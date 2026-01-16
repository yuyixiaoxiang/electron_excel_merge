/**
 * 主进程入口：负责创建 Electron 窗口、解析 git/Fork 传入的三方合并参数，
 * 并通过 IPC 向渲染进程提供 Excel 读写与三方 diff / merge 的能力。
 */
import { app, BrowserWindow, dialog, ipcMain } from 'electron';
import * as path from 'path';
import { spawn } from 'child_process';
import { Workbook, Worksheet, Row, Cell, CellValue } from 'exceljs';

// 保持对主窗口的引用，避免被 GC 回收导致窗口被意外关闭
let mainWindow: BrowserWindow | null = null;

const isDev = process.env.NODE_ENV === 'development';

/**
 * CLI three-way merge arguments for git/Fork integration.
 *
 * 约定（以 Fork / git mergetool 为例）：
 *   - diff 模式:   app.exe OURS THEIRS
 *   - merge 模式:  app.exe BASE OURS THEIRS [MERGED]
 *
 * 当带有 mergedPath 时，保存结果会直接写回 MERGED 文件；
 * 否则会回退到覆盖 ours（当前分支工作区文件）。
 */
interface CliThreeWayArgs {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  mergedPath?: string;
  mode: 'diff' | 'merge';
}

/**
 * 从 process.argv 中解析三方合并相关参数。
 *
 * - 开发环境下 argv 形如: [electron, main.js, '.', ...args]
 * - 打包后 exe 下 argv 形如: [app.exe, ...args]
 */
const parseCliThreeWayArgs = (): CliThreeWayArgs | null => {
  // 对于开发环境: process.argv = [electron, main.js, '.', ...args]
  // 对于打包后的 exe: process.argv = [app.exe, ...args]
  const argStartIndex = app.isPackaged ? 1 : 2;
  const rawArgs = process.argv.slice(argStartIndex);
  const userArgs = rawArgs.filter((arg) => !arg.startsWith('--'));

  // 2 个参数: 认为是 diff 模式 -> base 与 ours 相同（仅用于计算差异）
  if (userArgs.length === 2) {
    const [oursPath, theirsPath] = userArgs;
    return { basePath: oursPath, oursPath, theirsPath, mode: 'diff' };
  }

  if (userArgs.length < 3) {
    return null;
  }

  const [basePath, oursPath, theirsPath, mergedPath] = userArgs;
  return { basePath, oursPath, theirsPath, mergedPath, mode: 'merge' };
};

// 解析启动参数得到的三方合并信息（若无参数则为 null，走交互式模式）
const cliThreeWayArgs: CliThreeWayArgs | null = parseCliThreeWayArgs();

/**
 * 尝试在目标文件所在目录执行一次 `git add <filePath>`，
 * 方便在作为 merge tool 运行时自动标记冲突已解决。
 *
 * 注意：这里做的是“尽力而为”的操作，失败只会打印日志，不会中断主流程。
 */
const gitAddFile = (filePath: string): Promise<void> => {
  return new Promise((resolve) => {
    const cwd = path.dirname(filePath);
    const child = spawn('git', ['add', filePath], { cwd, stdio: 'ignore' });

    child.on('error', (err) => {
      console.error('git add failed', err);
      resolve();
    });

    child.on('close', (code) => {
      if (code !== 0) {
        console.error('git add exited with code', code);
      }
      resolve();
    });
  });
};

/**
 * 创建主浏览器窗口并加载前端页面。
 *
 * 开发模式下连接本地 webpack dev server，
 * 生产模式下加载打包到 dist 中的 index.html。
 */
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDev) {
    mainWindow.loadURL('http://localhost:3000');
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', '..', 'dist', 'index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// IPC types
interface SheetCell {
  address: string; // e.g. "A1"
  row: number;
  col: number;
  value: string | number | null;
}

interface SheetData {
  sheetName: string;
  rows: SheetCell[][];
}

interface MergeCell {
  address: string;
  row: number;
  col: number;
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  status: 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
  mergedValue: string | number | null;
}

interface MergeSheetData {
  sheetName: string;
  rows: MergeCell[][];
}

interface SaveMergeCellInput {
  address: string;
  value: string | number | null;
}

interface SaveMergeRequest {
  templatePath: string;
  cells: SaveMergeCellInput[];
}

interface SaveMergeResponse {
  success: boolean;
  filePath?: string;
  cancelled?: boolean;
  errorMessage?: string;
}

let currentFilePath: string | null = null;

/**
 * 处理渲染进程请求：选择并打开一个 Excel 文件。
 *
 * 返回：文件路径 + 所有工作表的二维单元格数据（仅包含“值”），
 * 用于单文件查看/编辑模式。
 */
ipcMain.handle('excel:open', async () => {
  if (!mainWindow) return null;

  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
    properties: ['openFile'],
  });

  if (canceled || filePaths.length === 0) {
    return null;
  }

  const filePath = filePaths[0];
  currentFilePath = filePath;

  const workbook = new Workbook();
  await workbook.xlsx.readFile(filePath);

  const buildSheetData = (worksheet: Worksheet): SheetData => {
    const rows: SheetCell[][] = [];

    const getSimpleValue = (raw: CellValue): string | number | null => {
      if (raw === null || raw === undefined) return null;
      // 富文本：raw.richText 是一个包含 { text } 的数组
      if (typeof raw === 'object' && Array.isArray((raw as any).richText)) {
        const parts = (raw as any).richText
          .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
          .join('');
        return parts;
      }
      if (typeof raw === 'object' && 'text' in raw) {
        return (raw as any).text ?? null;
      }
      if (typeof raw === 'object' && 'result' in raw) {
        return (raw as any).result ?? null;
      }
      if (typeof raw === 'string' || typeof raw === 'number') {
        return raw;
      }
      return String(raw);
    };

    worksheet.eachRow((row: Row, rowNumber: number) => {
      const rowCells: SheetCell[] = [];
      row.eachCell({ includeEmpty: true }, (cell: Cell, colNumber: number) => {
        const value = getSimpleValue(cell.value as any);

        rowCells.push({
          address: cell.address,
          row: rowNumber,
          col: colNumber,
          value,
        });
      });
      rows.push(rowCells);
    });

    return {
      sheetName: worksheet.name,
      rows,
    };
  };

  const sheets: SheetData[] = workbook.worksheets.map((ws) => buildSheetData(ws));

  return { filePath, sheet: sheets[0], sheets };
});

interface CellChange {
  address: string;
  newValue: string | number | null;
}

/**
 * 将单文件编辑模式下用户修改过的单元格写回原始 Excel 文件。
 *
 * 只修改单元格的 value，不动样式/公式等格式信息。
 */
ipcMain.handle('excel:saveChanges', async (_event, changes: CellChange[]) => {
  if (!currentFilePath) {
    throw new Error('No Excel file is currently loaded');
  }

  const workbook = new Workbook();
  await workbook.xlsx.readFile(currentFilePath);
  const worksheet = workbook.worksheets[0];

  for (const change of changes) {
    const cell = worksheet.getCell(change.address);
    cell.value = change.newValue as any; // only change value, keep formatting/styles
  }

  await workbook.xlsx.writeFile(currentFilePath);

  return { success: true };
});

// 保存三方 merge 结果到新的 Excel 文件，仅修改值，不改格式
//
// 在 git/Fork merge 模式下：
//   - 如果提供了 MERGED 参数，则结果写回 MERGED；
//   - 否则回退到覆盖 ours；
// 在 diff 模式下：
//   - 直接覆盖 ours（LOCAL）。
// 交互式模式下：
//   - 弹出保存对话框，由用户选择目标路径。
ipcMain.handle('excel:saveMergeResult', async (_event, req: SaveMergeRequest): Promise<SaveMergeResponse> => {
  if (!mainWindow) {
    throw new Error('Main window is not available');
  }

  try {
    const { templatePath, cells } = req as { templatePath: string; cells: { sheetName: string; address: string; value: string | number | null }[] };
    let targetPath: string | undefined;

    if (cliThreeWayArgs && cliThreeWayArgs.mode === 'merge') {
      // git / Fork merge 模式：优先写入 MERGED（工作区对应文件），如果命令只传了 base/ours/theirs 三个参数，则回退覆盖 ours。
      const oursPath = cliThreeWayArgs.oursPath;
      const mergedPath = cliThreeWayArgs.mergedPath;
      targetPath = mergedPath || oursPath;
    } else if (cliThreeWayArgs && cliThreeWayArgs.mode === 'diff') {
      targetPath = cliThreeWayArgs.oursPath;
    } else {
      const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
        title: '保存合并后的 Excel',
        defaultPath: templatePath,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
      });

      if (canceled || !filePath) {
        return { success: false, cancelled: true };
      }
      targetPath = filePath;
    }

    const workbook = new Workbook();
    await workbook.xlsx.readFile(templatePath);

    for (const cellInfo of cells) {
      const ws = workbook.getWorksheet(cellInfo.sheetName) ?? workbook.worksheets[0];
      const cell = ws.getCell(cellInfo.address);
      cell.value = cellInfo.value as any;
    }

    await workbook.xlsx.writeFile(targetPath);

    // 如果是通过 git/Fork 的 merge 模式启动，并且有明确的目标文件，尝试自动执行一次 git add
    if (cliThreeWayArgs && cliThreeWayArgs.mode === 'merge' && targetPath) {
      try {
        await gitAddFile(targetPath);
      } catch (e) {
        console.error('git add after merge failed', e);
      }
    }

    return { success: true, filePath: targetPath };
  } catch (err: any) {
    console.error('excel:saveMergeResult failed', err);
    return { success: false, errorMessage: err?.message ?? String(err) };
  }
});

// 三方 diff：base / ours / theirs，只比较单元格值，忽略格式
//
// 返回给渲染进程的数据是：
//   - base / ours / theirs 的文件路径；
//   - 每个工作表的三方单元格值 + 差异状态（unchanged / conflict 等）。
ipcMain.handle('excel:openThreeWay', async () => {
  if (!mainWindow) return null;

  // 对单个工作表做三方 diff：
  // - 将 base / ours / theirs 的同一坐标单元格抽取出来；
  // - 根据值的相等情况计算 status；
  // - 默认 mergedValue 从 base 或 ours 推导出一个初始值，之后可在前端调整。
  const buildMergeSheet = (baseWs: any, oursWs: any, theirsWs: any): MergeSheetData => {
    const rows: MergeCell[][] = [];

    const getSimpleValue = (v: any): string | number | null => {
      if (v === null || v === undefined) return null;
      // 富文本：v.richText 是一个包含 { text } 的数组
      if (typeof v === 'object' && Array.isArray((v as any).richText)) {
        const parts = (v as any).richText
          .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
          .join('');
        return parts;
      }
      if (typeof v === 'object' && 'text' in v) return (v as any).text ?? null;
      if (typeof v === 'object' && 'result' in v) return (v as any).result ?? null;
      if (typeof v === 'string' || typeof v === 'number') return v;
      return String(v);
    };

    const maxRow = Math.max(baseWs.rowCount, oursWs.rowCount, theirsWs.rowCount);
    const maxCol = Math.max(baseWs.columnCount, oursWs.columnCount, theirsWs.columnCount);

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      const mergeRow: MergeCell[] = [];

      for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
        const baseCell = baseWs.getRow(rowNumber).getCell(colNumber);
        const oursCell = oursWs.getRow(rowNumber).getCell(colNumber);
        const theirsCell = theirsWs.getRow(rowNumber).getCell(colNumber);

        const address =
          (oursCell && oursCell.address) ||
          (baseCell && baseCell.address) ||
          (theirsCell && theirsCell.address) ||
          `${colNumber}:${rowNumber}`;

        const baseValue = getSimpleValue(baseCell?.value);
        const oursValue = getSimpleValue(oursCell?.value);
        const theirsValue = getSimpleValue(theirsCell?.value);

        const same = (a: any, b: any) => a === b;

        let status: MergeCell['status'];
        let mergedValue: string | number | null = baseValue;

        const equalBO = same(baseValue, oursValue);
        const equalBT = same(baseValue, theirsValue);
        const equalOT = same(oursValue, theirsValue);

        if (equalBO && equalBT) {
          status = 'unchanged';
          mergedValue = baseValue;
        } else if (!equalBO && equalBT) {
          status = 'ours-changed';
          mergedValue = oursValue;
        } else if (equalBO && !equalBT) {
          status = 'theirs-changed';
          mergedValue = theirsValue;
        } else if (!equalBO && !equalBT && equalOT) {
          status = 'both-changed-same';
          mergedValue = oursValue;
        } else {
          status = 'conflict';
          mergedValue = oursValue; // 默认先用 ours，方便后续人工调整
        }

        mergeRow.push({
          address,
          row: rowNumber,
          col: colNumber,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
      }

      rows.push(mergeRow);
    }

    return {
      sheetName: baseWs.name,
      rows,
    };
  };

  // 按同名 sheet 做 diff
  /**
   * 按工作表名称对三个工作簿做对齐，仅对“同时存在于 base / ours / theirs 中的工作表”做 diff。
   */
  const buildMergeSheetsForWorkbooks = async (
    basePath: string,
    oursPath: string,
    theirsPath: string,
  ) => {
    const baseWb = new Workbook();
    const oursWb = new Workbook();
    const theirsWb = new Workbook();

    await baseWb.xlsx.readFile(basePath);
    await oursWb.xlsx.readFile(oursPath);
    await theirsWb.xlsx.readFile(theirsPath);

    const baseByName = new Map<string, any>();
    baseWb.worksheets.forEach((ws) => baseByName.set(ws.name, ws));
    const oursByName = new Map<string, any>();
    oursWb.worksheets.forEach((ws) => oursByName.set(ws.name, ws));
    const theirsByName = new Map<string, any>();
    theirsWb.worksheets.forEach((ws) => theirsByName.set(ws.name, ws));

    const commonNames = Array.from(baseByName.keys()).filter(
      (name) => oursByName.has(name) && theirsByName.has(name),
    );

    const mergeSheets: MergeSheetData[] = commonNames.map((name) =>
      buildMergeSheet(baseByName.get(name), oursByName.get(name), theirsByName.get(name)),
    );

    return { basePath, oursPath, theirsPath, mergeSheets };
  };

  // 如果从 git/Fork 传入了 base/ours/theirs，就直接使用这些路径
  if (cliThreeWayArgs) {
    const { basePath, oursPath, theirsPath } = cliThreeWayArgs;
    const { mergeSheets } = await buildMergeSheetsForWorkbooks(basePath, oursPath, theirsPath);

    return {
      basePath,
      oursPath,
      theirsPath,
      sheet: mergeSheets[0],
      sheets: mergeSheets,
    };
  }

  // 没有 CLI 参数时，回退到交互式选择文件的模式
  const pickFile = async (title: string) => {
    const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow!, {
      title,
      filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
      properties: ['openFile'],
    });
    if (canceled || filePaths.length === 0) return null;
    return filePaths[0];
  };

  const basePath = await pickFile('选择 base 版本 Excel');
  if (!basePath) return null;
  const oursPath = await pickFile('选择 ours (当前分支) Excel');
  if (!oursPath) return null;
  const theirsPath = await pickFile('选择 theirs (合并分支) Excel');
  if (!theirsPath) return null;

  const { mergeSheets } = await buildMergeSheetsForWorkbooks(basePath, oursPath, theirsPath);

  return { basePath, oursPath, theirsPath, sheet: mergeSheets[0], sheets: mergeSheets };
});

// 将 CLI three-way 信息暴露给渲染进程，便于自动加载
ipcMain.handle('excel:getCliThreeWayInfo', async () => {
  if (!cliThreeWayArgs) return null;
  return cliThreeWayArgs;
});
