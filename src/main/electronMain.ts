/**
 * 主进程入口：负责创建 Electron 窗口、解析git/Fork 传入的三方合并参数，
 * 并通过 IPC 向渲染进程提供Excel 读写与三方diff / merge 的能力。
 */
import { app, BrowserWindow, dialog, ipcMain, Menu, MenuItemConstructorOptions } from 'electron';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { spawn } from 'child_process';
import { Workbook, Worksheet, Row, Cell, CellValue } from 'exceljs';
import {
  ThreeWayCompareMode,
  classifyThreeWayCell,
  diffArraysToOps,
  getThreeWayRuntimeConfig,
  normalizeThreeWayCompareMode,
} from './threeWayDiffCore';

// 保持对主窗口的引用，避免被GC 回收导致窗口被意外关闭
let mainWindow: BrowserWindow | null = null;

const isDev = process.env.NODE_ENV === 'development';
const shouldBootstrapElectron = process.env.ELECTRON_EXCEL_MERGE_DISABLE_BOOTSTRAP !== '1';
const DEFAULT_FROZEN_HEADER_ROWS = 3;
const DEFAULT_ROW_SIMILARITY_THRESHOLD = 0.9;
const DEBUG_LOG_FILE_NAME = 'eMerge-debug.log';

interface DebugLogEntry {
  source: string;
  event: string;
  details?: unknown;
}

type ExcelJsWorksheetXformCompat = {
  prototype: {
    __eMergeCompatPatched?: boolean;
    reconcile: (model: any, options: any) => unknown;
  };
};
type ExcelJsStylesXformCompat = {
  prototype: {
    __eMergeCompatPatched?: boolean;
    getStyleModel: (id: number) => unknown;
    getDxfStyle?: (id: number) => unknown;
  };
};

const getDebugLogPath = (): string => path.join(app.getPath('temp'), DEBUG_LOG_FILE_NAME);
const debugFileLabel = (filePath: string | null | undefined): string | null =>
  filePath ? path.basename(filePath) : null;
const safeSerializeDebug = (value: unknown): string => {
  try {
    const seen = new WeakSet<object>();
    return JSON.stringify(value, (_key, current) => {
      if (current instanceof Error) {
        return {
          name: current.name,
          message: current.message,
          stack: current.stack,
        };
      }
      if (typeof current === 'string' && current.length > 500) {
        return `${current.slice(0, 500)}…(${current.length})`;
      }
      if (typeof current === 'object' && current !== null) {
        if (seen.has(current)) return '[Circular]';
        seen.add(current);
      }
      return current;
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return JSON.stringify({ serializationError: message, fallback: String(value) });
  }
};
const appendDebugLog = (source: string, event: string, details?: unknown) => {
  try {
    const line = `${new Date().toISOString()} [pid:${process.pid}] [${source}] ${event}${
      details === undefined ? '' : ` ${safeSerializeDebug(details)}`
    }\n`;
    fs.appendFile(getDebugLogPath(), line, () => {});
  } catch (error) {
    console.error('debug log write failed', error);
  }
};

process.on('uncaughtException', (error) => {
  appendDebugLog('main', 'uncaughtException', error);
  console.error(error);
});
process.on('unhandledRejection', (reason) => {
  appendDebugLog('main', 'unhandledRejection', reason);
  console.error(reason);
});

const patchExcelJsLegacyCommentCompat = () => {
  const worksheetXform = require('exceljs/lib/xlsx/xform/sheet/worksheet-xform') as ExcelJsWorksheetXformCompat;
  const stylesXform = require('exceljs/lib/xlsx/xform/style/styles-xform') as ExcelJsStylesXformCompat;
  let installedAnyPatch = false;
  if (!worksheetXform?.prototype?.__eMergeCompatPatched) {
    const originalReconcile = worksheetXform?.prototype?.reconcile;
    if (typeof originalReconcile !== 'function') {
      appendDebugLog('main', 'exceljs:legacy-comment-compat-skipped', { reason: 'missing-reconcile' });
      return;
    }
    worksheetXform.prototype.reconcile = function patchedWorksheetReconcile(model: any, options: any) {
      const nextOptions = options ?? {};
      const relationships = Array.isArray(model?.relationships) ? model.relationships : [];
      const missingCommentTargets: string[] = [];
      const missingVmlTargets: string[] = [];
      if (!nextOptions.comments) nextOptions.comments = {};
      if (!nextOptions.vmlDrawings) nextOptions.vmlDrawings = {};
      relationships.forEach((rel: any) => {
        const relType = String(rel?.Type ?? '');
        const target = typeof rel?.Target === 'string' ? rel.Target : '';
        if (!target) return;
        if (relType.endsWith('/comments') && !nextOptions.comments[target]) {
          nextOptions.comments[target] = { comments: [] };
          missingCommentTargets.push(target);
        }
        if (relType.endsWith('/vmlDrawing') && !nextOptions.vmlDrawings[target]) {
          nextOptions.vmlDrawings[target] = { comments: [] };
          missingVmlTargets.push(target);
        }
      });
      if (missingCommentTargets.length || missingVmlTargets.length) {
        appendDebugLog('main', 'exceljs:legacy-comment-compat-applied', {
          sheetName: model?.name ?? null,
          missingCommentTargets,
          missingVmlTargets,
        });
      }
      return originalReconcile.call(this, model, nextOptions);
    };
    worksheetXform.prototype.__eMergeCompatPatched = true;
    installedAnyPatch = true;
  }
  if (!stylesXform?.prototype?.__eMergeCompatPatched) {
    const originalGetStyleModel = stylesXform?.prototype?.getStyleModel;
    if (typeof originalGetStyleModel === 'function') {
      stylesXform.prototype.getStyleModel = function patchedGetStyleModel(this: any, id: number) {
        if (!this || !this.model || !Array.isArray(this.model.styles)) {
          return {};
        }
        return originalGetStyleModel.call(this, id) ?? {};
      };
    }
    const originalGetDxfStyle = stylesXform?.prototype?.getDxfStyle;
    if (typeof originalGetDxfStyle === 'function') {
      stylesXform.prototype.getDxfStyle = function patchedGetDxfStyle(this: any, id: number) {
        if (!this || !this.model || !Array.isArray(this.model.dxfs)) {
          return {};
        }
        return originalGetDxfStyle.call(this, id) ?? {};
      };
    }
    stylesXform.prototype.__eMergeCompatPatched = true;
    installedAnyPatch = true;
  }
  if (installedAnyPatch) {
    appendDebugLog('main', 'exceljs:legacy-comment-compat-installed');
  }
};

/**
 * CLI three-way merge arguments for git/Fork integration.
 *
 * 约定（以 Fork / git mergetool 为例）：
 *   - diff 模式:    app.exe OURS THEIRS
 *   - simple merge: app.exe OURS THEIRS MERGED
 *   - merge 模式:   app.exe BASE OURS THEIRS [MERGED]
 *
 * 当带有mergedPath 时，保存结果会直接写回MERGED 文件：
 * 否则会回退到覆盖ours（当前分支工作区文件）。
 */
interface CliThreeWayArgs {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  mergedPath?: string;
  mergedPathRaw?: string;
  mode: 'diff' | 'merge' | 'simple-merge';
}

const isCliMergeMode = (mode: CliThreeWayArgs['mode'] | null | undefined): boolean =>
  mode === 'merge' || mode === 'simple-merge';

const stripOuterQuotes = (value: string): string => {
  const trimmed = value.trim();
  if (trimmed.length < 2) return trimmed;
  const quotePairs: Array<readonly [string, string]> = [
    ['"', '"'],
    ["'", "'"],
    ['“', '”'],
    ['‘', '’'],
  ];
  for (const [start, end] of quotePairs) {
    if (trimmed.startsWith(start) && trimmed.endsWith(end)) {
      return trimmed.slice(start.length, trimmed.length - end.length).trim();
    }
  }
  return trimmed;
};
const normalizeCliPathText = (value: string): string => {
  const raw = stripOuterQuotes(value).trim();
  if (!raw) return raw;
  const msysDrivePathMatch = /^\/([a-zA-Z])\/(.*)$/.exec(raw);
  if (msysDrivePathMatch) {
    return `${msysDrivePathMatch[1].toUpperCase()}:\\${msysDrivePathMatch[2].replace(/\//g, '\\')}`;
  }
  return raw;
};
const isMsysTempPathText = (value: string): boolean => /^\/tmp(?:\/|$)/i.test(stripOuterQuotes(value).trim());
const resolveMsysTempPath = (value: string): string | null => {
  const raw = stripOuterQuotes(value).trim();
  if (!isMsysTempPathText(raw)) return null;
  const tempRoot = path.normalize(os.tmpdir());
  const relative = raw.replace(/^\/tmp\/?/i, '');
  if (!relative) return tempRoot;
  return path.normalize(path.join(tempRoot, ...relative.split('/').filter(Boolean)));
};
const getCliRelativePathRoots = (): string[] => {
  const roots: string[] = [];
  const pushRoot = (value: string | null | undefined) => {
    if (!value) return;
    const normalized = normalizeCliPathText(value);
    if (!normalized) return;
    const absolute = path.isAbsolute(normalized) ? path.normalize(normalized) : path.resolve(normalized);
    if (!roots.includes(absolute)) {
      roots.push(absolute);
    }
  };
  pushRoot(process.env.GIT_WORK_TREE);
  pushRoot(process.env.PWD);
  pushRoot(process.env.INIT_CWD);
  pushRoot(process.env.OLDPWD);
  pushRoot(process.cwd());
  return roots;
};
const resolveCliInputPath = (value: string): string => {
  const normalized = normalizeCliPathText(value);
  if (!normalized) return normalized;
  const msysTempPath = resolveMsysTempPath(normalized);
  if (msysTempPath) return msysTempPath;
  if (path.isAbsolute(normalized)) return path.normalize(normalized);
  for (const root of getCliRelativePathRoots()) {
    const candidate = path.resolve(root, normalized);
    if (fs.existsSync(candidate)) {
      return path.normalize(candidate);
    }
  }
  return path.resolve(process.cwd(), normalized);
};
const resolveCliOutputPath = (value: string | null | undefined): string | undefined => {
  if (!value) return undefined;
  const normalized = normalizeCliPathText(value);
  if (!normalized) return undefined;
  const msysTempPath = resolveMsysTempPath(normalized);
  if (msysTempPath) return msysTempPath;
  if (path.isAbsolute(normalized)) return path.normalize(normalized);
  for (const root of getCliRelativePathRoots()) {
    const candidate = path.resolve(root, normalized);
    if (fs.existsSync(candidate) || fs.existsSync(path.dirname(candidate))) {
      return path.normalize(candidate);
    }
  }
  return undefined;
};

/**
 * 从process.argv 中解析三方合并相关参数。
 *
 * - 开发环境下 argv 形如: [electron, main.js, '.', ...args]
 * - 打包后exe 与argv 形如: [app.exe, ...args]
 */
const parseCliThreeWayArgs = (): CliThreeWayArgs | null => {
  // 对于开发环境 process.argv = [electron, main.js, '.', ...args]
  // 对于打包后的 exe: process.argv = [app.exe, ...args]
  const argStartIndex = app?.isPackaged ? 1 : 2;
  const rawArgs = process.argv.slice(argStartIndex);
  const userArgs = rawArgs
    .map((arg) => stripOuterQuotes(arg))
    .filter((arg) => !!arg && !arg.startsWith('--'));
  // 兼容开发模式下 `electron .` 带来的app path 参数
  if (userArgs.length >= 3) {
    const first = userArgs[0];
    const appPath = app.getAppPath ? app.getAppPath() : '';
    const firstResolved = path.resolve(first);
    const appResolved = appPath ? path.resolve(appPath) : '';
    let isDir = false;
    try {
      isDir = fs.statSync(firstResolved).isDirectory();
    } catch {
      isDir = false;
    }
    if (first === '.' || (!!appResolved && firstResolved === appResolved) || isDir) {
      userArgs.shift();
    }
  }

  // 2 个参数，认为是 diff 模式 -> base 与 ours 相同（仅用于计算差异）
  if (userArgs.length === 2) {
    const [oursPath, theirsPath] = userArgs.map(resolveCliInputPath);
    return { basePath: oursPath, oursPath, theirsPath, mode: 'diff' };
  }

  if (userArgs.length === 3) {
    const [oursArg, theirsArg, mergedArg] = userArgs.map(normalizeCliPathText);
    const oursPath = resolveCliInputPath(oursArg);
    const theirsPath = resolveCliInputPath(theirsArg);
    const mergedPathRaw = mergedArg || undefined;
    const mergedPath = resolveCliOutputPath(mergedPathRaw);
    return { basePath: oursPath, oursPath, theirsPath, mergedPath, mergedPathRaw, mode: 'simple-merge' };
  }

  if (userArgs.length < 4) {
    return null;
  }

  const [baseArg, oursArg, theirsArg, mergedArg] = userArgs.map(normalizeCliPathText);
  const basePath = resolveCliInputPath(baseArg);
  const oursPath = resolveCliInputPath(oursArg);
  const theirsPath = resolveCliInputPath(theirsArg);
  const mergedPathRaw = mergedArg || undefined;
  const mergedPath = resolveCliOutputPath(mergedPathRaw);
  return { basePath, oursPath, theirsPath, mergedPath, mergedPathRaw, mode: 'merge' };
};

// 解析启动参数得到的三方合并信息（若无参数则为 null，走交互式模式）
const cliThreeWayArgs: CliThreeWayArgs | null = parseCliThreeWayArgs();
if (cliThreeWayArgs) {
  appendDebugLog('main', 'cli:parsed', {
    mode: cliThreeWayArgs.mode,
    cwd: process.cwd(),
    pwd: process.env.PWD ?? null,
    initCwd: process.env.INIT_CWD ?? null,
    gitWorkTree: process.env.GIT_WORK_TREE ?? null,
    basePath: debugFileLabel(cliThreeWayArgs.basePath),
    oursPath: debugFileLabel(cliThreeWayArgs.oursPath),
    theirsPath: debugFileLabel(cliThreeWayArgs.theirsPath),
    mergedPathRaw: cliThreeWayArgs.mergedPathRaw ?? null,
    mergedPath: cliThreeWayArgs.mergedPath ?? null,
  });
}
const getBundledGitInfo = (): { gitPath: string; env: NodeJS.ProcessEnv } | null => {
  const basePath = app?.isPackaged
    ? path.join(process.resourcesPath, 'git')
    : path.join(app.getAppPath(), 'resources', 'portable-git');
  const gitPath = path.join(basePath, 'cmd', 'git.exe');
  if (!fs.existsSync(gitPath)) return null;

  const env = { ...process.env };
  const extraPaths = [
    path.join(basePath, 'cmd'),
    path.join(basePath, 'mingw64', 'bin'),
    path.join(basePath, 'usr', 'bin'),
  ];
  const currentPath = env.PATH || env.Path || '';
  const newPath = [...extraPaths, currentPath].filter(Boolean).join(path.delimiter);
  env.PATH = newPath;
  env.Path = newPath;
  return { gitPath, env };
};

/**
 * 尝试在目标文件所在目录执行一次`git add <filePath>`：
 * 方便在作为merge tool 运行时自动标记冲突已解决。
 *
 * 注意：这里做的是“尽力而为”的操作，失败只会打印日志，不会中断主流程。
 */
const gitAddFile = (filePath: string): Promise<void> => {
  return new Promise((resolve) => {
    const cwd = path.dirname(filePath);
    const gitInfo = getBundledGitInfo();
    const gitCommand = gitInfo?.gitPath ?? 'git';
    const child = spawn(gitCommand, ['add', filePath], { cwd, stdio: 'ignore', env: gitInfo?.env });

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
 * 开发模式下连接本地 webpack dev server：
 * 生产模式下加载打包到 dist 中的 index.html。
 */
function createWindow() {
  appendDebugLog('main', 'createWindow:start', { isDev });
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
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', '..', 'dist', 'index.html'));
  }

  mainWindow.on('closed', () => {
    appendDebugLog('main', 'createWindow:closed');
    mainWindow = null;
  });
}

if (shouldBootstrapElectron) {
  app.whenReady().then(() => {
    appendDebugLog('main', 'app:ready', {
      isDev,
      isPackaged: app.isPackaged,
      cliMode: cliThreeWayArgs?.mode ?? null,
      logPath: getDebugLogPath(),
    });
    configureApplicationMenu();
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
}
type SimpleCellValue = string | number | null;

interface RowRecord {
  rowNumber: number; // 1-based Excel row number
  index: number; // 0-based index in extracted rows list
  values: SimpleCellValue[];
  nonEmptyCols: number[]; // 1-based column indices with non-empty values
  key?: string | null;
}
interface ColumnTypeSignature {
  num: number;
  str: number;
  empty: number;
  other: number;
}

interface ColumnRecord {
  colNumber: number; // 1-based Excel column number
  headerText: string; // normalized header text (joined by "|")
  headerKey: string; // stronger normalized key for matching
  typeSig: ColumnTypeSignature;
  sampleValues: string[]; // normalized sample values
}

interface AlignedColumn {
  baseCol?: number | null;
  oursCol?: number | null;
  theirsCol?: number | null;
}

interface AlignedRow {
  base?: RowRecord | null;
  ours?: RowRecord | null;
  theirs?: RowRecord | null;
  key?: string | null;
  ambiguousOurs?: boolean;
  ambiguousTheirs?: boolean;
}

/**
 * 将ExcelJS 的复杂单元格值转换为简单值（string | number | null）。
 * 
 * ExcelJS 的单元格值可能是：
 * - 简单类型：string、number
 * - 富文本：{ richText: [{text: '...'}] }
 * - 公式：{ formula: '...', result: value }
 * - 超链接等其他对象类型
 * 
 * 该函数统一提取其中的实际文本数值内容，忽略格式信息。
 */
const getSimpleValueForMerge = (v: any): SimpleCellValue => {
  if (v === null || v === undefined) return null;
  // 处理日期对象：转为 ISO 字符串，保持与 excel:open 中 getSimpleValue 一致
  if (v instanceof Date) return v.toISOString();
  // 处理富文本：拼接所有文本片段
  if (typeof v === 'object' && Array.isArray((v as any).richText)) {
    const parts = (v as any).richText
      .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
      .join('');
    return parts;
  }
  // 处理超链接等包含 text 属性的对象
  if (typeof v === 'object' && 'text' in v) return (v as any).text ?? null;
  // 处理公式单元格：取计算结果
  if (typeof v === 'object' && 'result' in v) return (v as any).result ?? null;
  // 简单类型直接返回
  if (typeof v === 'string' || typeof v === 'number') return v;
  // 其他类型转字符串
  return String(v);
};

const getFirstNonEmptyCellText = (ws: Worksheet | null | undefined, rowNumber: number, scanCols = 8): string | null => {
  if (!ws || rowNumber < 1) return null;
  const row = ws.getRow(rowNumber);
  const upperBound = Math.max(scanCols, ws.actualColumnCount ?? ws.columnCount ?? 0);
  for (let col = 1; col <= upperBound; col += 1) {
    const raw = getSimpleValueForMerge(row.getCell(col)?.value);
    if (raw == null) continue;
    const text = String(raw).trim();
    if (text) return text;
  }
  return null;
};

const resolveStructuredHeaderRowCount = (
  worksheets: Array<Worksheet | null | undefined>,
  requestedHeaderRowCount: number,
): number => {
  const requested = Math.max(0, Math.floor(requestedHeaderRowCount));
  let detected = 0;
  const scanLimit = Math.max(requested + 4, 8);
  worksheets.forEach((ws) => {
    if (!ws) return;
    let sheetDetected = 0;
    for (let rowNumber = 1; rowNumber <= scanLimit; rowNumber += 1) {
      const marker = getFirstNonEmptyCellText(ws, rowNumber);
      if (!marker || !marker.startsWith('##')) break;
      sheetDetected = rowNumber;
    }
    detected = Math.max(detected, sheetDetected);
  });
  return Math.max(requested, detected);
};

/**
 * 将单元格值标准化为字符串，用于比较和显示。
 * - null/undefined →空字符串
 * - 字符串→去除首尾空格
 * - 数字 →转字符串
 */
const normalizeCellValue = (v: SimpleCellValue): string => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  return String(v);
};

/**
 * 标准化主键列的值，用于行对齐。
 * 空字符串视为 null（即无主键），方便后续判断。
 */
const normalizeKeyValue = (v: SimpleCellValue): string | null => {
  const s = normalizeCellValue(v);
  return s === '' ? null : s;
};

/**
 * 标准化表头文本，用于列匹配。
 * 转为小写以忽略大小写差异。
 */
const normalizeHeaderText = (v: SimpleCellValue): string => {
  const s = normalizeCellValue(v);
  if (!s) return '';
  return s.toLowerCase();
};
/**
 * 生成更强的表头匹配键，用于精确匹配列。
 * - 转小写
 * - 去除所有空白
 * - 只保留字母、数字、中文字符
 * 
 * 例如：Icon名称, Asset..." →"icon名称asset"
 * 这样即使格式略有不同，也能匹配上相同语义的列。
 */
const normalizeHeaderKey = (text: string): string => {
  if (!text) return '';
  return text
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^0-9a-z\u4e00-\u9fa5]/gi, '');
};

/**
 * 为工作表的每一列提取特征信息，用于列对齐算法。
 * 
 * @param ws ExcelJS 工作表对象
 * @param headerCount 表头行数（前N行视为表头）
 * @param sampleRows 采样行数（用于类型和样本值统计）
 * @returns 列特征记录数组
 * 
 * 特征包括：
 * 1. headerText: 表头文本（多行用 | 分隔）
 * 2. headerKey: 标准化的表头键（用于精确匹配）
 * 3. typeSig: 数据类型签名（num/str/empty/other 的分布）
 * 4. sampleValues: 样本值集合（用于内容相似度比较）
 * 
 * 注意：完全空的列（表头和数据都为空）会被跳过，不生成记录。
 */
const buildColumnRecords = (
  ws: any,
  headerCount: number,
  sampleRows: number,
): ColumnRecord[] => {
  if (!ws) return [];
  // 获取工作表实际列数
  const actualColCount = Math.max(ws?.actualColumnCount ?? 0, ws?.columnCount ?? 0);
  const maxRow = Math.max(ws?.actualRowCount ?? 0, ws?.rowCount ?? 0, headerCount);
  const records: ColumnRecord[] = [];
  
  // 遍历每一列
  for (let col = 1; col <= actualColCount; col += 1) {
    // 1. 提取表头文本（拼接前 headerCount 行）
    const headerParts: string[] = [];
    for (let r = 1; r <= headerCount; r += 1) {
      const row = ws.getRow(r);
      const raw = getSimpleValueForMerge(row.getCell(col)?.value);
      const text = normalizeHeaderText(raw);
      if (text) headerParts.push(text);
    }
  const headerText = headerParts.join('|');
  const headerKey = normalizeHeaderKey(headerText);
    const typeSig: ColumnTypeSignature = { num: 0, str: 0, empty: 0, other: 0 };
    const sampleSet = new Set<string>();
    let sampled = 0;
    for (let r = headerCount + 1; r <= maxRow && sampled < sampleRows; r += 1) {
      const row = ws.getRow(r);
      const raw = getSimpleValueForMerge(row.getCell(col)?.value);
      const norm = normalizeCellValue(raw);
      if (norm === '') {
        typeSig.empty += 1;
        sampled += 1;
        continue;
      }
      if (typeof raw === 'number') typeSig.num += 1;
      else if (typeof raw === 'string') typeSig.str += 1;
      else typeSig.other += 1;
      sampleSet.add(norm);
      sampled += 1;
    }
    const sampleValues = Array.from(sampleSet).slice(0, 12);
    const hasDataSample = sampleValues.length > 0 || typeSig.num > 0 || typeSig.str > 0 || typeSig.other > 0;
    const isFullyEmpty = !headerText && !hasDataSample;
    if (isFullyEmpty) continue;

    records.push({
      colNumber: col,
      headerText,
      headerKey,
      typeSig,
      sampleValues,
    });
  }
  return records;
};

/**
 * 计算两个字符串的相似度（使用 Levenshtein 距离）。
 * 
 * @returns 0-1 之间的相似度，1 表示完全相同，0 表示完全不同。
 * 
 * 算法：Levenshtein 距离算法（动态规划）
 * - 计算将字符串 a 转换为b 所需的最小编辑步骤（插入、删除、替换）
 * - 相似度 = 1 - (距离 / 较长字符串长度）
 */
const stringSimilarity = (a: string, b: string): number => {
  if (!a && !b) return 1;
  if (!a || !b) return 0;
  const s = a.toLowerCase();
  const t = b.toLowerCase();
  if (s === t) return 1;
  const n = s.length;
  const m = t.length;
  if (n === 0 || m === 0) return 0;
  // 动态规划计算编辑距离
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  // 初始化：第i个字符转换为空需要i步
  for (let i = 0; i <= n; i += 1) dp[i][0] = i;
  for (let j = 0; j <= m; j += 1) dp[0][j] = j;
  // 填表：计算每个子问题的最小编辑距离
  for (let i = 1; i <= n; i += 1) {
    for (let j = 1; j <= m; j += 1) {
      const cost = s[i - 1] === t[j - 1] ? 0 : 1;  // 字符相同无需替换
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,       // 删除
        dp[i][j - 1] + 1,       // 插入
        dp[i - 1][j - 1] + cost, // 替换
      );
    }
  }
  const dist = dp[n][m];
  // 归一化为 0-1 之间的相似度
  return 1 - dist / Math.max(n, m);
};

/**
 * 计算两个列的数据类型签名相似度。
 * 
 * 类型签名 = { num, str, empty, other } 的分布比例。
 * 相似度= 1 - (比例差异的总和 / 2)。
 * 
 * 例如：
 * - A列：80% 数字，20% 字符串
 * - B列：85% 数字，15% 字符串
 * - 相似度很高，很可能是同一列
 */
const typeSignatureSimilarity = (a: ColumnTypeSignature, b: ColumnTypeSignature): number => {
  const totalA = a.num + a.str + a.empty + a.other;
  const totalB = b.num + b.str + b.empty + b.other;
  if (totalA === 0 && totalB === 0) return 1;
  if (totalA === 0 || totalB === 0) return 0;
  const pa = {
    num: a.num / totalA,
    str: a.str / totalA,
    empty: a.empty / totalA,
    other: a.other / totalA,
  };
  const pb = {
    num: b.num / totalB,
    str: b.str / totalB,
    empty: b.empty / totalB,
    other: b.other / totalB,
  };
  const dist =
    Math.abs(pa.num - pb.num) +
    Math.abs(pa.str - pb.str) +
    Math.abs(pa.empty - pb.empty) +
    Math.abs(pa.other - pb.other);
  return 1 - dist / 2;
};

const valueSimilarity = (a: string[], b: string[]): number => {
  if (a.length === 0 && b.length === 0) return 1;
  if (a.length === 0 || b.length === 0) return 0;
  const setA = new Set(a);
  const setB = new Set(b);
  let intersect = 0;
  setA.forEach((v) => {
    if (setB.has(v)) intersect += 1;
  });
  const union = setA.size + setB.size - intersect;
  if (union === 0) return 0;
  return intersect / union;
};

const columnSimilarity = (a: ColumnRecord, b: ColumnRecord): number => {
  const headerSim = stringSimilarity(a.headerKey || a.headerText, b.headerKey || b.headerText);
  const typeSim = typeSignatureSimilarity(a.typeSig, b.typeSig);
  const valSim = valueSimilarity(a.sampleValues, b.sampleValues);
  const hasHeader = (a.headerKey || a.headerText) && (b.headerKey || b.headerText);
  const wHeader = hasHeader ? 0.6 : 0.2;
  const wType = 0.2;
  const wVal = 0.2;
  const sum = wHeader + wType + wVal;
  return (wHeader * headerSim + wType * typeSim + wVal * valSim) / sum;
};

const alignColumnsBySimilarity = (
  baseCols: ColumnRecord[],
  sideCols: ColumnRecord[],
): { matched: Map<number, number>; gaps: Map<number, ColumnRecord[]> } => {
  const baseTokens = baseCols.map((c, i) => (c.headerKey || c.headerText ? (c.headerKey || c.headerText) : `__EMPTY_${i}`));
  const sideTokens = sideCols.map((c, i) => (c.headerKey || c.headerText ? (c.headerKey || c.headerText) : `__EMPTY_${i}`));
  const anchorPairs = lcsMatchPairs(baseTokens, sideTokens);
  const matched = new Map<number, number>();
  const usedSide = new Set<number>();
  for (const p of anchorPairs) {
    matched.set(p.aIndex, p.bIndex);
    usedSide.add(p.bIndex);
  }

  anchorPairs.sort((a, b) => a.aIndex - b.aIndex);

  const threshold = 0.55;
  const headerThreshold = 0.8;
  const matchSegment = (baseIdxs: number[], sideIdxs: number[]) => {
    if (baseIdxs.length === 0 || sideIdxs.length === 0) return;
    const pairs: Array<{ b: number; s: number; score: number }> = [];
    for (const b of baseIdxs) {
      for (const s of sideIdxs) {
        const headerA = baseCols[b].headerKey || baseCols[b].headerText;
        const headerB = sideCols[s].headerKey || sideCols[s].headerText;
        const headerSim = stringSimilarity(headerA, headerB);
        if (headerA && headerB && headerSim < headerThreshold) continue;
        const score = columnSimilarity(baseCols[b], sideCols[s]);
        if (score >= threshold) pairs.push({ b, s, score });
      }
    }
    pairs.sort((a, b) => b.score - a.score);
    for (const p of pairs) {
      if (matched.has(p.b)) continue;
      if (usedSide.has(p.s)) continue;
      matched.set(p.b, p.s);
      usedSide.add(p.s);
    }
  };

  let prevBase = -1;
  let prevSide = -1;
  for (const anchor of anchorPairs) {
    const baseIdxs: number[] = [];
    const sideIdxs: number[] = [];
    for (let b = prevBase + 1; b < anchor.aIndex; b += 1) baseIdxs.push(b);
    for (let s = prevSide + 1; s < anchor.bIndex; s += 1) sideIdxs.push(s);
    matchSegment(baseIdxs, sideIdxs);
    prevBase = anchor.aIndex;
    prevSide = anchor.bIndex;
  }
  if (prevBase < baseCols.length - 1 || prevSide < sideCols.length - 1) {
    const baseIdxs: number[] = [];
    const sideIdxs: number[] = [];
    for (let b = prevBase + 1; b < baseCols.length; b += 1) baseIdxs.push(b);
    for (let s = prevSide + 1; s < sideCols.length; s += 1) sideIdxs.push(s);
    matchSegment(baseIdxs, sideIdxs);
  }

  const gaps = new Map<number, ColumnRecord[]>();
  const matchedPairsBySide = Array.from(matched.entries())
    .map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }))
    .sort((a, b) => a.sideIndex - b.sideIndex);
  for (let s = 0; s < sideCols.length; s += 1) {
    if (usedSide.has(s)) continue;
    let gap = -1;
    for (const p of matchedPairsBySide) {
      if (p.sideIndex < s) gap = p.baseIndex;
      if (p.sideIndex >= s) break;
    }
    if (!gaps.has(gap)) gaps.set(gap, []);
    gaps.get(gap)!.push(sideCols[s]);
  }

  return { matched, gaps };
};

const buildAlignedColumns = (
  baseWs: any,
  oursWs: any,
  theirsWs: any,
  headerCount: number,
): AlignedColumn[] => {
  const sampleRows = 20;
  const baseCols = buildColumnRecords(baseWs, headerCount, sampleRows);
  const oursCols = buildColumnRecords(oursWs, headerCount, sampleRows);
  const theirsCols = buildColumnRecords(theirsWs, headerCount, sampleRows);

  const alignBase = baseCols.length > 0 ? baseCols : oursCols.length > 0 ? oursCols : theirsCols;
  const baseRefCols = alignBase;
  const oursAlign = alignColumnsBySimilarity(baseRefCols, oursCols);
  const theirsAlign = alignColumnsBySimilarity(baseRefCols, theirsCols);

  const aligned: AlignedColumn[] = [];
  const addGapCols = (gapIndex: number) => {
    const oursGap = oursAlign.gaps.get(gapIndex) ?? [];
    const theirsGap = theirsAlign.gaps.get(gapIndex) ?? [];
    for (const c of oursGap) aligned.push({ oursCol: c.colNumber ?? null });
    for (const c of theirsGap) aligned.push({ theirsCol: c.colNumber ?? null });
  };

  addGapCols(-1);
  for (let i = 0; i < baseRefCols.length; i += 1) {
    const baseColNumber = baseRefCols[i]?.colNumber ?? null;
    const oursIndex = oursAlign.matched.get(i);
    const theirsIndex = theirsAlign.matched.get(i);
    aligned.push({
      baseCol: baseColNumber,
      oursCol: typeof oursIndex === 'number' ? oursCols[oursIndex]?.colNumber ?? null : null,
      theirsCol: typeof theirsIndex === 'number' ? theirsCols[theirsIndex]?.colNumber ?? null : null,
    });
    addGapCols(i);
  }

  if (baseRefCols.length === 0) {
    // base/ours 모두为空时，直接按theirs 追加
    for (const c of theirsCols) {
      aligned.push({ theirsCol: c.colNumber ?? null });
    }
  }

  return aligned;
};

const colNumberToLabel = (colNumber: number): string => {
  let n = Math.max(1, Math.floor(colNumber));
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode('A'.charCodeAt(0) + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
};

/**
 * 从工作表中提取行记录（未对齐版本，用于单文件或列对齐前）。
 * 
 * @param ws ExcelJS 工作表对象
 * @param colCount 列数
 * @param primaryKeyCol 主键列号（1-based，-1 表示无主键）
 * @returns 行记录数组，每条记录包含：
 *   - rowNumber: Excel 中的原始行号
 *   - index: 在提取列表中的索引
 *   - values: 所有列的值数组
 *   - nonEmptyCols: 非空列的列号列表
 *   - key: 主键值（如果有）
 * 
 * 注意：完全空的行会被跳过。
 */
const buildRowRecords = (ws: any, colCount: number, primaryKeyCol: number): RowRecord[] => {
  const rows: RowRecord[] = [];
  let index = 0;
  // 遍历所有非空行
  ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    const values: SimpleCellValue[] = [];
    const nonEmptyCols: number[] = [];
    // 读取每一列的值
    for (let col = 1; col <= colCount; col += 1) {
      const cell = row.getCell(col);
      const value = getSimpleValueForMerge(cell?.value);
      values.push(value);
      if (value !== null && value !== '') {
        nonEmptyCols.push(col);
      }
    }
    // 跳过完全空的行
    if (nonEmptyCols.length === 0) return;
    // 提取主键值（如果有指定主键列）
    const key =
      primaryKeyCol >= 1 && primaryKeyCol <= colCount
        ? normalizeKeyValue(values[primaryKeyCol - 1])
        : null;
    rows.push({ rowNumber, index, values, nonEmptyCols, key });
    index += 1;
  });
  return rows;
};

const buildHeaderRowRecord = (ws: any, rowNumber: number, colCount: number, primaryKeyCol: number): RowRecord => {
  const values: SimpleCellValue[] = [];
  const nonEmptyCols: number[] = [];
  const row = ws.getRow(rowNumber);
  for (let col = 1; col <= colCount; col += 1) {
    const cell = row.getCell(col);
    const value = getSimpleValueForMerge(cell?.value);
    values.push(value);
    if (value !== null && value !== '') {
      nonEmptyCols.push(col);
    }
  }
  const key =
    primaryKeyCol >= 1 && primaryKeyCol <= colCount
      ? normalizeKeyValue(values[primaryKeyCol - 1])
      : null;
  return {
    rowNumber,
    index: rowNumber - 1,
    values,
    nonEmptyCols,
    key,
  };
};

const getWorksheetRowUpperBound = (ws: any): number => Math.max(ws?.rowCount ?? 0, ws?.actualRowCount ?? 0);

/**
 * 从工作表中提取行记录（列对齐版本）。
 * 
 * 为buildRowRecords 的区别：
 * - 使用对齐后的列顺序
 * - 根据 side 参数从对应的物理列读取值
 * - 如果某一列在该side 不存在，对应位置填null
 * 
 * @param alignedColumns 对齐后的列元信息
 * @param primaryKeyColAligned 主键列在对齐后序列中的位置
 * @param side 当前处理的是 base/ours/theirs 哪一侧
 * 
 * 例如：
 * - alignedColumns[2] = { baseCol: 3, oursCol: null, theirsCol: 2 }
 * - 对于 ours 侧，该对齐列的值会是 null（因为 ours 没有这一列）
 */
const buildRowRecordsAligned = (
  ws: any,
  alignedColumns: AlignedColumn[],
  primaryKeyColAligned: number,
  side: 'base' | 'ours' | 'theirs',
  includeEmptyRows = false,
): RowRecord[] => {
  const rows: RowRecord[] = [];
  let index = 0;
  const maxRow = getWorksheetRowUpperBound(ws);
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const row = ws.getRow(rowNumber);
    const values: SimpleCellValue[] = [];
    const nonEmptyCols: number[] = [];
    // 按照对齐后的列顺序读取值
    for (let i = 0; i < alignedColumns.length; i += 1) {
      const colMeta = alignedColumns[i];
      // 根据 side 获取对应的物理列号
      const colNumber =
        side === 'base' ? colMeta.baseCol : side === 'ours' ? colMeta.oursCol : colMeta.theirsCol;
      let value: SimpleCellValue = null;
      // 如果该side 有这一列，则读取值；否则为null
      if (colNumber) {
        const cell = row.getCell(colNumber);
        value = getSimpleValueForMerge(cell?.value);
      }
      values.push(value);
      if (value !== null && value !== '') nonEmptyCols.push(i + 1);
    }
    if (nonEmptyCols.length === 0 && !includeEmptyRows) continue;
    const key =
      primaryKeyColAligned >= 1 && primaryKeyColAligned <= alignedColumns.length
        ? normalizeKeyValue(values[primaryKeyColAligned - 1])
        : null;
    rows.push({ rowNumber, index, values, nonEmptyCols, key });
    index += 1;
  }
  return rows;
};

const buildHeaderRowRecordAligned = (
  ws: any,
  rowNumber: number,
  alignedColumns: AlignedColumn[],
  primaryKeyColAligned: number,
  side: 'base' | 'ours' | 'theirs',
): RowRecord => {
  const values: SimpleCellValue[] = [];
  const nonEmptyCols: number[] = [];
  const row = ws.getRow(rowNumber);
  for (let i = 0; i < alignedColumns.length; i += 1) {
    const colMeta = alignedColumns[i];
    const colNumber =
      side === 'base' ? colMeta.baseCol : side === 'ours' ? colMeta.oursCol : colMeta.theirsCol;
    let value: SimpleCellValue = null;
    if (colNumber) {
      const cell = row.getCell(colNumber);
      value = getSimpleValueForMerge(cell?.value);
    }
    values.push(value);
    if (value !== null && value !== '') nonEmptyCols.push(i + 1);
  }
  const key =
    primaryKeyColAligned >= 1 && primaryKeyColAligned <= alignedColumns.length
      ? normalizeKeyValue(values[primaryKeyColAligned - 1])
      : null;
  return {
    rowNumber,
    index: rowNumber - 1,
    values,
    nonEmptyCols,
    key,
  };
};

/**
 * 判断两行是否完全相等。
 * 
 * 相等的定义：所有非空列的值完全相同。
 * 只比较两行中至少有一行非空的列。
 */
const rowsEqual = (a: RowRecord, b: RowRecord): boolean => {
  // 收集两行的所有非空列
  const cols = new Set<number>();
  a.nonEmptyCols.forEach((c) => cols.add(c));
  b.nonEmptyCols.forEach((c) => cols.add(c));
  // 逐列比较
  for (const col of cols) {
    const av = normalizeCellValue(a.values[col - 1] ?? null);
    const bv = normalizeCellValue(b.values[col - 1] ?? null);
    if (av !== bv) return false;
  }
  return true;
};

/**
 * 计算两行的相似度。
 * 
 * @returns 0-1 之间的相似度，1 表示完全相同。
 * 
 * 算法：
 * 1. 收集两行的所有非空列
 * 2. 计算相同值的列数 / 总列数
 * 3. 跳过两边都为空的列（不计入总数）
 * 
 * 例如：
 * - A行 [1, "abc", null, "xyz"]
 * - B行 [1, "abc", "new", "xyz"]
 * - 相似度= 3/4 = 0.75（第3列不同）
 */
const rowSimilarity = (a: RowRecord, b: RowRecord): number => {
  const cols = new Set<number>();
  a.nonEmptyCols.forEach((c) => cols.add(c));
  b.nonEmptyCols.forEach((c) => cols.add(c));
  if (cols.size === 0) return 1;
  let same = 0;
  let total = 0;
  for (const col of cols) {
    const av = normalizeCellValue(a.values[col - 1] ?? null);
    const bv = normalizeCellValue(b.values[col - 1] ?? null);
    // 跳过两边都为空的列
    if (av === '' && bv === '') continue;
    total += 1;
    if (av === bv) same += 1;
  }
  if (total === 0) return 1;
  return same / total;
};

const rowTokenOf = (row: RowRecord): string => row.values.map((v) => normalizeCellValue(v)).join('||');

/**
 * 计算行的状态（基于三方对比）。
 * 
 * @returns 行状态：
 *   - 'ambiguous': 匹配有歧义（多个候选行）
 *   - 'added': 新增行（base 没有，side 有）
 *   - 'deleted': 删除行（base 有，side 没有）
 *   - 'unchanged': 未变化（内容完全相同）
 *   - 'modified': 修改行（内容不同）
 */
const computeRowStatus = (
  baseRow: RowRecord | null | undefined,
  sideRow: RowRecord | null | undefined,
  isAmbiguous: boolean | undefined,
): RowStatus => {
  if (isAmbiguous) return 'ambiguous';
  if (!baseRow && sideRow) return 'added';
  if (baseRow && !sideRow) return 'deleted';
  if (!baseRow && !sideRow) return 'unchanged';
  if (baseRow && sideRow && rowsEqual(baseRow, sideRow)) return 'unchanged';
  return 'modified';
};


const makeAddress = (col: number, row: number): string => {
  return `${colNumberToLabel(col)}${row}`;
};

const estimateSideIndex = (
  baseIndex: number,
  matchedPairs: Array<{ baseIndex: number; sideIndex: number }>,
): number => {
  if (matchedPairs.length === 0) return baseIndex;
  let prev: { baseIndex: number; sideIndex: number } | null = null;
  let next: { baseIndex: number; sideIndex: number } | null = null;
  for (const p of matchedPairs) {
    if (p.baseIndex < baseIndex) prev = p;
    if (p.baseIndex > baseIndex) {
      next = p;
      break;
    }
  }
  if (prev && next) {
    const t = (baseIndex - prev.baseIndex) / Math.max(1, next.baseIndex - prev.baseIndex);
    return Math.round(prev.sideIndex + t * (next.sideIndex - prev.sideIndex));
  }
  if (prev) return prev.sideIndex + (baseIndex - prev.baseIndex);
  if (next) return next.sideIndex - (next.baseIndex - baseIndex);
  return baseIndex;
};

/**
 * 计算最长公共子序列（LCS）并返回匹配对。
 * 
 * 用于删行对齐的锁点匹配：找到两个序列中确定相同的元素作为“锁点”。
 * 
 * @param a 第一个字符串数组
 * @param b 第二个字符串数组
 * @returns 匹配对数组，按照出现顺序排列
 * 
 * 例如：
 * - a = ["A", "B", "C", "D"]
 * - b = ["A", "X", "B", "D"]
 * - 返回: [{ aIndex: 0, bIndex: 0 }, { aIndex: 1, bIndex: 2 }, { aIndex: 3, bIndex: 3 }]
 * - 即A, B, D 三个元素是公共的
 * 
 * 算法：动态规划+ 回溯
 * - dp[i][j] = a[0..i-1] 和b[0..j-1] 的LCS 长度
 * - 回溯找到实际匹配的位置
 */
const lcsMatchPairs = (a: string[], b: string[]): Array<{ aIndex: number; bIndex: number }> => {
  const n = a.length;
  const m = b.length;
  // 动态规划表：dp[i][j] = LCS 长度
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  // 填表：计算LCS 长度
  for (let i = 1; i <= n; i += 1) {
    for (let j = 1; j <= m; j += 1) {
      if (a[i - 1] === b[j - 1]) dp[i][j] = dp[i - 1][j - 1] + 1;  // 匹配，长度+1
      else dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);        // 不匹配，取最大值
    }
  }
  // 回溯：从 dp 表中提取实际匹配对
  const pairs: Array<{ aIndex: number; bIndex: number }> = [];
  let i = n;
  let j = m;
  while (i > 0 && j > 0) {
    if (a[i - 1] === b[j - 1]) {
      // 当前元素匹配，记录并继续回溯
      pairs.push({ aIndex: i - 1, bIndex: j - 1 });
      i -= 1;
      j -= 1;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      i -= 1;  // 向上回溯
    } else {
      j -= 1;  // 向左回溯
    }
  }
  // 回溯是从后往前，需要反转
  return pairs.reverse();
};


/**
 * 基于主键列对齐行。
 * 
 * 这是行对齐的主要方法，适用于有唯一标识列（如ID）的数据。
 * 
 * @param baseRows base 的行记录
 * @param oursRows ours 的行记录
 * @param theirsRows theirs 的行记录
 * @param keyCol 主键列号（1-based）
 * @param rowSimilarityThreshold 相似度阈值（用于歧义检测）
 * @returns 对齐结果 + 歧义行集合
 * 
 * 算法步骤：
 * 1. 按主键值分组：Map<key, RowRecord[]>
 * 2. 对每个主键值：
 *    - 如果 base/ours/theirs 都有且每侧只有 1 条 → 直接匹配
 *    - 如果某侧有多条相同主键 → 检测歧义（相似度匹配）
 * 3. 返回对齐后的三元组：(base, ours, theirs)
 * 
 * 歧义场景：
 * - 主键值相同但其他列内容不同的多行
 * - 此时无法确定哪一行对应哪一行，标记为 ambiguous
 */
const alignRowsByKey = (
  baseRows: RowRecord[],
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
  keyCol: number,
  rowSimilarityThreshold: number,
): { aligned: AlignedRow[]; ambiguousOurs: Set<number>; ambiguousTheirs: Set<number> } => {
  const groupByKey = (rows: RowRecord[]) => {
    const m = new Map<string, RowRecord[]>();
    rows.forEach((r) => {
      if (!r.key) return;
      if (!m.has(r.key)) m.set(r.key, []);
      m.get(r.key)!.push(r);
    });
    return m;
  };
  const rowSimilarityIgnoringKey = (a: RowRecord, b: RowRecord): number => {
    if (keyCol < 1) return rowSimilarity(a, b);
    const cols = new Set<number>();
    a.nonEmptyCols.forEach((c) => cols.add(c));
    b.nonEmptyCols.forEach((c) => cols.add(c));
    if (cols.size === 0) return 1;
    let same = 0;
    let total = 0;
    for (const col of cols) {
      if (col === keyCol) continue;
      const av = normalizeCellValue(a.values[col - 1] ?? null);
      const bv = normalizeCellValue(b.values[col - 1] ?? null);
      if (av === '' && bv === '') continue;
      total += 1;
      if (av === bv) same += 1;
    }
    if (total === 0) return 1;
    return same / total;
  };

  const baseByKeyList = groupByKey(baseRows);
  const oursByKeyList = groupByKey(oursRows);
  const theirsByKeyList = groupByKey(theirsRows);

  const baseCounts = new Map<string, number>();
  baseByKeyList.forEach((list, key) => baseCounts.set(key, list.length));
  const oursCounts = new Map<string, number>();
  oursByKeyList.forEach((list, key) => oursCounts.set(key, list.length));
  const theirsCounts = new Map<string, number>();
  theirsByKeyList.forEach((list, key) => theirsCounts.set(key, list.length));

  const occurrenceIndex = (rows: RowRecord[]) => {
    const occ = new Map<number, number>();
    const counters = new Map<string, number>();
    rows.forEach((r) => {
      if (!r.key) return;
      const next = (counters.get(r.key) ?? 0) + 1;
      counters.set(r.key, next);
      occ.set(r.index, next - 1);
    });
    return occ;
  };

  const baseOcc = occurrenceIndex(baseRows);

  const matchedOursRows = new Set<number>();
  const matchedTheirsRows = new Set<number>();

  const matchedInOurs: Array<{ baseIndex: number; sideIndex: number }> = [];
  const matchedInTheirs: Array<{ baseIndex: number; sideIndex: number }> = [];
  const pickNoKeyRowMatch = (
    baseRow: RowRecord,
    sideRows: RowRecord[],
    matchedRows: Set<number>,
  ): { row: RowRecord | null; ambiguous: boolean } => {
    const candidates = sideRows.filter((row) => !matchedRows.has(row.index) && !row.key);
    if (candidates.length === 0) {
      return { row: null, ambiguous: false };
    }

    const baseToken = rowTokenOf(baseRow);
    const exactMatches = candidates.filter((row) => rowTokenOf(row) === baseToken);
    if (exactMatches.length > 0) {
      exactMatches.sort((a, b) => Math.abs(a.index - baseRow.index) - Math.abs(b.index - baseRow.index));
      return { row: exactMatches[0] ?? null, ambiguous: false };
    }

    const scored = candidates
      .map((row) => ({ row, score: rowSimilarity(baseRow, row) }))
      .filter((entry) => entry.score >= rowSimilarityThreshold)
      .sort((a, b) => b.score - a.score);
    const best = scored[0];
    const second = scored[1];
    if (!best) {
      return { row: null, ambiguous: false };
    }
    if (second && best.score - second.score < 0.1) {
      return { row: null, ambiguous: true };
    }
    return { row: best.row, ambiguous: false };
  };

  const alignedBase: AlignedRow[] = baseRows.map((baseRow) => {
    const key = baseRow.key ?? null;
    if (!key) {
      const oursFallback = pickNoKeyRowMatch(baseRow, oursRows, matchedOursRows);
      const theirsFallback = pickNoKeyRowMatch(baseRow, theirsRows, matchedTheirsRows);
      if (oursFallback.row) {
        matchedOursRows.add(oursFallback.row.index);
        matchedInOurs.push({ baseIndex: baseRow.index, sideIndex: oursFallback.row.index });
      }
      if (theirsFallback.row) {
        matchedTheirsRows.add(theirsFallback.row.index);
        matchedInTheirs.push({ baseIndex: baseRow.index, sideIndex: theirsFallback.row.index });
      }
      return {
        base: baseRow,
        ours: oursFallback.row,
        theirs: theirsFallback.row,
        key,
        ambiguousOurs: oursFallback.ambiguous,
        ambiguousTheirs: theirsFallback.ambiguous,
      };
    }

    const baseList = baseByKeyList.get(key) ?? [];
    const oursList = oursByKeyList.get(key) ?? [];
    const theirsList = theirsByKeyList.get(key) ?? [];
    const baseCount = baseList.length;
    const oursCount = oursList.length;
    const theirsCount = theirsList.length;
    const occIndex = baseOcc.get(baseRow.index) ?? 0;

    let ours: RowRecord | null = null;
    let theirs: RowRecord | null = null;
    let ambiguousOurs = false;
    let ambiguousTheirs = false;
    const pickBestMatch = (
      candidates: RowRecord[],
      similarityFn: (a: RowRecord, b: RowRecord) => number,
      threshold: number,
      delta: number,
    ) => {
      if (candidates.length === 0) return null;
      const scored = candidates
        .map((r) => ({ row: r, score: similarityFn(baseRow, r) }))
        .sort((a, b) => b.score - a.score);
      const best = scored[0];
      const second = scored[1];
      if (!best || best.score < threshold) return null;
      if (second && best.score - second.score < delta) return null;
      return best.row;
    };

    if (oursCount === 0) {
      const candidates = oursRows.filter((r) => !matchedOursRows.has(r.index));
      const best = pickBestMatch(candidates, rowSimilarityIgnoringKey, rowSimilarityThreshold, 0.05);
      if (best) ours = best;
      else ours = null;
    } else if (oursCount === 1 && baseCount === 1) {
      ours = oursList[0] ?? null;
    } else if (oursCount === baseCount && baseCount > 0) {
      ours = oursList[occIndex] ?? null;
    } else {
      const candidates = oursList.filter((r) => !matchedOursRows.has(r.index));
      if (candidates.length === 1) {
        const only = candidates[0];
        if (rowSimilarity(baseRow, only) >= rowSimilarityThreshold) ours = only;
        else ambiguousOurs = true;
      } else {
        const best = pickBestMatch(candidates, rowSimilarity, rowSimilarityThreshold, 0.1);
        if (best) ours = best;
        else ambiguousOurs = true;
      }
    }

    if (theirsCount === 0) {
      const candidates = theirsRows.filter((r) => !matchedTheirsRows.has(r.index));
      const best = pickBestMatch(candidates, rowSimilarityIgnoringKey, rowSimilarityThreshold, 0.05);
      if (best) theirs = best;
      else theirs = null;
    } else if (theirsCount === 1 && baseCount === 1) {
      theirs = theirsList[0] ?? null;
    } else if (theirsCount === baseCount && baseCount > 0) {
      theirs = theirsList[occIndex] ?? null;
    } else {
      const candidates = theirsList.filter((r) => !matchedTheirsRows.has(r.index));
      if (candidates.length === 1) {
        const only = candidates[0];
        if (rowSimilarity(baseRow, only) >= rowSimilarityThreshold) theirs = only;
        else ambiguousTheirs = true;
      } else {
        const best = pickBestMatch(candidates, rowSimilarity, rowSimilarityThreshold, 0.1);
        if (best) theirs = best;
        else ambiguousTheirs = true;
      }
    }

    if (ours) {
      matchedOursRows.add(ours.index);
      matchedInOurs.push({ baseIndex: baseRow.index, sideIndex: ours.index });
    }
    if (theirs) {
      matchedTheirsRows.add(theirs.index);
      matchedInTheirs.push({ baseIndex: baseRow.index, sideIndex: theirs.index });
    }

    return {
      base: baseRow,
      ours,
      theirs,
      key,
      ambiguousOurs,
      ambiguousTheirs,
    };
  });

  matchedInOurs.sort((a, b) => a.sideIndex - b.sideIndex);
  matchedInTheirs.sort((a, b) => a.sideIndex - b.sideIndex);

  const gapsOurs = new Map<number, RowRecord[]>();
  const gapsTheirs = new Map<number, RowRecord[]>();

  const pushGap = (gaps: Map<number, RowRecord[]>, gap: number, row: RowRecord) => {
    if (!gaps.has(gap)) gaps.set(gap, []);
    gaps.get(gap)!.push(row);
  };

  const placeInGaps = (
    rows: RowRecord[],
    matchedRowIndices: Set<number>,
    matchedPairs: Array<{ baseIndex: number; sideIndex: number }>,
    gaps: Map<number, RowRecord[]>,
  ) => {
    const matchedBaseBySideIndex = matchedPairs.slice().sort((a, b) => a.sideIndex - b.sideIndex);
    for (const row of rows) {
      if (matchedRowIndices.has(row.index)) continue;
      let gap = -1;
      for (const p of matchedBaseBySideIndex) {
        if (p.sideIndex < row.index) gap = p.baseIndex;
        if (p.sideIndex >= row.index) break;
      }
      pushGap(gaps, gap, row);
    }
  };

  placeInGaps(oursRows, matchedOursRows, matchedInOurs, gapsOurs);
  placeInGaps(theirsRows, matchedTheirsRows, matchedInTheirs, gapsTheirs);

  const aligned: AlignedRow[] = [];
  const addGapRows = (gapIndex: number) => {
    const oursGap = gapsOurs.get(gapIndex) ?? [];
    const theirsGap = gapsTheirs.get(gapIndex) ?? [];
    const unmatchedTheirsByKey = new Map<string, RowRecord[]>();
    const unmatchedTheirsWithoutKey: RowRecord[] = [];
    for (const r of theirsGap) {
      if (!r.key) {
        unmatchedTheirsWithoutKey.push(r);
        continue;
      }
      if (!unmatchedTheirsByKey.has(r.key)) unmatchedTheirsByKey.set(r.key, []);
      unmatchedTheirsByKey.get(r.key)!.push(r);
    }
    for (const r of oursGap) {
      const sameKeyTheirs = r.key ? unmatchedTheirsByKey.get(r.key) ?? [] : [];
      const matchedTheirs = sameKeyTheirs.shift() ?? null;
      if (r.key && sameKeyTheirs.length === 0) unmatchedTheirsByKey.delete(r.key);
      if (matchedTheirs) {
        aligned.push({ ours: r, theirs: matchedTheirs, key: r.key, ambiguousOurs: false, ambiguousTheirs: false });
        continue;
      }
      const ambiguous = !r.key;
      aligned.push({ ours: r, key: r.key ?? null, ambiguousOurs: ambiguous });
    }
    for (const r of unmatchedTheirsWithoutKey) {
      const ambiguous = !r.key;
      aligned.push({ theirs: r, key: r.key ?? null, ambiguousTheirs: ambiguous });
    }
    for (const rows of unmatchedTheirsByKey.values()) {
      for (const r of rows) {
        const ambiguous = !r.key;
        aligned.push({ theirs: r, key: r.key ?? null, ambiguousTheirs: ambiguous });
      }
    }
  };

  addGapRows(-1);
  for (const baseRow of alignedBase) {
    aligned.push(baseRow);
    addGapRows(baseRow.base?.index ?? -1);
  }

  return { aligned, ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
};

const alignRowsBySequence = (
  baseRows: RowRecord[],
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
): { aligned: AlignedRow[]; ambiguousOurs: Set<number>; ambiguousTheirs: Set<number> } => {
  const buildTokens = (rows: RowRecord[]) => rows.map((r) => rowTokenOf(r));

  const similarityThreshold = 0.7;
  const similarityDelta = 0.05;
  const windowSize = 12;

  const alignOneSide = (sideRows: RowRecord[]) => {
    const baseTokens = buildTokens(baseRows);
    const sideTokens = buildTokens(sideRows);
    const ops = diffArraysToOps(baseTokens, sideTokens);
    const matched = new Map<number, number>();
    const deletes = new Set<number>();
    const inserts = new Set<number>();
    for (const op of ops) {
      const hasBase = (idx: number) => idx >= 0 && idx < baseRows.length;
      const hasSide = (idx: number) => idx >= 0 && idx < sideRows.length;
      if (op.type === 'equal') {
        if (hasBase(op.aIndex) && hasSide(op.bIndex)) {
          matched.set(op.aIndex, op.bIndex);
        }
      } else if (op.type === 'delete') {
        if (hasBase(op.aIndex)) deletes.add(op.aIndex);
      } else {
        if (hasSide(op.bIndex)) inserts.add(op.bIndex);
      }
    }

    const unmatchedDeletes = new Set<number>(deletes);
    const unmatchedInserts = new Set<number>(inserts);

    // 优先匹配“完全相同”的行（token 相同），避免重复行造成错配
    const insertByToken = new Map<string, number[]>();
    for (const idx of unmatchedInserts) {
      const token = sideTokens[idx] ?? '';
      if (!insertByToken.has(token)) insertByToken.set(token, []);
      insertByToken.get(token)!.push(idx);
    }
    insertByToken.forEach((list) => list.sort((a, b) => a - b));

    const matchExactToken = (baseIndex: number) => {
      const token = baseTokens[baseIndex] ?? '';
      const list = insertByToken.get(token);
      if (!list || list.length === 0) return null;
      // 选择距离期望位置最近的插入点
      const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
      matchedPairs.sort((a, b) => a.baseIndex - b.baseIndex);
      const expected = estimateSideIndex(baseIndex, matchedPairs);
      let bestPos = 0;
      let bestDist = Math.abs(list[0] - expected);
      for (let i = 1; i < list.length; i += 1) {
        const dist = Math.abs(list[i] - expected);
        if (dist < bestDist) {
          bestDist = dist;
          bestPos = i;
        }
      }
      const sideIndex = list.splice(bestPos, 1)[0];
      if (list.length === 0) insertByToken.delete(token);
      return sideIndex ?? null;
    };

    for (const baseIndex of deletes) {
      const sideIndex = matchExactToken(baseIndex);
      if (sideIndex == null) continue;
      matched.set(baseIndex, sideIndex);
      unmatchedDeletes.delete(baseIndex);
      unmatchedInserts.delete(sideIndex);
    }

    const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
    matchedPairs.sort((a, b) => a.baseIndex - b.baseIndex);

    const ambiguousBase = new Set<number>();
    const ambiguousSide = new Set<number>();
    for (const baseIndex of unmatchedDeletes) {
      const baseRow = baseRows[baseIndex];
      if (!baseRow) continue;
      const expected = estimateSideIndex(baseIndex, matchedPairs);
      const candidates: Array<{ index: number; score: number }> = [];
      for (const sideIndex of unmatchedInserts) {
        if (sideIndex < expected - windowSize || sideIndex > expected + windowSize) continue;
        const sideRow = sideRows[sideIndex];
        if (!sideRow) continue;
        const score = rowSimilarity(baseRow, sideRow);
        if (score >= similarityThreshold) candidates.push({ index: sideIndex, score });
      }
      if (candidates.length === 0) continue;
      candidates.sort((a, b) => b.score - a.score);
      const best = candidates[0];
      const second = candidates[1];
      if (second && second.score >= similarityThreshold && best.score - second.score < similarityDelta) {
        ambiguousBase.add(baseIndex);
        candidates.forEach((c) => ambiguousSide.add(c.index));
        continue;
      }
      matched.set(baseIndex, best.index);
      unmatchedInserts.delete(best.index);
    }

    return { matched, unmatchedInserts, ambiguousBase, ambiguousSide };
  };

  const oursAlign = alignOneSide(oursRows);
  const theirsAlign = alignOneSide(theirsRows);

  const gapsOurs = new Map<number, RowRecord[]>();
  const gapsTheirs = new Map<number, RowRecord[]>();

  const buildGaps = (
    sideRows: RowRecord[],
    matched: Map<number, number>,
    unmatchedInserts: Set<number>,
    gaps: Map<number, RowRecord[]>,
  ) => {
    const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
    matchedPairs.sort((a, b) => a.sideIndex - b.sideIndex);
    for (const sideIndex of unmatchedInserts) {
      const row = sideRows[sideIndex];
      if (!row) continue;
      let gap = -1;
      for (const p of matchedPairs) {
        if (p.sideIndex < sideIndex) gap = p.baseIndex;
        if (p.sideIndex >= sideIndex) break;
      }
      if (!gaps.has(gap)) gaps.set(gap, []);
      gaps.get(gap)!.push(row);
    }
  };

  buildGaps(oursRows, oursAlign.matched, oursAlign.unmatchedInserts, gapsOurs);
  buildGaps(theirsRows, theirsAlign.matched, theirsAlign.unmatchedInserts, gapsTheirs);

  const aligned: AlignedRow[] = [];
  const addGapRows = (gapIndex: number) => {
    const oursGap = gapsOurs.get(gapIndex) ?? [];
    const theirsGap = gapsTheirs.get(gapIndex) ?? [];
    for (const r of oursGap) {
      aligned.push({ ours: r, ambiguousOurs: oursAlign.ambiguousSide.has(r.index) });
    }
    for (const r of theirsGap) {
      aligned.push({ theirs: r, ambiguousTheirs: theirsAlign.ambiguousSide.has(r.index) });
    }
  };

  addGapRows(-1);
  for (let i = 0; i < baseRows.length; i += 1) {
    const baseRow = baseRows[i];
    const oursIndex = oursAlign.matched.get(i);
    const theirsIndex = theirsAlign.matched.get(i);
    aligned.push({
      base: baseRow,
      ours: typeof oursIndex === 'number' ? oursRows[oursIndex] : null,
      theirs: typeof theirsIndex === 'number' ? theirsRows[theirsIndex] : null,
      ambiguousOurs: oursAlign.ambiguousBase.has(i) || (typeof oursIndex === 'number' && oursAlign.ambiguousSide.has(oursIndex)),
      ambiguousTheirs:
        theirsAlign.ambiguousBase.has(i) || (typeof theirsIndex === 'number' && theirsAlign.ambiguousSide.has(theirsIndex)),
    });
    addGapRows(i);
  }

  return { aligned, ambiguousOurs: oursAlign.ambiguousSide, ambiguousTheirs: theirsAlign.ambiguousSide };
};

// Align rows by content using unique anchors, then diff segments to reduce misalignment noise.
const alignRowsByContent = (
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
): { aligned: AlignedRow[]; ambiguousOurs: Set<number>; ambiguousTheirs: Set<number> } => {
  if (oursRows.length === 0 && theirsRows.length === 0) {
    return { aligned: [], ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
  }
  if (oursRows.length === 0) {
    return { aligned: theirsRows.map((r) => ({ theirs: r })), ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
  }
  if (theirsRows.length === 0) {
    return {
      aligned: oursRows.map((r) => ({ base: r, ours: r })),
      ambiguousOurs: new Set(),
      ambiguousTheirs: new Set(),
    };
  }

  const tokenOf = (r: RowRecord) => rowTokenOf(r);
  const oursTokens = oursRows.map((r) => tokenOf(r));
  const theirsTokens = theirsRows.map((r) => tokenOf(r));

  const countTokens = (tokens: string[]) => {
    const m = new Map<string, number>();
    tokens.forEach((t) => m.set(t, (m.get(t) ?? 0) + 1));
    return m;
  };
  const oursCount = countTokens(oursTokens);
  const theirsCount = countTokens(theirsTokens);
  const theirsUniqueIndex = new Map<string, number>();
  theirsTokens.forEach((t, idx) => {
    if ((theirsCount.get(t) ?? 0) === 1) theirsUniqueIndex.set(t, idx);
  });

  const anchors: Array<{ o: number; t: number }> = [];
  oursTokens.forEach((t, o) => {
    if ((oursCount.get(t) ?? 0) !== 1) return;
    const tIdx = theirsUniqueIndex.get(t);
    if (typeof tIdx === 'number') anchors.push({ o, t: tIdx });
  });

  const selectIncreasingAnchors = (pairs: Array<{ o: number; t: number }>) => {
    if (pairs.length === 0) return [];
    // pairs are already in ours order; compute LIS on t
    const tails: number[] = [];
    const prev = new Array(pairs.length).fill(-1);
    for (let i = 0; i < pairs.length; i += 1) {
      const tVal = pairs[i].t;
      let l = 0;
      let r = tails.length;
      while (l < r) {
        const m = Math.floor((l + r) / 2);
        if (pairs[tails[m]].t < tVal) l = m + 1;
        else r = m;
      }
      if (l > 0) prev[i] = tails[l - 1];
      if (l === tails.length) tails.push(i);
      else tails[l] = i;
    }
    const result: Array<{ o: number; t: number }> = [];
    let k = tails[tails.length - 1];
    while (k >= 0) {
      result.push(pairs[k]);
      k = prev[k];
    }
    return result.reverse();
  };

  const inOrderAnchors = selectIncreasingAnchors(anchors);
  if (inOrderAnchors.length === 0) {
    // fallback to sequence alignment with ours as base
    return alignRowsBySequence(oursRows, oursRows, theirsRows);
  }

  const aligned: AlignedRow[] = [];
  const addSegment = (oStart: number, oEnd: number, tStart: number, tEnd: number) => {
    const oSeg = oursRows.slice(oStart, oEnd);
    const tSeg = theirsRows.slice(tStart, tEnd);
    if (oSeg.length === 0 && tSeg.length === 0) return;
    if (oSeg.length === 0) {
      tSeg.forEach((r) => aligned.push({ theirs: r }));
      return;
    }
    if (tSeg.length === 0) {
      oSeg.forEach((r) => aligned.push({ base: r, ours: r }));
      return;
    }
    const segAligned = alignRowsBySequence(oSeg, oSeg, tSeg).aligned;
    aligned.push(...segAligned);
  };

  let prevO = -1;
  let prevT = -1;
  for (const anchor of inOrderAnchors) {
    addSegment(prevO + 1, anchor.o, prevT + 1, anchor.t);
    aligned.push({
      base: oursRows[anchor.o],
      ours: oursRows[anchor.o],
      theirs: theirsRows[anchor.t],
    });
    prevO = anchor.o;
    prevT = anchor.t;
  }
  addSegment(prevO + 1, oursRows.length, prevT + 1, theirsRows.length);

  return { aligned, ambiguousOurs: new Set(), ambiguousTheirs: new Set() };
};

const buildMergeSheetWithRowAlign = (
  baseWs: any,
  oursWs: any,
  theirsWs: any,
  primaryKeyCol: number,
  frozenRowCount: number,
  rowSimilarityThreshold: number,
  compareMode: ThreeWayCompareMode,
): MergeSheetData => {
  const runtimeConfig = getThreeWayRuntimeConfig(compareMode);
  const sheetsEqualByCoordinate = (a: any, b: any) => {
    const maxRow = Math.max(getRowCount(a), getRowCount(b));
    const maxCol = Math.max(getColCount(a), getColCount(b));
    for (let r = 1; r <= maxRow; r += 1) {
      const rowA = a.getRow(r);
      const rowB = b.getRow(r);
      for (let c = 1; c <= maxCol; c += 1) {
        const av = normalizeCellValue(getSimpleValueForMerge(rowA.getCell(c)?.value));
        const bv = normalizeCellValue(getSimpleValueForMerge(rowB.getCell(c)?.value));
        if (av !== bv) return false;
      }
    }
    return true;
  };
  const getRowCount = (ws: any) =>
    getWorksheetRowUpperBound(ws);
  const getColCount = (ws: any) =>
    (ws?.actualColumnCount ?? 0) > 0 ? ws.actualColumnCount : ws?.columnCount ?? 0;
  // note: hasExactDiff will be derived from visible diff cells (ours/theirs/conflict)
  const detectKeyColByThreshold = (
    rows: RowRecord[],
    totalCols: number,
    minCoverage: number,
    minUniq: number,
  ) => {
    const total = rows.length;
    if (total === 0) return null;
    const minNonEmpty = Math.max(3, Math.floor(total * minCoverage));
    let bestCol: number | null = null;
    let bestScore = 0;
    for (let col = 1; col <= totalCols; col += 1) {
      let nonEmpty = 0;
      const uniq = new Set<string>();
      for (const row of rows) {
        const v = normalizeKeyValue(row.values[col - 1] ?? null);
        if (v == null) continue;
        nonEmpty += 1;
        uniq.add(v);
      }
      if (nonEmpty < minNonEmpty) continue;
      const coverage = nonEmpty / total;
      const uniqueness = uniq.size / Math.max(1, nonEmpty);
      if (coverage < minCoverage || uniqueness < minUniq) continue;
      const score = coverage * uniqueness;
      if (score > bestScore) {
        bestScore = score;
        bestCol = col;
      }
    }
    return bestCol;
  };
  const detectImplicitKeyCol = (rows: RowRecord[], totalCols: number) =>
    detectKeyColByThreshold(rows, totalCols, 0.8, 0.9);
  const detectWeakKeyCol = (rows: RowRecord[], totalCols: number) =>
    detectKeyColByThreshold(rows, totalCols, 0.6, 0.9);
  const detectHeaderKeyCol = (ws: any, totalCols: number, headerRows: number) => {
    const maxHeader = Math.max(1, Math.min(Math.floor(headerRows), 3));
    for (let r = 1; r <= maxHeader; r += 1) {
      const row = ws.getRow(r);
      for (let c = 1; c <= totalCols; c += 1) {
        const raw = getSimpleValueForMerge(row.getCell(c)?.value);
        if (raw == null) continue;
        const text = String(raw).trim();
        if (!text) continue;
        if (/id/i.test(text) || /编号|主键/.test(text)) {
          return c;
        }
      }
    }
    return null;
  };
  const applyKeyFromColumn = (rows: RowRecord[], col: number): RowRecord[] =>
    rows.map((r) => ({
      ...r,
      key: col >= 1 ? normalizeKeyValue(r.values[col - 1] ?? null) : null,
    }));
  const rawColCount = Math.max(
    baseWs?.actualColumnCount ?? baseWs?.columnCount ?? 0,
    oursWs?.actualColumnCount ?? oursWs?.columnCount ?? 0,
    theirsWs?.actualColumnCount ?? theirsWs?.columnCount ?? 0,
  );
  const headerCount = resolveStructuredHeaderRowCount([baseWs, oursWs, theirsWs], frozenRowCount);
  const baseWsForAlign = runtimeConfig.alignBaseSide === 'ours' ? oursWs : baseWs;
  const alignedColumns = buildAlignedColumns(baseWsForAlign, oursWs, theirsWs, headerCount);
  const columnsMeta: MergeColumnMeta[] = alignedColumns.map((c, idx) => ({
    col: idx + 1,
    baseCol: c.baseCol ?? null,
    oursCol: c.oursCol ?? null,
    theirsCol: c.theirsCol ?? null,
  }));
  const colCount = Math.max(alignedColumns.length, 0);
  const useKey = primaryKeyCol >= 1 && primaryKeyCol <= rawColCount;
  const shouldAutoDetectKey = primaryKeyCol < 0;
  const shouldShortCircuit = runtimeConfig.shortCircuitWhenOursEqualsTheirs && sheetsEqualByCoordinate(oursWs, theirsWs);
  const mapRawToAligned = (rawCol: number, side: 'base' | 'ours' | 'theirs'): number | null => {
    if (rawCol < 1) return null;
    const idx = alignedColumns.findIndex((c) =>
      side === 'base' ? c.baseCol === rawCol : side === 'ours' ? c.oursCol === rawCol : c.theirsCol === rawCol,
    );
    return idx >= 0 ? idx + 1 : null;
  };
  const keyColAligned = useKey ? mapRawToAligned(primaryKeyCol, 'ours') ?? -1 : -1;

  const baseRowsForKeyDetection = buildRowRecordsAligned(baseWsForAlign, alignedColumns, keyColAligned, 'base').filter(
    (r) => r.rowNumber > headerCount,
  );
  const oursRowsForKeyDetection = buildRowRecordsAligned(oursWs, alignedColumns, keyColAligned, 'ours').filter(
    (r) => r.rowNumber > headerCount,
  );
  const theirsRowsForKeyDetection = buildRowRecordsAligned(theirsWs, alignedColumns, keyColAligned, 'theirs').filter(
    (r) => r.rowNumber > headerCount,
  );
  const implicitKeyCol = shouldAutoDetectKey ? detectImplicitKeyCol(baseRowsForKeyDetection, colCount) : null;
  const headerKeyColRaw =
    shouldAutoDetectKey && implicitKeyCol == null ? detectHeaderKeyCol(baseWsForAlign, rawColCount, headerCount) : null;
  const headerKeyCol = headerKeyColRaw ? mapRawToAligned(headerKeyColRaw, 'base') : null;
  const weakKeyCol =
    shouldAutoDetectKey && implicitKeyCol == null && headerKeyCol == null
      ? detectWeakKeyCol(baseRowsForKeyDetection, colCount)
      : null;
  const alignKeyCol = useKey ? keyColAligned ?? -1 : shouldAutoDetectKey ? implicitKeyCol ?? headerKeyCol ?? weakKeyCol ?? -1 : -1;
  const includeEmptyAlignedRows = compareMode === 'merge' && alignKeyCol < 1;
  const baseRows = buildRowRecordsAligned(
    baseWsForAlign,
    alignedColumns,
    keyColAligned,
    'base',
    includeEmptyAlignedRows,
  ).filter((r) => r.rowNumber > headerCount);
  const oursRows = buildRowRecordsAligned(
    oursWs,
    alignedColumns,
    keyColAligned,
    'ours',
    includeEmptyAlignedRows,
  ).filter((r) => r.rowNumber > headerCount);
  const theirsRows = buildRowRecordsAligned(
    theirsWs,
    alignedColumns,
    keyColAligned,
    'theirs',
    includeEmptyAlignedRows,
  ).filter((r) => r.rowNumber > headerCount);
  const primaryKeySource: PrimaryKeySource = useKey
    ? alignKeyCol >= 1
      ? 'manual'
      : 'none'
    : implicitKeyCol != null
      ? 'auto-implicit'
      : headerKeyCol != null
        ? 'auto-header'
        : weakKeyCol != null
          ? 'auto-weak'
          : 'none';
  const primaryKeyAlignedCol = alignKeyCol >= 1 ? alignKeyCol : null;
  const primaryKeyOursCol = primaryKeyAlignedCol ? alignedColumns[primaryKeyAlignedCol - 1]?.oursCol ?? null : null;
  if (shouldShortCircuit) {
    return {
      sheetName: baseWs.name,
      cells: [],
      rowsMeta: [],
      hasExactDiff: false,
      columnsMeta,
      primaryKeyAlignedCol,
      primaryKeyOursCol,
      primaryKeySource,
    };
  }
  const alignedResult =
    alignKeyCol >= 1
      ? alignRowsByKey(
          applyKeyFromColumn(baseRows, alignKeyCol),
          applyKeyFromColumn(oursRows, alignKeyCol),
          applyKeyFromColumn(theirsRows, alignKeyCol),
          alignKeyCol,
          rowSimilarityThreshold,
        )
      : runtimeConfig.useTwoWayContentAlignment
        ? alignRowsByContent(oursRows, theirsRows)
        : alignRowsBySequence(baseRows, oursRows, theirsRows);

  const aligned = alignedResult.aligned;

  const rowsMeta: MergeRowMeta[] = [];
  // 1) Header rows: compare by fixed row number (no alignment)
  const metaKeyCol = alignKeyCol >= 1 ? alignKeyCol : keyColAligned;
  for (let r = 1; r <= headerCount; r += 1) {
    const baseRow = buildHeaderRowRecordAligned(baseWsForAlign, r, alignedColumns, metaKeyCol, 'base');
    const oursRow = buildHeaderRowRecordAligned(oursWs, r, alignedColumns, metaKeyCol, 'ours');
    const theirsRow = buildHeaderRowRecordAligned(theirsWs, r, alignedColumns, metaKeyCol, 'theirs');
    const oursSim = rowSimilarity(baseRow, oursRow);
    const theirsSim = rowSimilarity(baseRow, theirsRow);
    rowsMeta.push({
      visualRowNumber: r,
      key: baseRow.key ?? oursRow.key ?? theirsRow.key ?? null,
      baseRowNumber: r,
      oursRowNumber: r,
      theirsRowNumber: r,
      oursSimilarity: oursSim,
      theirsSimilarity: theirsSim,
      oursStatus: computeRowStatus(baseRow, oursRow, false),
      theirsStatus: computeRowStatus(baseRow, theirsRow, false),
    });
  }
  // 2) Body rows: aligned
  aligned.forEach((row, idx) => {
    const visualRowNumber = headerCount + idx + 1;
    const oursSim = row.base && row.ours ? rowSimilarity(row.base, row.ours) : null;
    const theirsSim = row.base && row.theirs ? rowSimilarity(row.base, row.theirs) : null;
    rowsMeta.push({
      visualRowNumber,
      key: alignKeyCol >= 1 ? row.key ?? row.base?.key ?? row.ours?.key ?? row.theirs?.key ?? null : null,
      baseRowNumber: row.base?.rowNumber ?? null,
      oursRowNumber: row.ours?.rowNumber ?? null,
      theirsRowNumber: row.theirs?.rowNumber ?? null,
      oursSimilarity: oursSim,
      theirsSimilarity: theirsSim,
      oursStatus: computeRowStatus(row.base ?? null, row.ours ?? null, row.ambiguousOurs),
      theirsStatus: computeRowStatus(row.base ?? null, row.theirs ?? null, row.ambiguousTheirs),
    });
  });

  const cells: MergeCell[] = [];
  let hasExactDiff = false;

  // Header rows diff by fixed row number and classify with base / ours / theirs
  for (let r = 1; r <= headerCount; r += 1) {
    const baseRow = buildHeaderRowRecordAligned(baseWsForAlign, r, alignedColumns, metaKeyCol, 'base');
    const oursRow = buildHeaderRowRecordAligned(oursWs, r, alignedColumns, metaKeyCol, 'ours');
    const theirsRow = buildHeaderRowRecordAligned(theirsWs, r, alignedColumns, metaKeyCol, 'theirs');
    const cols = new Set<number>();
    baseRow.nonEmptyCols.forEach((c) => cols.add(c));
    oursRow.nonEmptyCols.forEach((c) => cols.add(c));
    theirsRow.nonEmptyCols.forEach((c) => cols.add(c));
    for (const col of cols) {
      const baseValue = baseRow.values[col - 1] ?? null;
      const oursValue = oursRow.values[col - 1] ?? null;
      const theirsValue = theirsRow.values[col - 1] ?? null;
      const { status, mergedValue } = classifyThreeWayCell(baseValue, oursValue, theirsValue);

      if (status !== 'unchanged') {
        const colMeta = alignedColumns[col - 1];
        cells.push({
          address: makeAddress(col, r),
          row: r,
          col,
          baseCol: colMeta?.baseCol ?? null,
          oursCol: colMeta?.oursCol ?? null,
          theirsCol: colMeta?.theirsCol ?? null,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
        hasExactDiff = true;
      }
    }
  }

  // Body rows diff via alignment and classify with base / ours / theirs
  aligned.forEach((row, visualIndex) => {
    const visualRowNumber = headerCount + visualIndex + 1;
    const cols = new Set<number>();
    row.base?.nonEmptyCols.forEach((c) => cols.add(c));
    row.ours?.nonEmptyCols.forEach((c) => cols.add(c));
    row.theirs?.nonEmptyCols.forEach((c) => cols.add(c));
    if (cols.size === 0) return;

    for (const col of cols) {
      const baseValue = row.base?.values[col - 1] ?? null;
      const oursValue = row.ours?.values[col - 1] ?? null;
      const theirsValue = row.theirs?.values[col - 1] ?? null;
      const { status, mergedValue } = classifyThreeWayCell(baseValue, oursValue, theirsValue);

      if (status !== 'unchanged') {
        const addressRow =
          row.ours?.rowNumber ?? row.base?.rowNumber ?? row.theirs?.rowNumber ?? visualRowNumber;
        const colMeta = alignedColumns[col - 1];
        cells.push({
          address: makeAddress(col, addressRow),
          row: visualRowNumber,
          col,
          baseCol: colMeta?.baseCol ?? null,
          oursCol: colMeta?.oursCol ?? null,
          theirsCol: colMeta?.theirsCol ?? null,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
        hasExactDiff = true;
      }
    }
  });

  // 如果有差异列，为冻结行补齐这些列的内容（即使未变化），用于显示表头冻结行上下文
  if (headerCount > 0 && cells.length > 0) {
    const diffColumns = new Set<number>(cells.map((c) => c.col));
    if (diffColumns.size > 0) {
      const existing = new Set<string>(cells.map((c) => `${c.row}:${c.col}`));
      for (let r = 1; r <= headerCount; r += 1) {
        const baseRow = buildHeaderRowRecordAligned(baseWsForAlign, r, alignedColumns, metaKeyCol, 'base');
        const oursRow = buildHeaderRowRecordAligned(oursWs, r, alignedColumns, metaKeyCol, 'ours');
        const theirsRow = buildHeaderRowRecordAligned(theirsWs, r, alignedColumns, metaKeyCol, 'theirs');
        for (const col of diffColumns) {
          const key = `${r}:${col}`;
          if (existing.has(key)) continue;
          const baseValue = baseRow.values[col - 1] ?? null;
          const oursValue = oursRow.values[col - 1] ?? null;
          const theirsValue = theirsRow.values[col - 1] ?? null;
          const colMeta = alignedColumns[col - 1];
          cells.push({
            address: makeAddress(col, r),
            row: r,
            col,
            baseCol: colMeta?.baseCol ?? null,
            oursCol: colMeta?.oursCol ?? null,
            theirsCol: colMeta?.theirsCol ?? null,
            baseValue,
            oursValue,
            theirsValue,
            status: 'unchanged',
            mergedValue: baseValue,
          });
          existing.add(key);
        }
      }
    }
  }
  cells.sort((a, b) => a.row - b.row || a.col - b.col);

  return {
    sheetName: baseWs.name,
    cells,
    rowsMeta,
    hasExactDiff,
    columnsMeta,
    primaryKeyAlignedCol,
    primaryKeyOursCol,
    primaryKeySource,
  };
};

const convertMergeSheetToSimpleMerge = (sheet: MergeSheetData): MergeSheetData => ({
  ...sheet,
  cells: (sheet.cells ?? []).map((cell) => ({
    ...cell,
    baseCol: null,
    baseValue: null,
    status: cell.status === 'unchanged' ? 'unchanged' : 'conflict',
    mergedValue:
      cell.status === 'unchanged'
        ? cell.oursValue ?? cell.theirsValue ?? cell.mergedValue ?? null
        : cell.oursValue,
  })),
  rowsMeta: (sheet.rowsMeta ?? []).map((rowMeta) => ({
    ...rowMeta,
    baseRowNumber: null,
  })),
  columnsMeta: sheet.columnsMeta?.map((columnMeta) => ({
    ...columnMeta,
    baseCol: null,
  })),
});

// 简单缓存：同一次应用生命周期内重复读取同一个xlsx 时复用workbook，减少IO
const workbookCache = new Map<string, Workbook>();

const loadWorkbookCached = async (filePath: string): Promise<Workbook> => {
  patchExcelJsLegacyCommentCompat();
  const hit = workbookCache.get(filePath);
  if (hit) {
    appendDebugLog('main', 'workbook-cache:hit', { filePath: debugFileLabel(filePath) });
    return hit;
  }
  const startedAt = Date.now();
  appendDebugLog('main', 'workbook-cache:miss', { filePath: debugFileLabel(filePath) });
  const wb = new Workbook();
  await wb.xlsx.readFile(filePath);
  workbookCache.set(filePath, wb);
  appendDebugLog('main', 'workbook-cache:store', {
    filePath: debugFileLabel(filePath),
    durationMs: Date.now() - startedAt,
    sheetCount: wb.worksheets.length,
  });
  return wb;
};

const getWorksheetSafe = (wb: Workbook, sheetName?: string, sheetIndex?: number): any => {
  if (sheetName) {
    const byName = wb.getWorksheet(sheetName);
    if (byName) return byName;
  }
  if (typeof sheetIndex === 'number' && sheetIndex >= 0 && sheetIndex < wb.worksheets.length) {
    return wb.worksheets[sheetIndex];
  }
  return wb.worksheets[0];
};

const buildMergeSheetsForWorkbooks = async (
  basePath: string,
  oursPath: string,
  theirsPath: string,
  primaryKeyCol: number,
  frozenRowCount: number,
  rowSimilarityThreshold: number,
  compareMode: ThreeWayCompareMode,
) => {
  const startedAt = Date.now();
  appendDebugLog('main', 'buildMergeSheetsForWorkbooks:start', {
    compareMode,
    primaryKeyCol,
    frozenRowCount,
    rowSimilarityThreshold,
    basePath: debugFileLabel(basePath),
    oursPath: debugFileLabel(oursPath),
    theirsPath: debugFileLabel(theirsPath),
  });
  // 复用缓存，避免每次调取重新 diff 时重复从磁盘读取
  const effectiveBasePath = compareMode === 'simple-merge' ? oursPath : basePath;
  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(effectiveBasePath),
    loadWorkbookCached(oursPath),
    loadWorkbookCached(theirsPath),
  ]);

  const baseList = baseWb.worksheets;
  const oursList = oursWb.worksheets;
  const theirsList = theirsWb.worksheets;

  const baseByName = new Map<string, { ws: any; idx: number }>();
  baseList.forEach((ws, idx) => {
    if (!baseByName.has(ws.name)) baseByName.set(ws.name, { ws, idx });
  });
  const oursByName = new Map<string, { ws: any; idx: number }>();
  oursList.forEach((ws, idx) => {
    if (!oursByName.has(ws.name)) oursByName.set(ws.name, { ws, idx });
  });
  const theirsByName = new Map<string, { ws: any; idx: number }>();
  theirsList.forEach((ws, idx) => {
    if (!theirsByName.has(ws.name)) theirsByName.set(ws.name, { ws, idx });
  });

  // 规则：优先按同名工作表对齐；对剩余未匹配的工作表，再按索引对齐（第1 张对第1 张……）。
  const usedBaseIdx = new Set<number>();
  const usedOursIdx = new Set<number>();
  const usedTheirsIdx = new Set<number>();

  const mergeSheets: MergeSheetData[] = [];

  // 1) 同名匹配：以 base 的顺序为准
  for (let i = 0; i < baseList.length; i += 1) {
    const baseWs = baseList[i];
    const oursHit = oursByName.get(baseWs.name);
    const theirsHit = theirsByName.get(baseWs.name);
    if (!oursHit || !theirsHit) continue;

    usedBaseIdx.add(i);
    usedOursIdx.add(oursHit.idx);
    usedTheirsIdx.add(theirsHit.idx);

    mergeSheets.push(
      compareMode === 'simple-merge'
        ? convertMergeSheetToSimpleMerge(
            buildMergeSheetWithRowAlign(
              oursHit.ws,
              oursHit.ws,
              theirsHit.ws,
              primaryKeyCol,
              frozenRowCount,
              rowSimilarityThreshold,
              'diff',
            ),
          )
        : buildMergeSheetWithRowAlign(
            baseWs,
            oursHit.ws,
            theirsHit.ws,
            primaryKeyCol,
            frozenRowCount,
            rowSimilarityThreshold,
            compareMode,
          ),
    );
  }

  // 2) 索引兜底：仅对“同一 idx 在三边都没被用过”的位置做对齐
  const count = Math.min(baseList.length, oursList.length, theirsList.length);
  for (let idx = 0; idx < count; idx += 1) {
    if (usedBaseIdx.has(idx) || usedOursIdx.has(idx) || usedTheirsIdx.has(idx)) continue;
    usedBaseIdx.add(idx);
    usedOursIdx.add(idx);
    usedTheirsIdx.add(idx);
    mergeSheets.push(
      compareMode === 'simple-merge'
        ? convertMergeSheetToSimpleMerge(
            buildMergeSheetWithRowAlign(
              oursList[idx],
              oursList[idx],
              theirsList[idx],
              primaryKeyCol,
              frozenRowCount,
              rowSimilarityThreshold,
              'diff',
            ),
          )
        : buildMergeSheetWithRowAlign(
            baseList[idx],
            oursList[idx],
            theirsList[idx],
            primaryKeyCol,
            frozenRowCount,
            rowSimilarityThreshold,
            compareMode,
          ),
    );
  }

  appendDebugLog('main', 'buildMergeSheetsForWorkbooks:end', {
    compareMode,
    durationMs: Date.now() - startedAt,
    sheetCount: mergeSheets.length,
    diffCellCount: mergeSheets.reduce((sum, sheet) => sum + (sheet.cells?.length ?? 0), 0),
  });
  return { basePath, oursPath, theirsPath, mergeSheets };
};

const normalizeThreeWayResult = (
  basePath: string,
  oursPath: string,
  theirsPath: string,
  compareMode: ThreeWayCompareMode,
  mergeSheets: MergeSheetData[],
) => {
  const emptySheet: MergeSheetData = { sheetName: '', cells: [], rowsMeta: [] };
  return {
    basePath,
    oursPath,
    theirsPath,
    compareMode,
    sheet: mergeSheets[0] ?? emptySheet,
    sheets: mergeSheets,
  };
};

// IPC types
interface SheetCell {
  address: string; // e.g. "A1"
  row: number;
  col: number;
  value: string | number | null;
}

type RowStatus = 'unchanged' | 'added' | 'deleted' | 'modified' | 'ambiguous';

interface MergeRowMeta {
  /** 视觉行号（diff/merge 视图中的 1-based 行号）*/
  visualRowNumber: number;
  /** 如果启用了主键列，这里记录主键（normalize 后） */
  key?: string | null;
  /** 三方文件中各自对应的原始行号（-based）；不存在则为null */
  baseRowNumber: number | null;
  oursRowNumber: number | null;
  theirsRowNumber: number | null;
  /** 行相似度（相对base，范围0-1）*/
  oursSimilarity?: number | null;
  theirsSimilarity?: number | null;
  /** 该视觉行在对度side 相对 base 的状态*/
  oursStatus: RowStatus;
  theirsStatus: RowStatus;
}

interface SheetData {
  sheetName: string;
  rows: SheetCell[][];
}

interface MergeCell {
  address: string;
  row: number;
  col: number;
  baseCol?: number | null;
  oursCol?: number | null;
  theirsCol?: number | null;
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  status: 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
  mergedValue: string | number | null;
}
interface MergeColumnMeta {
  col: number; // aligned column index (1-based)
  baseCol: number | null;
  oursCol: number | null;
  theirsCol: number | null;
}

type PrimaryKeySource = 'manual' | 'auto-implicit' | 'auto-header' | 'auto-weak' | 'none';

interface MergeSheetData {
  sheetName: string;
  cells: MergeCell[];
  rowsMeta?: MergeRowMeta[];
  hasExactDiff?: boolean;
  columnsMeta?: MergeColumnMeta[];
  primaryKeyAlignedCol?: number | null;
  primaryKeyOursCol?: number | null;
  primaryKeySource?: PrimaryKeySource;
}

interface SaveMergeCellInput {
  address: string;
  value: string | number | null;
}
interface SaveMergeRowOp {
  sheetName: string;
  action: 'insert' | 'delete' | 'skip-insert';
  targetRowNumber: number; // 1-based in template (ours)
  values?: (string | number | null)[];
  visualRowNumber?: number;
}

interface SaveMergeColOp {
  sheetName: string;
  action: 'insert' | 'delete';
  targetColNumber: number; // 1-based in template (ours)
  alignedColNumber?: number; // 1-based aligned column index
  values?: (string | number | null)[];
  source?: 'theirs' | 'base' | 'ours';
}

interface SaveMergeRequest {
  templatePath: string;
  cells: SaveMergeCellInput[];
  rowOps?: SaveMergeRowOp[];
  colOps?: SaveMergeColOp[];
  basePath?: string;
  oursPath?: string;
  theirsPath?: string;
}

interface SaveMergeResponse {
  success: boolean;
  filePath?: string;
  cancelled?: boolean;
  errorMessage?: string;
}

let currentFilePath: string | null = null;

const readWorkbookOpenResult = async (filePath: string) => {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filePath);

  const buildSheetData = (worksheet: Worksheet): SheetData => {
    const rows: SheetCell[][] = [];

    const getSimpleValue = (raw: CellValue): string | number | null => {
      if (raw === null || raw === undefined) return null;

      if (raw instanceof Date) {
        return raw.toISOString();
      }

      if (typeof raw === 'object' && Array.isArray((raw as any).richText)) {
        const parts = (raw as any).richText
          .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
          .join('');
        return parts;
      }

      if (typeof raw === 'object' && raw && 'text' in (raw as any)) {
        const t = (raw as any).text;
        if (t === null || t === undefined) return null;
        return typeof t === 'string' || typeof t === 'number' ? (t as any) : String(t);
      }

      if (typeof raw === 'object' && raw && 'result' in (raw as any)) {
        const r = (raw as any).result;
        if (r === null || r === undefined) return null;
        if (typeof r === 'string' || typeof r === 'number') return r;
        if (r instanceof Date) return r.toISOString();
        return String(r);
      }

      if (typeof raw === 'string' || typeof raw === 'number') {
        return raw;
      }

      if (typeof raw === 'object') {
        try {
          return JSON.stringify(raw);
        } catch {
          return String(raw);
        }
      }

      return String(raw);
    };

    const maxRow = getWorksheetRowUpperBound(worksheet);
    const maxCol =
      (worksheet as any).actualColumnCount && (worksheet as any).actualColumnCount > 0
        ? (worksheet as any).actualColumnCount
        : worksheet.columnCount;

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      const rowCells: SheetCell[] = [];
      const row = worksheet.getRow(rowNumber);
      for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
        const cell = row.getCell(colNumber);
        const value = getSimpleValue(cell.value as any);
        rowCells.push({
          address: cell.address,
          row: rowNumber,
          col: colNumber,
          value,
        });
      }
      rows.push(rowCells);
    }

    return {
      sheetName: worksheet.name,
      rows,
    };
  };

  const sheets: SheetData[] = workbook.worksheets.map((ws) => buildSheetData(ws));
  return { filePath, sheet: sheets[0], sheets };
};

/**
 * 处理渲染进程请求：选择并打开一个Excel 文件。
 *
 * 返回：文件路径+ 所有工作表的二维单元格数据（仅包含“值”）：
 * 用于单文件查看编辑模式。
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
  return readWorkbookOpenResult(filePath);
});

interface FolderExcelFileInfo {
  relativePath: string;
  absolutePath: string;
  sizeBytes: number;
  modifiedAtMs: number;
}
const isExcelLikeFile = (fileName: string): boolean => {
  const ext = path.extname(fileName).toLowerCase();
  return ext === '.xlsx' || ext === '.xlsm' || ext === '.xls';
};
const walkExcelFiles = (rootDir: string, currentDir: string, bucket: FolderExcelFileInfo[]): void => {
  const entries = fs.readdirSync(currentDir, { withFileTypes: true });
  entries.forEach((entry) => {
    const absolutePath = path.join(currentDir, entry.name);
    if (entry.isDirectory()) {
      walkExcelFiles(rootDir, absolutePath, bucket);
      return;
    }
    if (!entry.isFile() || !isExcelLikeFile(entry.name)) return;
    let stat: fs.Stats;
    try {
      stat = fs.statSync(absolutePath);
    } catch {
      return;
    }
    const relativePath = path.relative(rootDir, absolutePath).split(path.sep).join('/');
    bucket.push({
      relativePath,
      absolutePath,
      sizeBytes: stat.size,
      modifiedAtMs: stat.mtimeMs,
    });
  });
};

type WorkspaceTabMenuKind = 'folder' | 'diff' | 'merge';
const emitWorkspaceTabCommand = (kind: WorkspaceTabMenuKind) => {
  const targetWindow = BrowserWindow.getFocusedWindow() ?? mainWindow;
  targetWindow?.webContents.send('workspace:newTab', { kind });
};
const configureApplicationMenu = () => {
  const fileSubmenu: MenuItemConstructorOptions[] = [
    {
      label: 'Excel 文件夹比较',
      click: () => emitWorkspaceTabCommand('folder'),
    },
    {
      label: 'Excel 比较',
      click: () => emitWorkspaceTabCommand('diff'),
    },
    {
      label: 'Merge 模式',
      click: () => emitWorkspaceTabCommand('merge'),
    },
    { type: 'separator' },
    process.platform === 'darwin' ? { role: 'close' } : { role: 'quit' },
  ];
  const template: MenuItemConstructorOptions[] = [
    { label: 'File', submenu: fileSubmenu },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' },
    { role: 'help', submenu: [{ role: 'about' }] },
  ];
  if (process.platform === 'darwin') {
    template.unshift({ role: 'appMenu' });
  }
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
};
ipcMain.handle('excel:pickFolder', async (): Promise<string | null> => {
  if (!mainWindow) return null;
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: '选择 Excel 文件夹',
    properties: ['openDirectory'],
  });
  if (canceled || filePaths.length === 0) return null;
  return filePaths[0];
});
ipcMain.handle('excel:listExcelFilesInFolder', async (_event, folderPath: string): Promise<FolderExcelFileInfo[]> => {
  if (!folderPath) return [];
  const normalized = path.resolve(folderPath);
  let stat: fs.Stats;
  try {
    stat = fs.statSync(normalized);
  } catch {
    return [];
  }
  if (!stat.isDirectory()) return [];
  const files: FolderExcelFileInfo[] = [];
  walkExcelFiles(normalized, normalized, files);
  return files.sort((a, b) => a.relativePath.localeCompare(b.relativePath));
});

ipcMain.handle('excel:loadWorkbook', async (_event, filePath: string) => {
  if (!filePath) return null;
  return readWorkbookOpenResult(filePath);
});

interface CellChange {
  address: string;
  newValue: string | number | null;
}
interface GetSheetDataRequest {
  path: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
}

/**
 * 将单文件编辑模式下用户修改过的单元格写回原始 Excel 文件。
 *
 * 只修改单元格的value，不动样式公式等格式信息。
 */
ipcMain.handle('excel:saveChanges', async (_event, req: CellChange[] | { changes: CellChange[]; sheetName?: string; sheetIndex?: number; filePath?: string; rowOps?: SaveMergeRowOp[] }) => {
  const targetFilePath = !Array.isArray(req) ? req?.filePath ?? currentFilePath : currentFilePath;
  if (!targetFilePath) {
    throw new Error('No Excel file is currently loaded');
  }
  const changes: CellChange[] = Array.isArray(req) ? req : (req?.changes ?? []);
  const sheetName = !Array.isArray(req) ? req?.sheetName : undefined;
  const sheetIndex = !Array.isArray(req) ? req?.sheetIndex : undefined;
  const rowOps = !Array.isArray(req) ? req?.rowOps ?? [] : [];

  const workbook = new Workbook();
  await workbook.xlsx.readFile(targetFilePath);
  let worksheet = sheetName ? workbook.getWorksheet(sheetName) ?? undefined : undefined;
  if (!worksheet && typeof sheetIndex === 'number' && sheetIndex >= 0) {
    worksheet = workbook.worksheets[sheetIndex];
  }
  if (!worksheet) worksheet = workbook.worksheets[0];

  if (rowOps.length > 0) {
    const rowOpsBySheet = new Map<string, SaveMergeRowOp[]>();
    rowOps.forEach((op) => {
      const key = op.sheetName || sheetName || '';
      if (!rowOpsBySheet.has(key)) rowOpsBySheet.set(key, []);
      rowOpsBySheet.get(key)!.push(op);
    });
    rowOpsBySheet.forEach((ops, targetSheetName) => {
      let ws = targetSheetName ? workbook.getWorksheet(targetSheetName) ?? undefined : undefined;
      if (!ws && sheetName) ws = workbook.getWorksheet(sheetName) ?? undefined;
      if (!ws && typeof sheetIndex === 'number' && sheetIndex >= 0) {
        ws = workbook.worksheets[sheetIndex];
      }
      if (!ws) ws = workbook.worksheets[0];
      const sorted = ops
        .slice()
        .sort((a, b) => a.targetRowNumber - b.targetRowNumber || (a.visualRowNumber ?? 0) - (b.visualRowNumber ?? 0));
      for (const op of sorted) {
        const rowNumber = Math.max(1, Math.floor(op.targetRowNumber));
        if (op.action === 'insert') {
          const maxCol = Math.max(
            ws?.actualColumnCount ?? ws?.columnCount ?? 0,
            op.values?.length ?? 0,
          );
          const values: (string | number | null)[] = [];
          for (let i = 0; i < maxCol; i += 1) {
            values.push(op.values && i < op.values.length ? op.values[i] ?? null : null);
          }
          ws.spliceRows(rowNumber, 0, values);
        } else if (op.action === 'delete') {
          ws.spliceRows(rowNumber, 1);
        } else if (op.action === 'skip-insert') {
          continue;
        }
      }
    });
    worksheet = sheetName ? workbook.getWorksheet(sheetName) ?? worksheet : worksheet;
    if (!worksheet && typeof sheetIndex === 'number' && sheetIndex >= 0) {
      worksheet = workbook.worksheets[sheetIndex];
    }
    if (!worksheet) worksheet = workbook.worksheets[0];
  }

  for (const change of changes) {
    const cell = worksheet.getCell(change.address);
    cell.value = change.newValue as any; // only change value, keep formatting/styles
  }

  await workbook.xlsx.writeFile(targetFilePath);
  // invalidate cache to avoid stale reads
  if (workbookCache.has(targetFilePath)) {
    workbookCache.delete(targetFilePath);
  }

  return { success: true };
});

// 读取指定文件的指定工作表（用了merge 模式下显示全表）
ipcMain.handle('excel:getSheetData', async (_event, req: GetSheetDataRequest): Promise<SheetData | null> => {
  if (!req || !req.path) return null;
  const wb = await loadWorkbookCached(req.path);
  const ws = getWorksheetSafe(wb, req.sheetName, req.sheetIndex);
  if (!ws) return null;

  const maxRow = getWorksheetRowUpperBound(ws);
  const maxCol =
    (ws as any).actualColumnCount && (ws as any).actualColumnCount > 0
      ? (ws as any).actualColumnCount
      : ws.columnCount;

  const rows: SheetCell[][] = [];
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const rowCells: SheetCell[] = [];
    const row = ws.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
      const cell = row.getCell(colNumber);
      const value = getSimpleValueForMerge(cell?.value);
      rowCells.push({
        address: cell.address,
        row: rowNumber,
        col: colNumber,
        value,
      });
    }
    rows.push(rowCells);
  }

  return { sheetName: ws.name, rows };
});

// 保存三方 merge 结果到新的Excel 文件，仅修改值，不改格式
//
// 在git/Fork merge 模式下：
//   - 如果提供了MERGED 参数，则结果写回 MERGED：
//   - 否则回退到覆盖ours：
// 在diff 模式下：
//   - 直接覆盖 ours（LOCAL）。
// 交互式模式下：
//   - 弹出保存对话框，由用户选择目标路径。
const saveMergeResultInternal = async (
  req: SaveMergeRequest,
  options?: { targetPathOverride?: string; skipMainWindowRequirement?: boolean },
): Promise<SaveMergeResponse> => {
  if (!options?.skipMainWindowRequirement && !mainWindow) {
    throw new Error('Main window is not available');
  }

  try {
    const { templatePath, cells, rowOps, colOps } = req as {
      templatePath: string;
      cells: { sheetName: string; address: string; value: string | number | null }[];
      rowOps?: SaveMergeRowOp[];
      colOps?: SaveMergeColOp[];
    };
    let targetPath: string | undefined = options?.targetPathOverride;
    const unresolvedCliMergedPath =
      isCliMergeMode(cliThreeWayArgs?.mode) && cliThreeWayArgs?.mergedPathRaw && !cliThreeWayArgs.mergedPath
        ? cliThreeWayArgs.mergedPathRaw
        : null;

    if (!targetPath && cliThreeWayArgs && isCliMergeMode(cliThreeWayArgs.mode)) {
      // git / Fork merge 模式：
      //   - 如果 MERGED 能解析成绝对路径，则优先写回 MERGED；
      //   - 如果根本没传 MERGED，则回退覆盖 ours；
      //   - 如果传了相对 MERGED 但无法可靠解析，则改为让用户确认保存位置，避免误写到便携版临时目录。
      if (cliThreeWayArgs.mergedPath) {
        targetPath = cliThreeWayArgs.mergedPath;
      } else if (!cliThreeWayArgs.mergedPathRaw) {
        targetPath = cliThreeWayArgs.oursPath;
      }
    } else if (!targetPath && cliThreeWayArgs && cliThreeWayArgs.mode === 'diff') {
      targetPath = cliThreeWayArgs.oursPath;
    }

    if (!targetPath) {
      const windowForSave = mainWindow;
      if (!windowForSave) {
        throw new Error('Main window is not available');
      }
      const unresolvedMergedLeaf =
        unresolvedCliMergedPath && unresolvedCliMergedPath.trim()
          ? path.basename(unresolvedCliMergedPath)
          : '';
      const { canceled, filePath } = await dialog.showSaveDialog(windowForSave, {
        title: unresolvedCliMergedPath ? '无法自动定位 Git 的 MERGED 文件，请确认保存位置' : '保存合并后的 Excel',
        defaultPath: unresolvedMergedLeaf || templatePath,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
      });

      if (canceled || !filePath) {
        return { success: false, cancelled: true };
      }
      targetPath = filePath;
    }

    const workbook = new Workbook();
    await workbook.xlsx.readFile(templatePath);

    // IMPORTANT: 必须先执行列/行操作，再修改单元格
    // 因为删行操作会改变索引，如果先修改单元格，地址会错乱
    
    const colOpsBySheet = new Map<string, SaveMergeColOp[]>();
    const rowOpsBySheet = new Map<string, SaveMergeRowOp[]>();

    // 1. 先执行列操作
    if (colOps && colOps.length > 0) {
      colOps.forEach((op) => {
        const key = op.sheetName || '';
        if (!colOpsBySheet.has(key)) colOpsBySheet.set(key, []);
        colOpsBySheet.get(key)!.push(op);
      });
      colOpsBySheet.forEach((ops, sheetName) => {
        const ws = workbook.getWorksheet(sheetName) ?? workbook.worksheets[0];
        const sorted = ops.slice().sort((a, b) => {
          const va = a.alignedColNumber ?? 0;
          const vb = b.alignedColNumber ?? 0;
          if (va !== vb) return va - vb;
          return a.targetColNumber - b.targetColNumber;
        });
        // Process deletes first (sorted by col descending to maintain positions)
        const deletes = sorted.filter(op => op.action === 'delete').sort((a, b) => b.targetColNumber - a.targetColNumber);
        for (const op of deletes) {
          const colNumber = Math.max(1, Math.floor(op.targetColNumber));
          if (typeof (ws as any).spliceColumns === 'function') {
            (ws as any).spliceColumns(colNumber, 1);
          } else {
            // fallback: manual delete by shifting cells left
            const maxRow = getWorksheetRowUpperBound(ws);
            const maxCol = ws?.actualColumnCount ?? ws?.columnCount ?? 0;
            for (let r = 1; r <= maxRow; r += 1) {
              for (let c = colNumber; c < maxCol; c += 1) {
                const from = ws.getRow(r).getCell(c + 1);
                const to = ws.getRow(r).getCell(c);
                to.value = from.value as any;
              }
              // Clear last column
              ws.getRow(r).getCell(maxCol).value = null;
            }
          }
        }
        // Then process inserts (sorted by aligned col ascending)
        const inserts = sorted.filter(op => op.action === 'insert');
        let offset = 0;
        for (const op of inserts) {
          // 注意：targetColNumber 是原始 ours 列号，需要去掉前面已执行的 delete 偏移
          let baseCol = Math.max(1, Math.floor(op.targetColNumber));
          for (const delOp of deletes) {
            const delCol = Math.max(1, Math.floor(delOp.targetColNumber));
            if (baseCol > delCol) baseCol -= 1;
          }
          const colNumber = baseCol + offset;
          const maxRow = Math.max(getWorksheetRowUpperBound(ws), op.values?.length ?? 0);
          const values: (string | number | null)[] = [];
          for (let i = 0; i < maxRow; i += 1) {
            values.push(op.values && i < op.values.length ? op.values[i] ?? null : null);
          }
          if (typeof (ws as any).spliceColumns === 'function') {
            (ws as any).spliceColumns(colNumber, 0, values);
          } else {
            // fallback: manual insert by shifting cells (rare)
            for (let r = maxRow; r >= 1; r -= 1) {
              for (let c = (ws?.actualColumnCount ?? ws?.columnCount ?? 0); c >= colNumber; c -= 1) {
                const from = ws.getRow(r).getCell(c);
                const to = ws.getRow(r).getCell(c + 1);
                to.value = from.value as any;
              }
              const cell = ws.getRow(r).getCell(colNumber);
              cell.value = values[r - 1] ?? null;
            }
          }
          offset += 1;
        }
      });
    }
    // 2. 再执行行操作
    if (rowOps && rowOps.length > 0) {
      rowOps.forEach((op) => {
        const key = op.sheetName || '';
        if (!rowOpsBySheet.has(key)) rowOpsBySheet.set(key, []);
        rowOpsBySheet.get(key)!.push(op);
      });
      rowOpsBySheet.forEach((ops, sheetName) => {
        const ws = workbook.getWorksheet(sheetName) ?? workbook.worksheets[0];
        const sorted = ops.slice().sort((a, b) => {
          const va = a.visualRowNumber ?? 0;
          const vb = b.visualRowNumber ?? 0;
          if (va !== vb) return va - vb;
          return a.targetRowNumber - b.targetRowNumber;
        });
        let offset = 0;
        for (const op of sorted) {
          const baseRow = Math.max(1, Math.floor(op.targetRowNumber));
          const rowNumber = baseRow + offset;
          if (op.action === 'insert') {
            const maxCol = Math.max(
              ws?.actualColumnCount ?? ws?.columnCount ?? 0,
              op.values?.length ?? 0,
            );
            const values: (string | number | null)[] = [];
            for (let i = 0; i < maxCol; i += 1) {
              values.push(op.values && i < op.values.length ? op.values[i] ?? null : null);
            }
            ws.spliceRows(rowNumber, 0, values);
            offset += 1;
          } else if (op.action === 'delete') {
            ws.spliceRows(rowNumber, 1);
            offset -= 1;
          } else if (op.action === 'skip-insert') {
            continue;
          }
        }
      });
    }

    const colLabelToNumber = (label: string): number => {
      const s = label.toUpperCase();
      let n = 0;
      for (let i = 0; i < s.length; i += 1) {
        const code = s.charCodeAt(i);
        if (code < 65 || code > 90) return NaN;
        n = n * 26 + (code - 64);
      }
      return n;
    };
    const parseAddress = (address: string): { col: number; row: number } | null => {
      const m = /^([A-Z]+)(\d+)$/i.exec(address);
      if (!m) return null;
      const col = colLabelToNumber(m[1]);
      const row = Number(m[2]);
      if (!Number.isFinite(col) || !Number.isFinite(row)) return null;
      return { col, row };
    };
    const buildRowMapper = (ops: SaveMergeRowOp[]) => {
      const sorted = ops.slice().sort((a, b) => {
        const va = a.visualRowNumber ?? 0;
        const vb = b.visualRowNumber ?? 0;
        if (va !== vb) return va - vb;
        return a.targetRowNumber - b.targetRowNumber;
      });
      return (row: number): number | null => {
        let r = row;
        let offset = 0;
        for (const op of sorted) {
          const baseRow = Math.max(1, Math.floor(op.targetRowNumber));
          const rowNumber = baseRow + offset;
          if (op.action === 'insert') {
            if (r >= rowNumber) r += 1;
            offset += 1;
          } else if (op.action === 'delete') {
            if (r === rowNumber) return null;
            if (r > rowNumber) r -= 1;
            offset -= 1;
          }
        }
        return r;
      };
    };
    const buildColMapper = (ops: SaveMergeColOp[]) => {
      const sorted = ops.slice().sort((a, b) => {
        const va = a.alignedColNumber ?? 0;
        const vb = b.alignedColNumber ?? 0;
        if (va !== vb) return va - vb;
        return a.targetColNumber - b.targetColNumber;
      });
      const deletes = sorted
        .filter((op) => op.action === 'delete')
        .sort((a, b) => b.targetColNumber - a.targetColNumber);
      const inserts = sorted.filter((op) => op.action === 'insert');
      return (col: number): number | null => {
        let c = col;
        for (const op of deletes) {
          const colNumber = Math.max(1, Math.floor(op.targetColNumber));
          if (c === colNumber) return null;
          if (c > colNumber) c -= 1;
        }
        let offset = 0;
        for (const op of inserts) {
          // 注意：insert 的 targetColNumber 需从已执行 delete 偏移后计算
          let adjustedBase = Math.max(1, Math.floor(op.targetColNumber));
          for (const delOp of deletes) {
            const delCol = Math.max(1, Math.floor(delOp.targetColNumber));
            if (adjustedBase > delCol) adjustedBase -= 1;
          }
          const insertAt = adjustedBase + offset;
          if (c >= insertAt) c += 1;
          offset += 1;
        }
        return c;
      };
    };
    const rowMapperCache = new Map<string, (row: number) => number | null>();
    const colMapperCache = new Map<string, (col: number) => number | null>();
    const getRowMapper = (sheetKey: string) => {
      if (!rowMapperCache.has(sheetKey)) {
        rowMapperCache.set(sheetKey, buildRowMapper(rowOpsBySheet.get(sheetKey) ?? []));
      }
      return rowMapperCache.get(sheetKey)!;
    };
    const getColMapper = (sheetKey: string) => {
      if (!colMapperCache.has(sheetKey)) {
        colMapperCache.set(sheetKey, buildColMapper(colOpsBySheet.get(sheetKey) ?? []));
      }
      return colMapperCache.get(sheetKey)!;
    };

    // 3. 最后修改单元格值（此时删行索引已经稳定）
    for (const cellInfo of cells) {
      const sheetKey = cellInfo.sheetName || '';
      const ws = workbook.getWorksheet(cellInfo.sheetName) ?? workbook.worksheets[0];
      const parsed = parseAddress(cellInfo.address);
      if (!parsed) continue;
      const newCol = getColMapper(sheetKey)(parsed.col);
      if (newCol == null) continue;
      const newRow = getRowMapper(sheetKey)(parsed.row);
      if (newRow == null) continue;
      const newAddress = makeAddress(newCol, newRow);
      const cell = ws.getCell(newAddress);
      cell.value = cellInfo.value as any;
    }

    normalizeSharedFormulas(workbook);
    const targetDir = path.dirname(targetPath);
    if (targetDir && !fs.existsSync(targetDir)) {
      fs.mkdirSync(targetDir, { recursive: true });
    }
    await workbook.xlsx.writeFile(targetPath);
    // invalidate cache to avoid stale reads
    if (targetPath && workbookCache.has(targetPath)) {
      workbookCache.delete(targetPath);
    }
    if (templatePath && templatePath !== targetPath && workbookCache.has(templatePath)) {
      workbookCache.delete(templatePath);
    }

    // 如果是通过 git/Fork 的merge 模式启动，并且有明确的目标文件，尝试自动执行一次git add
    if (cliThreeWayArgs && isCliMergeMode(cliThreeWayArgs.mode) && targetPath) {
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
};

ipcMain.handle('excel:saveMergeResult', async (_event, req: SaveMergeRequest): Promise<SaveMergeResponse> => {
  return saveMergeResultInternal(req);
});
ipcMain.on('excel:debugLog', (_event, entry: DebugLogEntry) => {
  appendDebugLog(entry?.source ?? 'renderer', entry?.event ?? 'unknown', entry?.details);
});
ipcMain.handle('excel:getDebugLogPath', async () => getDebugLogPath());

// 三方 diff：base / ours / theirs，只比较单元格值，忽略格式
//
// 返回给渲染进程的数据是：
//   - base / ours / theirs 的文件路径；
//   - 每个工作表的三方单元格值+ 差异状态（unchanged / conflict 等）。
ipcMain.handle('excel:openThreeWay', async () => {
  if (!mainWindow) return null;
  const primaryKeyCol = 1;
  const frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS;
  const rowSimilarityThreshold = DEFAULT_ROW_SIMILARITY_THRESHOLD;

  if (cliThreeWayArgs) {
    const { basePath, oursPath, theirsPath, mode } = cliThreeWayArgs;
    const compareMode = normalizeThreeWayCompareMode(mode);
    const { mergeSheets } = await buildMergeSheetsForWorkbooks(
      basePath,
      oursPath,
      theirsPath,
      primaryKeyCol,
      frozenRowCount,
      rowSimilarityThreshold,
      compareMode,
    );
    return normalizeThreeWayResult(basePath, oursPath, theirsPath, compareMode, mergeSheets);
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

  const { mergeSheets } = await buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    primaryKeyCol,
    frozenRowCount,
    rowSimilarityThreshold,
    'merge',
  );

  return normalizeThreeWayResult(basePath, oursPath, theirsPath, 'merge', mergeSheets);
});
interface ThreeWayDiffRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  compareMode?: ThreeWayCompareMode;
  primaryKeyCol: number; // 1-based manual key; -1 means auto-detect; 0 means force no primary key
  frozenRowCount?: number; // header rows compared by coordinates
  rowSimilarityThreshold?: number; // 0-1
  debugRequestId?: string;
}

ipcMain.handle('excel:computeThreeWayDiff', async (_event, req: ThreeWayDiffRequest) => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath) return null;
  const requestId = req.debugRequestId ?? `compute-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const startedAt = Date.now();
  const primaryKeyCol =
    typeof req.primaryKeyCol === 'number' && !Number.isNaN(req.primaryKeyCol) ? Math.floor(req.primaryKeyCol) : -1;
  const frozenRowCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const rowSimilarityThreshold =
    typeof req.rowSimilarityThreshold === 'number' && !Number.isNaN(req.rowSimilarityThreshold)
      ? Math.min(1, Math.max(0, req.rowSimilarityThreshold))
      : DEFAULT_ROW_SIMILARITY_THRESHOLD;
  const compareMode = normalizeThreeWayCompareMode(req.compareMode);
  appendDebugLog('main', 'computeThreeWayDiff:start', {
    requestId,
    compareMode,
    primaryKeyCol,
    frozenRowCount,
    rowSimilarityThreshold,
    basePath: debugFileLabel(req.basePath),
    oursPath: debugFileLabel(req.oursPath),
    theirsPath: debugFileLabel(req.theirsPath),
  });
  try {
    const { mergeSheets } = await buildMergeSheetsForWorkbooks(
      req.basePath,
      req.oursPath,
      req.theirsPath,
      primaryKeyCol,
      frozenRowCount,
      rowSimilarityThreshold,
      compareMode,
    );
    appendDebugLog('main', 'computeThreeWayDiff:end', {
      requestId,
      durationMs: Date.now() - startedAt,
      sheetCount: mergeSheets.length,
      diffCellCount: mergeSheets.reduce((sum, sheet) => sum + (sheet.cells?.length ?? 0), 0),
    });
    return normalizeThreeWayResult(req.basePath, req.oursPath, req.theirsPath, compareMode, mergeSheets);
  } catch (error: any) {
    appendDebugLog('main', 'computeThreeWayDiff:error', {
      requestId,
      durationMs: Date.now() - startedAt,
      message: error?.message ?? String(error),
      stack: error?.stack,
    });
    throw error;
  }
});

// 将CLI three-way 信息暴露给渲染进程，便于自动加载
ipcMain.handle('excel:getCliThreeWayInfo', async () => {
  if (!cliThreeWayArgs) return null;
  return cliThreeWayArgs;
});

// 读取三方文件的“某一行”数据，用于底部行级对比视图
interface ThreeWayRowRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  compareMode?: ThreeWayCompareMode;
  sheetName?: string;
  sheetIndex?: number; // 0-based
  frozenRowCount?: number;
  rowNumber?: number; // 1-based fallback for all sides
  baseRowNumber?: number | null;
  oursRowNumber?: number | null;
  theirsRowNumber?: number | null;
  debugRequestId?: string;
}

interface ThreeWayRowResult {
  sheetName: string;
  rowNumber?: number;
  baseRowNumber: number | null;
  oursRowNumber: number | null;
  theirsRowNumber: number | null;
  colCount: number;
  base: (string | number | null)[];
  ours: (string | number | null)[];
  theirs: (string | number | null)[];
}
interface ThreeWayRowsRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  compareMode?: ThreeWayCompareMode;
  sheetName?: string;
  sheetIndex?: number; // 0-based
  frozenRowCount?: number;
  debugRequestId?: string;
  rows: Array<{
    rowNumber?: number;
    baseRowNumber?: number | null;
    oursRowNumber?: number | null;
    theirsRowNumber?: number | null;
  }>;
}
interface ThreeWayRowsResult {
  sheetName: string;
  colCount: number;
  rows: ThreeWayRowResult[];
}

const normalizeSharedFormulas = (workbook: Workbook) => {
  workbook.worksheets.forEach((ws) => {
    ws.eachRow({ includeEmpty: true }, (row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        const v: any = cell.value as any;
        if (!v || typeof v !== 'object') return;
        const isShared = v.sharedFormula || v.shareType === 'shared';
        if (!isShared) return;
        const model: any = (cell as any).model || {};
        const formula = model.formula || v.formula;
        const result = model.result !== undefined ? model.result : v.result;
        if (formula) {
          cell.value = { formula, result } as any;
          return;
        }
        if (result !== undefined) {
          cell.value = result as any;
          return;
        }
        cell.value = null as any;
      });
    });
  });
};

const resolveRequestedRowNumber = (
  value: number | null | undefined,
  fallbackRow: number | null,
): number | null => {
  if (typeof value === 'number' && !Number.isNaN(value)) {
    return Math.max(1, Math.floor(value));
  }
  if (value === null) {
    return null;
  }
  return fallbackRow;
};


ipcMain.handle('excel:getThreeWayRow', async (_event, req: ThreeWayRowRequest): Promise<ThreeWayRowResult | null> => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath) return null;
  const requestId = req.debugRequestId ?? `row-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const startedAt = Date.now();
  const fallbackRow =
    typeof req.rowNumber === 'number' && !Number.isNaN(req.rowNumber)
      ? Math.max(1, Math.floor(req.rowNumber))
      : null;
  const baseRowNumber = resolveRequestedRowNumber(req.baseRowNumber, fallbackRow);
  const oursRowNumber = resolveRequestedRowNumber(req.oursRowNumber, fallbackRow);
  const theirsRowNumber = resolveRequestedRowNumber(req.theirsRowNumber, fallbackRow);

  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(req.basePath),
    loadWorkbookCached(req.oursPath),
    loadWorkbookCached(req.theirsPath),
  ]);

  const baseWs = getWorksheetSafe(baseWb, req.sheetName, req.sheetIndex);
  const oursWs = getWorksheetSafe(oursWb, req.sheetName, req.sheetIndex);
  const theirsWs = getWorksheetSafe(theirsWb, req.sheetName, req.sheetIndex);

  const resolvedSheetName = baseWs?.name ?? req.sheetName ?? '';
  const headerCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const resolvedHeaderCount = resolveStructuredHeaderRowCount([baseWs, oursWs, theirsWs], headerCount);
  const compareMode = normalizeThreeWayCompareMode(req.compareMode);
  const runtimeConfig = getThreeWayRuntimeConfig(compareMode);
  const baseWsForAlign = runtimeConfig.alignBaseSide === 'ours' ? oursWs : baseWs;
  const alignedColumns = buildAlignedColumns(baseWsForAlign, oursWs, theirsWs, resolvedHeaderCount);
  const colCount = alignedColumns.length;
  appendDebugLog('main', 'getThreeWayRow:start', {
    requestId,
    compareMode,
    sheetName: resolvedSheetName,
    fallbackRow,
    baseRowNumber,
    oursRowNumber,
    theirsRowNumber,
    frozenRowCount: resolvedHeaderCount,
  });

  const readRowAligned = (
    ws: any,
    rowNum: number | null,
    side: 'base' | 'ours' | 'theirs',
  ): (string | number | null)[] => {
    const arr: (string | number | null)[] = [];
    if (!rowNum) {
      for (let col = 1; col <= colCount; col += 1) arr.push(null);
      return arr;
    }
    const row = ws.getRow(rowNum);
    for (let i = 0; i < alignedColumns.length; i += 1) {
      const meta = alignedColumns[i];
      const colNumber =
        side === 'base' ? meta.baseCol : side === 'ours' ? meta.oursCol : meta.theirsCol;
      if (!colNumber) {
        arr.push(null);
        continue;
      }
      const cell = row.getCell(colNumber);
      arr.push(getSimpleValueForMerge(cell?.value));
    }
    return arr;
  };

  const result = {
    sheetName: resolvedSheetName,
    rowNumber: fallbackRow ?? undefined,
    baseRowNumber: baseRowNumber ?? null,
    oursRowNumber: oursRowNumber ?? null,
    theirsRowNumber: theirsRowNumber ?? null,
    colCount,
    base: readRowAligned(baseWs, baseRowNumber ?? null, 'base'),
    ours: readRowAligned(oursWs, oursRowNumber ?? null, 'ours'),
    theirs: readRowAligned(theirsWs, theirsRowNumber ?? null, 'theirs'),
  };
  appendDebugLog('main', 'getThreeWayRow:end', {
    requestId,
    durationMs: Date.now() - startedAt,
    sheetName: resolvedSheetName,
    colCount,
  });
  return result;
});
ipcMain.handle('excel:getThreeWayRows', async (_event, req: ThreeWayRowsRequest): Promise<ThreeWayRowsResult | null> => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath || !Array.isArray(req.rows)) return null;
  const requestId = req.debugRequestId ?? `rows-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const startedAt = Date.now();

  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(req.basePath),
    loadWorkbookCached(req.oursPath),
    loadWorkbookCached(req.theirsPath),
  ]);

  const baseWs = getWorksheetSafe(baseWb, req.sheetName, req.sheetIndex);
  const oursWs = getWorksheetSafe(oursWb, req.sheetName, req.sheetIndex);
  const theirsWs = getWorksheetSafe(theirsWb, req.sheetName, req.sheetIndex);

  const resolvedSheetName = baseWs?.name ?? req.sheetName ?? '';
  const headerCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const resolvedHeaderCount = resolveStructuredHeaderRowCount([baseWs, oursWs, theirsWs], headerCount);
  const compareMode = normalizeThreeWayCompareMode(req.compareMode);
  const runtimeConfig = getThreeWayRuntimeConfig(compareMode);
  const baseWsForAlign = runtimeConfig.alignBaseSide === 'ours' ? oursWs : baseWs;
  const alignedColumns = buildAlignedColumns(baseWsForAlign, oursWs, theirsWs, resolvedHeaderCount);
  const colCount = alignedColumns.length;
  appendDebugLog('main', 'getThreeWayRows:start', {
    requestId,
    compareMode,
    sheetName: resolvedSheetName,
    rowCount: req.rows.length,
    frozenRowCount: resolvedHeaderCount,
  });

  const readRowAligned = (
    ws: any,
    rowNum: number | null,
    side: 'base' | 'ours' | 'theirs',
  ): (string | number | null)[] => {
    const arr: (string | number | null)[] = [];
    if (!rowNum) {
      for (let col = 1; col <= colCount; col += 1) arr.push(null);
      return arr;
    }
    const row = ws.getRow(rowNum);
    for (let i = 0; i < alignedColumns.length; i += 1) {
      const meta = alignedColumns[i];
      const colNumber =
        side === 'base' ? meta.baseCol : side === 'ours' ? meta.oursCol : meta.theirsCol;
      if (!colNumber) {
        arr.push(null);
        continue;
      }
      const cell = row.getCell(colNumber);
      arr.push(getSimpleValueForMerge(cell?.value));
    }
    return arr;
  };

  const rows: ThreeWayRowResult[] = req.rows.map((r) => {
    const fallbackRow =
      typeof r.rowNumber === 'number' && !Number.isNaN(r.rowNumber) ? Math.max(1, Math.floor(r.rowNumber)) : null;
    const baseRowNumber = resolveRequestedRowNumber(r.baseRowNumber, fallbackRow);
    const oursRowNumber = resolveRequestedRowNumber(r.oursRowNumber, fallbackRow);
    const theirsRowNumber = resolveRequestedRowNumber(r.theirsRowNumber, fallbackRow);

    return {
      sheetName: resolvedSheetName,
      rowNumber: fallbackRow ?? undefined,
      baseRowNumber: baseRowNumber ?? null,
      oursRowNumber: oursRowNumber ?? null,
      theirsRowNumber: theirsRowNumber ?? null,
      colCount,
      base: readRowAligned(baseWs, baseRowNumber ?? null, 'base'),
      ours: readRowAligned(oursWs, oursRowNumber ?? null, 'ours'),
      theirs: readRowAligned(theirsWs, theirsRowNumber ?? null, 'theirs'),
    };
  });

  appendDebugLog('main', 'getThreeWayRows:end', {
    requestId,
    durationMs: Date.now() - startedAt,
    sheetName: resolvedSheetName,
    rowCount: rows.length,
    colCount,
  });
  return { sheetName: resolvedSheetName, colCount, rows };
});

export const __testOnly = {
  buildMergeSheetsForWorkbooks,
  clearWorkbookCache: () => workbookCache.clear(),
  saveMergeResultDirect: (req: SaveMergeRequest, targetPath: string) =>
    saveMergeResultInternal(req, { targetPathOverride: targetPath, skipMainWindowRequirement: true }),
};
