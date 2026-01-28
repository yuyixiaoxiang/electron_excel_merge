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
const DEFAULT_FROZEN_HEADER_ROWS = 3;

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
  const argStartIndex = app?.isPackaged ? 1 : 2;
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
    mainWindow.webContents.openDevTools();
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
type SimpleCellValue = string | number | null;

interface RowRecord {
  rowNumber: number; // 1-based Excel row number
  index: number; // 0-based index in extracted rows list
  values: SimpleCellValue[];
  nonEmptyCols: number[]; // 1-based column indices with non-empty values
  key?: string | null;
}

interface AlignedRow {
  base?: RowRecord | null;
  ours?: RowRecord | null;
  theirs?: RowRecord | null;
  key?: string | null;
  ambiguousOurs?: boolean;
  ambiguousTheirs?: boolean;
}

const getSimpleValueForMerge = (v: any): SimpleCellValue => {
  if (v === null || v === undefined) return null;
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

const normalizeCellValue = (v: SimpleCellValue): string => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  return String(v);
};

const normalizeKeyValue = (v: SimpleCellValue): string | null => {
  const s = normalizeCellValue(v);
  return s === '' ? null : s;
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

const buildRowRecords = (ws: any, colCount: number, primaryKeyCol: number): RowRecord[] => {
  const rows: RowRecord[] = [];
  let index = 0;
  ws.eachRow({ includeEmpty: false }, (row: any, rowNumber: number) => {
    const values: SimpleCellValue[] = [];
    const nonEmptyCols: number[] = [];
    for (let col = 1; col <= colCount; col += 1) {
      const cell = row.getCell(col);
      const value = getSimpleValueForMerge(cell?.value);
      values.push(value);
      if (value !== null && value !== '') {
        nonEmptyCols.push(col);
      }
    }
    if (nonEmptyCols.length === 0) return;
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

const rowsEqual = (a: RowRecord, b: RowRecord): boolean => {
  const cols = new Set<number>();
  a.nonEmptyCols.forEach((c) => cols.add(c));
  b.nonEmptyCols.forEach((c) => cols.add(c));
  for (const col of cols) {
    const av = normalizeCellValue(a.values[col - 1] ?? null);
    const bv = normalizeCellValue(b.values[col - 1] ?? null);
    if (av !== bv) return false;
  }
  return true;
};

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
    if (av === '' && bv === '') continue;
    total += 1;
    if (av === bv) same += 1;
  }
  if (total === 0) return 1;
  return same / total;
};

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

type DiffOp =
  | { type: 'equal'; aIndex: number; bIndex: number }
  | { type: 'delete'; aIndex: number }
  | { type: 'insert'; bIndex: number };

const myersDiff = (a: string[], b: string[]): DiffOp[] => {
  const n = a.length;
  const m = b.length;
  const max = n + m;
  let v = new Map<number, number>();
  v.set(1, 0);
  const trace: Map<number, number>[] = [];

  for (let d = 0; d <= max; d += 1) {
    const vSnap = new Map<number, number>();
    for (let k = -d; k <= d; k += 2) {
      let x: number;
      if (k === -d || (k !== d && (v.get(k - 1) ?? 0) < (v.get(k + 1) ?? 0))) {
        x = v.get(k + 1) ?? 0;
      } else {
        x = (v.get(k - 1) ?? 0) + 1;
      }
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) {
        x += 1;
        y += 1;
      }
      vSnap.set(k, x);
      if (x >= n && y >= m) {
        trace.push(vSnap);
        // backtrack
        const ops: DiffOp[] = [];
        let x2 = n;
        let y2 = m;
        for (let d2 = trace.length - 1; d2 >= 0; d2 -= 1) {
          const v2 = trace[d2];
          const k2 = x2 - y2;
          let prevK: number;
          if (k2 === -d2 || (k2 !== d2 && (v2.get(k2 - 1) ?? 0) < (v2.get(k2 + 1) ?? 0))) {
            prevK = k2 + 1;
          } else {
            prevK = k2 - 1;
          }
          const prevX = v2.get(prevK) ?? 0;
          const prevY = prevX - prevK;
          while (x2 > prevX && y2 > prevY) {
            ops.push({ type: 'equal', aIndex: x2 - 1, bIndex: y2 - 1 });
            x2 -= 1;
            y2 -= 1;
          }
          if (d2 === 0) break;
          if (x2 === prevX) {
            ops.push({ type: 'insert', bIndex: y2 - 1 });
            y2 -= 1;
          } else {
            ops.push({ type: 'delete', aIndex: x2 - 1 });
            x2 -= 1;
          }
        }
        return ops.reverse();
      }
    }
    trace.push(vSnap);
    v = vSnap;
  }
  return [];
};

const alignRowsByKey = (
  baseRows: RowRecord[],
  oursRows: RowRecord[],
  theirsRows: RowRecord[],
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

  const alignedBase: AlignedRow[] = baseRows.map((baseRow) => {
    const key = baseRow.key ?? null;
    if (!key) {
      return {
        base: baseRow,
        ours: null,
        theirs: null,
        key,
        ambiguousOurs: true,
        ambiguousTheirs: true,
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

    if (oursCount === 0) {
      ours = null;
    } else if (oursCount === 1 && baseCount === 1) {
      ours = oursList[0] ?? null;
    } else if (oursCount === baseCount && baseCount > 0) {
      ours = oursList[occIndex] ?? null;
    } else {
      ambiguousOurs = true;
    }

    if (theirsCount === 0) {
      theirs = null;
    } else if (theirsCount === 1 && baseCount === 1) {
      theirs = theirsList[0] ?? null;
    } else if (theirsCount === baseCount && baseCount > 0) {
      theirs = theirsList[occIndex] ?? null;
    } else {
      ambiguousTheirs = true;
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
    for (const r of oursGap) {
      const ambiguous = !r.key;
      aligned.push({ ours: r, key: r.key ?? null, ambiguousOurs: ambiguous });
    }
    for (const r of theirsGap) {
      const ambiguous = !r.key;
      aligned.push({ theirs: r, key: r.key ?? null, ambiguousTheirs: ambiguous });
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
  const buildTokens = (rows: RowRecord[]) =>
    rows.map((r) => r.values.map((v) => normalizeCellValue(v)).join('||'));

  const similarityThreshold = 0.7;
  const similarityDelta = 0.05;
  const windowSize = 3;

  const alignOneSide = (sideRows: RowRecord[]) => {
    const baseTokens = buildTokens(baseRows);
    const sideTokens = buildTokens(sideRows);
    const ops = myersDiff(baseTokens, sideTokens);
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

    const matchedPairs = Array.from(matched.entries()).map(([baseIndex, sideIndex]) => ({ baseIndex, sideIndex }));
    matchedPairs.sort((a, b) => a.baseIndex - b.baseIndex);

    const ambiguousBase = new Set<number>();
    const ambiguousSide = new Set<number>();
    const unmatchedInserts = new Set<number>(inserts);

    for (const baseIndex of deletes) {
      const baseRow = baseRows[baseIndex];
      if (!baseRow) continue;
      const expected = estimateSideIndex(baseIndex, matchedPairs);
      const candidates: Array<{ index: number; score: number }> = [];
      for (const sideIndex of inserts) {
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

const buildMergeSheetWithRowAlign = (
  baseWs: any,
  oursWs: any,
  theirsWs: any,
  primaryKeyCol: number,
  frozenRowCount: number,
): MergeSheetData => {
  const getRowCount = (ws: any) =>
    (ws?.actualRowCount ?? 0) > 0 ? ws.actualRowCount : ws?.rowCount ?? 0;
  const getColCount = (ws: any) =>
    (ws?.actualColumnCount ?? 0) > 0 ? ws.actualColumnCount : ws?.columnCount ?? 0;
  const hasExactCellDiff = (base: any, ours: any, theirs: any) => {
    const maxRow = Math.max(getRowCount(base), getRowCount(ours), getRowCount(theirs));
    const maxCol = Math.max(getColCount(base), getColCount(ours), getColCount(theirs));
    for (let r = 1; r <= maxRow; r += 1) {
      const baseRow = base.getRow(r);
      const oursRow = ours.getRow(r);
      const theirsRow = theirs.getRow(r);
      for (let c = 1; c <= maxCol; c += 1) {
        const baseValue = getSimpleValueForMerge(baseRow.getCell(c)?.value);
        const oursValue = getSimpleValueForMerge(oursRow.getCell(c)?.value);
        const theirsValue = getSimpleValueForMerge(theirsRow.getCell(c)?.value);
        if (baseValue !== oursValue || baseValue !== theirsValue || oursValue !== theirsValue) {
          return true;
        }
      }
    }
    return false;
  };
  const colCount = Math.max(
    baseWs?.actualColumnCount ?? baseWs?.columnCount ?? 0,
    oursWs?.actualColumnCount ?? oursWs?.columnCount ?? 0,
    theirsWs?.actualColumnCount ?? theirsWs?.columnCount ?? 0,
  );
  const headerCount = Math.max(0, Math.floor(frozenRowCount));
  const useKey = primaryKeyCol >= 1 && primaryKeyCol <= colCount;
  const keyCol = useKey ? primaryKeyCol : -1;
  const baseRows = buildRowRecords(baseWs, colCount, keyCol).filter((r) => r.rowNumber > headerCount);
  const oursRows = buildRowRecords(oursWs, colCount, keyCol).filter((r) => r.rowNumber > headerCount);
  const theirsRows = buildRowRecords(theirsWs, colCount, keyCol).filter((r) => r.rowNumber > headerCount);

  const alignedResult = useKey
    ? alignRowsByKey(baseRows, oursRows, theirsRows)
    : alignRowsBySequence(baseRows, oursRows, theirsRows);

  const aligned = alignedResult.aligned;

  const rowsMeta: MergeRowMeta[] = [];
  // 1) Header rows: compare by fixed row number (no alignment)
  for (let r = 1; r <= headerCount; r += 1) {
    const baseRow = buildHeaderRowRecord(baseWs, r, colCount, keyCol);
    const oursRow = buildHeaderRowRecord(oursWs, r, colCount, keyCol);
    const theirsRow = buildHeaderRowRecord(theirsWs, r, colCount, keyCol);
    rowsMeta.push({
      visualRowNumber: r,
      key: baseRow.key ?? oursRow.key ?? theirsRow.key ?? null,
      baseRowNumber: r,
      oursRowNumber: r,
      theirsRowNumber: r,
      oursStatus: computeRowStatus(baseRow, oursRow, false),
      theirsStatus: computeRowStatus(baseRow, theirsRow, false),
    });
  }
  // 2) Body rows: aligned
  aligned.forEach((row, idx) => {
    const visualRowNumber = headerCount + idx + 1;
    rowsMeta.push({
      visualRowNumber,
      key: row.key ?? row.base?.key ?? row.ours?.key ?? row.theirs?.key ?? null,
      baseRowNumber: row.base?.rowNumber ?? null,
      oursRowNumber: row.ours?.rowNumber ?? null,
      theirsRowNumber: row.theirs?.rowNumber ?? null,
      oursStatus: computeRowStatus(row.base ?? null, row.ours ?? null, row.ambiguousOurs),
      theirsStatus: computeRowStatus(row.base ?? null, row.theirs ?? null, row.ambiguousTheirs),
    });
  });

  const same = (a: SimpleCellValue, b: SimpleCellValue) => normalizeCellValue(a) === normalizeCellValue(b);
  const cells: MergeCell[] = [];
  const hasExactDiff = hasExactCellDiff(baseWs, oursWs, theirsWs);

  // Header rows diff by fixed row number
  for (let r = 1; r <= headerCount; r += 1) {
    const baseRow = buildHeaderRowRecord(baseWs, r, colCount, keyCol);
    const oursRow = buildHeaderRowRecord(oursWs, r, colCount, keyCol);
    const theirsRow = buildHeaderRowRecord(theirsWs, r, colCount, keyCol);
    const cols = new Set<number>();
    baseRow.nonEmptyCols.forEach((c) => cols.add(c));
    oursRow.nonEmptyCols.forEach((c) => cols.add(c));
    theirsRow.nonEmptyCols.forEach((c) => cols.add(c));
    for (const col of cols) {
      const baseValue = baseRow.values[col - 1] ?? null;
      const oursValue = oursRow.values[col - 1] ?? null;
      const theirsValue = theirsRow.values[col - 1] ?? null;

      const equalBO = same(baseValue, oursValue);
      const equalBT = same(baseValue, theirsValue);
      const equalOT = same(oursValue, theirsValue);

      let status: MergeCell['status'];
      let mergedValue: SimpleCellValue = baseValue;

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
        mergedValue = oursValue;
      }

      if (status !== 'unchanged') {
        cells.push({
          address: makeAddress(col, r),
          row: r,
          col,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
      }
    }
  }

  // Body rows diff via alignment
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

      const equalBO = same(baseValue, oursValue);
      const equalBT = same(baseValue, theirsValue);
      const equalOT = same(oursValue, theirsValue);

      let status: MergeCell['status'];
      let mergedValue: SimpleCellValue = baseValue;

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
        mergedValue = oursValue;
      }

      if (status !== 'unchanged') {
        const addressRow =
          row.ours?.rowNumber ?? row.base?.rowNumber ?? row.theirs?.rowNumber ?? visualRowNumber;
        cells.push({
          address: makeAddress(col, addressRow),
          row: visualRowNumber,
          col,
          baseValue,
          oursValue,
          theirsValue,
          status,
          mergedValue,
        });
      }
    }
  });

  // 如果有差异列，为冻结行补齐这些列的内容（即使未变化），用于显示表头/冻结行上下文
  if (headerCount > 0 && cells.length > 0) {
    const diffColumns = new Set<number>(cells.map((c) => c.col));
    if (diffColumns.size > 0) {
      const existing = new Set<string>(cells.map((c) => `${c.row}:${c.col}`));
      for (let r = 1; r <= headerCount; r += 1) {
        const baseRow = buildHeaderRowRecord(baseWs, r, colCount, keyCol);
        const oursRow = buildHeaderRowRecord(oursWs, r, colCount, keyCol);
        const theirsRow = buildHeaderRowRecord(theirsWs, r, colCount, keyCol);
        for (const col of diffColumns) {
          const key = `${r}:${col}`;
          if (existing.has(key)) continue;
          const baseValue = baseRow.values[col - 1] ?? null;
          const oursValue = oursRow.values[col - 1] ?? null;
          const theirsValue = theirsRow.values[col - 1] ?? null;
          cells.push({
            address: makeAddress(col, r),
            row: r,
            col,
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
  };
};

const buildMergeSheetsForWorkbooks = async (
  basePath: string,
  oursPath: string,
  theirsPath: string,
  primaryKeyCol: number,
  frozenRowCount: number,
) => {
  const baseWb = new Workbook();
  const oursWb = new Workbook();
  const theirsWb = new Workbook();

  await baseWb.xlsx.readFile(basePath);
  await oursWb.xlsx.readFile(oursPath);
  await theirsWb.xlsx.readFile(theirsPath);

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

  // 规则：优先按同名工作表对齐；对剩余未匹配的工作表，再按索引对齐（第 1 张对第 1 张……）。
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

    mergeSheets.push(buildMergeSheetWithRowAlign(baseWs, oursHit.ws, theirsHit.ws, primaryKeyCol, frozenRowCount));
  }

  // 2) 索引兜底：仅对“同一 idx 在三边都没被用过”的位置做对齐
  const count = Math.min(baseList.length, oursList.length, theirsList.length);
  for (let idx = 0; idx < count; idx += 1) {
    if (usedBaseIdx.has(idx) || usedOursIdx.has(idx) || usedTheirsIdx.has(idx)) continue;
    usedBaseIdx.add(idx);
    usedOursIdx.add(idx);
    usedTheirsIdx.add(idx);
    mergeSheets.push(
      buildMergeSheetWithRowAlign(baseList[idx], oursList[idx], theirsList[idx], primaryKeyCol, frozenRowCount),
    );
  }

  return { basePath, oursPath, theirsPath, mergeSheets };
};

const normalizeThreeWayResult = (
  basePath: string,
  oursPath: string,
  theirsPath: string,
  mergeSheets: MergeSheetData[],
) => {
  const emptySheet: MergeSheetData = { sheetName: '', cells: [], rowsMeta: [] };
  return {
    basePath,
    oursPath,
    theirsPath,
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
  /** 视觉行号（diff/merge 视图中的 1-based 行号） */
  visualRowNumber: number;
  /** 如果启用了主键列，这里记录主键（normalize 后） */
  key?: string | null;
  /** 三方文件中各自对应的原始行号（1-based）；不存在则为 null */
  baseRowNumber: number | null;
  oursRowNumber: number | null;
  theirsRowNumber: number | null;
  /** 该视觉行在对应 side 相对 base 的状态 */
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
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  status: 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
  mergedValue: string | number | null;
}

interface MergeSheetData {
  sheetName: string;
  cells: MergeCell[];
  rowsMeta?: MergeRowMeta[];
  hasExactDiff?: boolean;
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

      // Date
      if (raw instanceof Date) {
        // 保持可读性，避免显示为 [object Object]
        return raw.toISOString();
      }

      // 富文本：raw.richText 是一个包含 { text } 的数组
      if (typeof raw === 'object' && Array.isArray((raw as any).richText)) {
        const parts = (raw as any).richText
          .map((p: any) => (p && typeof p.text === 'string' ? p.text : ''))
          .join('');
        return parts;
      }

      // Hyperlink / text-like objects
      if (typeof raw === 'object' && raw && 'text' in (raw as any)) {
        const t = (raw as any).text;
        if (t === null || t === undefined) return null;
        return typeof t === 'string' || typeof t === 'number' ? (t as any) : String(t);
      }

      // Formula / shared formula 等：优先显示 result
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

      // 兜底：尽量 JSON 序列化，避免 [object Object]
      if (typeof raw === 'object') {
        try {
          return JSON.stringify(raw);
        } catch {
          return String(raw);
        }
      }

      return String(raw);
    };

    // 重要：确保每一行的列数一致。
    // 否则会出现“数据行列数 > 表头/冻结行列数”造成错位。
    const maxRow =
      (worksheet as any).actualRowCount && (worksheet as any).actualRowCount > 0
        ? (worksheet as any).actualRowCount
        : worksheet.rowCount;
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
  const primaryKeyCol = 1;
  const frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS;

  if (cliThreeWayArgs) {
    const { basePath, oursPath, theirsPath } = cliThreeWayArgs;
    const { mergeSheets } = await buildMergeSheetsForWorkbooks(
      basePath,
      oursPath,
      theirsPath,
      primaryKeyCol,
      frozenRowCount,
    );
    return normalizeThreeWayResult(basePath, oursPath, theirsPath, mergeSheets);
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

  const { mergeSheets } = await buildMergeSheetsForWorkbooks(basePath, oursPath, theirsPath, primaryKeyCol, frozenRowCount);

  return normalizeThreeWayResult(basePath, oursPath, theirsPath, mergeSheets);
});
interface ThreeWayDiffRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  primaryKeyCol: number; // 1-based, -1 means no primary key
  frozenRowCount?: number; // header rows compared by coordinates
}

ipcMain.handle('excel:computeThreeWayDiff', async (_event, req: ThreeWayDiffRequest) => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath) return null;
  const primaryKeyCol =
    typeof req.primaryKeyCol === 'number' && !Number.isNaN(req.primaryKeyCol) ? Math.floor(req.primaryKeyCol) : 1;
  const frozenRowCount =
    typeof req.frozenRowCount === 'number' && !Number.isNaN(req.frozenRowCount)
      ? Math.max(0, Math.floor(req.frozenRowCount))
      : DEFAULT_FROZEN_HEADER_ROWS;
  const { mergeSheets } = await buildMergeSheetsForWorkbooks(
    req.basePath,
    req.oursPath,
    req.theirsPath,
    primaryKeyCol,
    frozenRowCount,
  );
  return normalizeThreeWayResult(req.basePath, req.oursPath, req.theirsPath, mergeSheets);
});

// 将 CLI three-way 信息暴露给渲染进程，便于自动加载
ipcMain.handle('excel:getCliThreeWayInfo', async () => {
  if (!cliThreeWayArgs) return null;
  return cliThreeWayArgs;
});

// 读取三方文件的“某一行”数据，用于底部行级对比视图
interface ThreeWayRowRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
  rowNumber?: number; // 1-based fallback for all sides
  baseRowNumber?: number | null;
  oursRowNumber?: number | null;
  theirsRowNumber?: number | null;
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

// 简单缓存：同一次应用生命周期内重复读取同一个 xlsx 时复用 workbook，减少 IO
const workbookCache = new Map<string, Workbook>();

const loadWorkbookCached = async (filePath: string): Promise<Workbook> => {
  const hit = workbookCache.get(filePath);
  if (hit) return hit;
  const wb = new Workbook();
  await wb.xlsx.readFile(filePath);
  workbookCache.set(filePath, wb);
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

const getSimpleValueForThreeWay = (v: any): string | number | null => {
  if (v === null || v === undefined) return null;
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

ipcMain.handle('excel:getThreeWayRow', async (_event, req: ThreeWayRowRequest): Promise<ThreeWayRowResult | null> => {
  if (!req || !req.basePath || !req.oursPath || !req.theirsPath) return null;
  const fallbackRow =
    typeof req.rowNumber === 'number' && !Number.isNaN(req.rowNumber)
      ? Math.max(1, Math.floor(req.rowNumber))
      : null;
  const baseRowNumber =
    typeof req.baseRowNumber === 'number' && !Number.isNaN(req.baseRowNumber)
      ? Math.max(1, Math.floor(req.baseRowNumber))
      : fallbackRow;
  const oursRowNumber =
    typeof req.oursRowNumber === 'number' && !Number.isNaN(req.oursRowNumber)
      ? Math.max(1, Math.floor(req.oursRowNumber))
      : fallbackRow;
  const theirsRowNumber =
    typeof req.theirsRowNumber === 'number' && !Number.isNaN(req.theirsRowNumber)
      ? Math.max(1, Math.floor(req.theirsRowNumber))
      : fallbackRow;

  const [baseWb, oursWb, theirsWb] = await Promise.all([
    loadWorkbookCached(req.basePath),
    loadWorkbookCached(req.oursPath),
    loadWorkbookCached(req.theirsPath),
  ]);

  const baseWs = getWorksheetSafe(baseWb, req.sheetName, req.sheetIndex);
  const oursWs = getWorksheetSafe(oursWb, req.sheetName, req.sheetIndex);
  const theirsWs = getWorksheetSafe(theirsWb, req.sheetName, req.sheetIndex);

  const resolvedSheetName = baseWs?.name ?? req.sheetName ?? '';

  const colCount = Math.max(
    baseWs?.actualColumnCount ?? baseWs?.columnCount ?? 0,
    oursWs?.actualColumnCount ?? oursWs?.columnCount ?? 0,
    theirsWs?.actualColumnCount ?? theirsWs?.columnCount ?? 0,
  );

  const readRow = (ws: any, rowNum: number | null): (string | number | null)[] => {
    const arr: (string | number | null)[] = [];
    if (!rowNum) {
      for (let col = 1; col <= colCount; col += 1) arr.push(null);
      return arr;
    }
    const row = ws.getRow(rowNum);
    for (let col = 1; col <= colCount; col += 1) {
      const cell = row.getCell(col);
      arr.push(getSimpleValueForThreeWay(cell?.value));
    }
    return arr;
  };

  return {
    sheetName: resolvedSheetName,
    rowNumber: fallbackRow ?? undefined,
    baseRowNumber: baseRowNumber ?? null,
    oursRowNumber: oursRowNumber ?? null,
    theirsRowNumber: theirsRowNumber ?? null,
    colCount,
    base: readRow(baseWs, baseRowNumber ?? null),
    ours: readRow(oursWs, oursRowNumber ?? null),
    theirs: readRow(theirsWs, theirsRowNumber ?? null),
  };
});
