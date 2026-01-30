/**
 * 预加载脚本：在隔离上下文中暴露一个 typesafe 的 excelAPI 到 window，
 * 让渲染进程只能通过这里定义好的 IPC 通道访问主进程能力。
 */
import { contextBridge, ipcRenderer } from 'electron';

// 以下接口与 main.ts 中的同名类型需要保持结构一致，
// 方便在 renderer 侧进行类型推导与复用。
interface SheetCell {
  address: string;
  row: number;
  col: number;
  value: string | number | null;
}

interface SheetData {
  sheetName: string;
  rows: SheetCell[][];
}
interface GetSheetDataRequest {
  path: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
}

interface OpenResult {
  filePath: string;
  sheet: SheetData; // 兼容旧字段：第一个 sheet
  sheets: SheetData[];
}

interface CellChange {
  address: string;
  newValue: string | number | null;
}

// Merge diff types
type RowStatus = 'unchanged' | 'added' | 'deleted' | 'modified' | 'ambiguous';

interface MergeRowMeta {
  visualRowNumber: number;
  key?: string | null;
  baseRowNumber: number | null;
  oursRowNumber: number | null;
  theirsRowNumber: number | null;
  oursSimilarity?: number | null;
  theirsSimilarity?: number | null;
  oursStatus: RowStatus;
  theirsStatus: RowStatus;
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
  // 性能优化：仅传输“可能产生差异”的单元格列表（稀疏结构），避免把整张表矩阵通过 IPC 传到渲染进程
  cells: MergeCell[];
  rowsMeta?: MergeRowMeta[];
  hasExactDiff?: boolean;
}

interface ThreeWayOpenResult {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheet: MergeSheetData; // 第一个 sheet
  sheets: MergeSheetData[];
}

interface ThreeWayDiffRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  primaryKeyCol: number; // 1-based, -1 means no primary key
  frozenRowCount?: number; // number of header rows to compare by coordinates
  rowSimilarityThreshold?: number; // 0-1
}

interface SaveMergeCellInput {
  sheetName: string;
  address: string;
  value: string | number | null;
}
interface SaveMergeRowOp {
  sheetName: string;
  action: 'insert' | 'delete';
  targetRowNumber: number; // 1-based in template (ours)
  values?: (string | number | null)[];
  visualRowNumber?: number; // for stable ordering
}

interface SaveMergeRequest {
  templatePath: string;
  cells: SaveMergeCellInput[];
  rowOps?: SaveMergeRowOp[];
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

interface CliThreeWayInfo {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  mergedPath?: string;
  mode: 'diff' | 'merge';
}

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
interface ThreeWayRowsRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheetName?: string;
  sheetIndex?: number; // 0-based
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

/**
 * 暴露给渲染进程的所有 Excel 相关操作。
 *
 * 注意：这里只包装了 ipcRenderer.invoke，真正的实现都在 main.ts 中。
 */
const excelAPI = {
  openFile: async (): Promise<OpenResult | null> => {
    const result = await ipcRenderer.invoke('excel:open');
    return result as OpenResult | null;
  },
  saveChanges: async (changes: CellChange[]): Promise<void> => {
    await ipcRenderer.invoke('excel:saveChanges', changes);
  },
  openThreeWay: async (): Promise<ThreeWayOpenResult | null> => {
    const result = await ipcRenderer.invoke('excel:openThreeWay');
    return result as ThreeWayOpenResult | null;
  },
  getSheetData: async (req: GetSheetDataRequest): Promise<SheetData | null> => {
    const result = await ipcRenderer.invoke('excel:getSheetData', req);
    return result as SheetData | null;
  },
  computeThreeWayDiff: async (req: ThreeWayDiffRequest): Promise<ThreeWayOpenResult | null> => {
    const result = await ipcRenderer.invoke('excel:computeThreeWayDiff', req);
    return result as ThreeWayOpenResult | null;
  },
  saveMergeResult: async (req: SaveMergeRequest): Promise<SaveMergeResponse> => {
    const result = await ipcRenderer.invoke('excel:saveMergeResult', req);
    return result as SaveMergeResponse;
  },
  getCliThreeWayInfo: async (): Promise<CliThreeWayInfo | null> => {
    const result = await ipcRenderer.invoke('excel:getCliThreeWayInfo');
    return result as CliThreeWayInfo | null;
  },
  getThreeWayRow: async (req: ThreeWayRowRequest): Promise<ThreeWayRowResult | null> => {
    const result = await ipcRenderer.invoke('excel:getThreeWayRow', req);
    return result as ThreeWayRowResult | null;
  },
  getThreeWayRows: async (req: ThreeWayRowsRequest): Promise<ThreeWayRowsResult | null> => {
    const result = await ipcRenderer.invoke('excel:getThreeWayRows', req);
    return result as ThreeWayRowsResult | null;
  },
};

contextBridge.exposeInMainWorld('excelAPI', excelAPI);

export type {
  SheetCell,
  SheetData,
  GetSheetDataRequest,
  OpenResult,
  CellChange,
  MergeCell,
  MergeSheetData,
  MergeRowMeta,
  RowStatus,
  ThreeWayOpenResult,
  ThreeWayDiffRequest,
  SaveMergeCellInput,
  SaveMergeRowOp,
  SaveMergeRequest,
  SaveMergeResponse,
  CliThreeWayInfo,
  ThreeWayRowRequest,
  ThreeWayRowResult,
  ThreeWayRowsRequest,
  ThreeWayRowsResult,
};
