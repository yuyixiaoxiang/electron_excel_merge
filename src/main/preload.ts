/**
 * 共享契约：保留原 preload 中的类型定义，供 Tauri 版前端和服务层复用。
 */

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
interface FolderExcelFileInfo {
  relativePath: string;
  absolutePath: string;
  sizeBytes: number;
  modifiedAtMs: number;
}
interface WorkspaceTabMenuEvent {
  kind: 'folder' | 'diff' | 'merge';
}

interface CellChange {
  address: string;
  newValue: string | number | null;
}
interface SaveChangesRequest {
  changes: CellChange[];
  sheetName?: string;
  sheetIndex?: number; // 0-based
  filePath?: string;
  rowOps?: SaveMergeRowOp[];
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
  baseCol?: number | null;
  oursCol?: number | null;
  theirsCol?: number | null;
  formulaControlled?: boolean;
  sharedControlled?: boolean;
  sharedControlGroupKey?: string | null;
  sharedControlMasterSheetName?: string | null;
  sharedControlIsMaster?: boolean;
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  status: 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
  mergedValue: string | number | null;
}
interface MergeColumnMeta {
  col: number;
  baseCol: number | null;
  oursCol: number | null;
  theirsCol: number | null;
}

type PrimaryKeySource = 'manual' | 'auto-implicit' | 'auto-header' | 'auto-weak' | 'none';

interface MergeSheetData {
  sheetName: string;
  // 性能优化：仅传输“可能产生差异”的单元格列表（稀疏结构），避免把整张表矩阵通过 IPC 传到渲染进程
  cells: MergeCell[];
  rowsMeta?: MergeRowMeta[];
  hasExactDiff?: boolean;
  columnsMeta?: MergeColumnMeta[];
  primaryKeyAlignedCol?: number | null;
  primaryKeyOursCol?: number | null;
  primaryKeySource?: PrimaryKeySource;
}

interface ThreeWayOpenResult {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheet: MergeSheetData; // 第一个 sheet
  sheets: MergeSheetData[];
}

type ThreeWayCompareMode = 'diff' | 'merge';

interface ThreeWayDiffRequest {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  compareMode?: ThreeWayCompareMode;
  primaryKeyCol: number; // 1-based manual key; -1 means auto-detect; 0 means force no primary key
  frozenRowCount?: number; // number of header rows to compare by coordinates
  rowSimilarityThreshold?: number; // 0-1
  debugRequestId?: string;
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

interface CliThreeWayInfo {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  mergedPath?: string;
  mergedPathRaw?: string;
  mode: 'diff' | 'merge';
}
interface DebugLogEntry {
  source: string;
  event: string;
  details?: unknown;
}

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

export type {
  SheetCell,
  SheetData,
  GetSheetDataRequest,
  OpenResult,
  FolderExcelFileInfo,
  WorkspaceTabMenuEvent,
  CellChange,
  SaveChangesRequest,
  MergeCell,
  MergeColumnMeta,
  MergeSheetData,
  MergeRowMeta,
  RowStatus,
  ThreeWayOpenResult,
  ThreeWayCompareMode,
  ThreeWayDiffRequest,
  SaveMergeCellInput,
  SaveMergeRowOp,
  SaveMergeColOp,
  SaveMergeRequest,
  SaveMergeResponse,
  CliThreeWayInfo,
  DebugLogEntry,
  ThreeWayRowRequest,
  ThreeWayRowResult,
  ThreeWayRowsRequest,
  ThreeWayRowsResult,
};
