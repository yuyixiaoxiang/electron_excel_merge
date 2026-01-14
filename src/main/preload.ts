import { contextBridge, ipcRenderer } from 'electron';

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

interface ThreeWayOpenResult {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  sheet: MergeSheetData; // 第一个 sheet
  sheets: MergeSheetData[];
}

interface SaveMergeCellInput {
  sheetName: string;
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

interface CliThreeWayInfo {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  mergedPath?: string;
  mode: 'diff' | 'merge';
}

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
  saveMergeResult: async (req: SaveMergeRequest): Promise<SaveMergeResponse> => {
    const result = await ipcRenderer.invoke('excel:saveMergeResult', req);
    return result as SaveMergeResponse;
  },
  getCliThreeWayInfo: async (): Promise<CliThreeWayInfo | null> => {
    const result = await ipcRenderer.invoke('excel:getCliThreeWayInfo');
    return result as CliThreeWayInfo | null;
  },
};

contextBridge.exposeInMainWorld('excelAPI', excelAPI);

export type {
  SheetCell,
  SheetData,
  OpenResult,
  CellChange,
  MergeCell,
  MergeSheetData,
  ThreeWayOpenResult,
  SaveMergeCellInput,
  SaveMergeRequest,
  SaveMergeResponse,
  CliThreeWayInfo,
};
