import type {
  CellChange,
  CliThreeWayInfo,
  DebugLogEntry,
  FolderExcelFileInfo,
  GetSheetDataRequest,
  OpenResult,
  SaveChangesRequest,
  SaveMergeRequest,
  SaveMergeResponse,
  SheetData,
  ThreeWayDiffRequest,
  ThreeWayOpenResult,
  ThreeWayRowRequest,
  ThreeWayRowResult,
  ThreeWayRowsRequest,
  ThreeWayRowsResult,
  WorkspaceTabMenuEvent,
} from '../main/preload';

/**
 * 声明在预加载脚本中通过 contextBridge 暴露到 window 上的 excelAPI，
 * 这样在 React 代码中使用 window.excelAPI 时可以获得完整的类型提示。
 */
declare global {
  interface Window {
    excelAPI: {
      pickFolder: () => Promise<string | null>;
      listExcelFilesInFolder: (folderPath: string) => Promise<FolderExcelFileInfo[]>;
      openFile: () => Promise<OpenResult | null>;
      loadWorkbook: (filePath: string) => Promise<OpenResult | null>;
      saveChanges: (req: SaveChangesRequest | CellChange[]) => Promise<void>;
      openThreeWay: () => Promise<ThreeWayOpenResult | null>;
      getSheetData: (req: GetSheetDataRequest) => Promise<SheetData | null>;
      computeThreeWayDiff: (req: ThreeWayDiffRequest) => Promise<ThreeWayOpenResult | null>;
      saveMergeResult: (req: SaveMergeRequest) => Promise<SaveMergeResponse>;
      getCliThreeWayInfo: () => Promise<CliThreeWayInfo | null>;
      getThreeWayRow: (req: ThreeWayRowRequest) => Promise<ThreeWayRowResult | null>;
      getThreeWayRows: (req: ThreeWayRowsRequest) => Promise<ThreeWayRowsResult | null>;
      debugLog: (entry: DebugLogEntry) => void;
      getDebugLogPath: () => Promise<string>;
      onWorkspaceNewTab: (handler: (payload: WorkspaceTabMenuEvent) => void) => () => void;
    };
  }
}

export {};
