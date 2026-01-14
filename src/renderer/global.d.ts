import type {
  CellChange,
  CliThreeWayInfo,
  OpenResult,
  SaveMergeRequest,
  SaveMergeResponse,
  ThreeWayOpenResult,
} from '../main/preload';

declare global {
  interface Window {
    excelAPI: {
      openFile: () => Promise<OpenResult | null>;
      saveChanges: (changes: CellChange[]) => Promise<void>;
      openThreeWay: () => Promise<ThreeWayOpenResult | null>;
      saveMergeResult: (req: SaveMergeRequest) => Promise<SaveMergeResponse>;
      getCliThreeWayInfo: () => Promise<CliThreeWayInfo | null>;
    };
  }
}

export {};
