import type {
  CellChange,
  CliThreeWayInfo,
  OpenResult,
  SaveMergeRequest,
  SaveMergeResponse,
  ThreeWayDiffRequest,
  ThreeWayOpenResult,
  ThreeWayRowRequest,
  ThreeWayRowResult,
} from '../main/preload';

/**
 * 声明在预加载脚本中通过 contextBridge 暴露到 window 上的 excelAPI，
 * 这样在 React 代码中使用 window.excelAPI 时可以获得完整的类型提示。
 */
declare global {
  interface Window {
    excelAPI: {
      openFile: () => Promise<OpenResult | null>;
      saveChanges: (changes: CellChange[]) => Promise<void>;
      openThreeWay: () => Promise<ThreeWayOpenResult | null>;
      computeThreeWayDiff: (req: ThreeWayDiffRequest) => Promise<ThreeWayOpenResult | null>;
      saveMergeResult: (req: SaveMergeRequest) => Promise<SaveMergeResponse>;
      getCliThreeWayInfo: () => Promise<CliThreeWayInfo | null>;
      getThreeWayRow: (req: ThreeWayRowRequest) => Promise<ThreeWayRowResult | null>;
    };
  }
}

export {};
