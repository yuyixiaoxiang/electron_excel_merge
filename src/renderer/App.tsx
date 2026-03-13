import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type {
  CellChange,
  CliThreeWayInfo,
  MergeCell,
  MergeColumnMeta,
  MergeRowMeta,
  MergeSheetData,
  OpenResult,
  SaveMergeColOp,
  SaveMergeRowOp,
  SaveMergeRequest,
  SheetCell,
  SheetData,
  ThreeWayDiffRequest,
  ThreeWayRowResult,
} from '../main/preload';
import { DiffCellData, DiffSideBySide } from './DiffSideBySide';
import { MergeWorkbench } from './MergeWorkbench';
import { VirtualGrid } from './VirtualGrid';

/**
 * 应用根组件：
 * - diff 模式：双文件对比、左右并排编辑；
 * - merge 模式：base / ours / theirs 三方合并与结果写回。
 */
type ViewMode = 'diff' | 'merge';
type DiffSide = 'left' | 'right';
type PrimaryKeyMode = 'auto' | 'manual' | 'none';
type DiffHistorySnapshot = {
  leftWorkbook: OpenResult | null;
  rightWorkbook: OpenResult | null;
  leftChangesBySheet: Map<string, Map<string, CellChange>>;
  rightChangesBySheet: Map<string, Map<string, CellChange>>;
  leftRowOpsBySheet: Map<string, SaveMergeRowOp[]>;
  rightRowOpsBySheet: Map<string, SaveMergeRowOp[]>;
  diffSheets: MergeSheetData[];
  selectedDiffSheetIndex: number;
  diffSelectedCell: { rowIndex: number; colIndex: number } | null;
};
type MergeHistorySnapshot = {
  mergeSheets: MergeSheetData[];
  resolvedBySheet: Map<number, Set<string>>;
  mergeRowOpsBySheet: Map<number, Map<number, SaveMergeRowOp>>;
  mergeColOpsBySheet: Map<number, Map<number, SaveMergeColOp>>;
  selectedMergeSheetIndex: number;
  selectedMergeCell: { rowIndex: number; colIndex: number } | null;
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

const makeAddress = (colNumber: number, rowNumber: number): string => {
  return `${colNumberToLabel(colNumber)}${rowNumber}`;
};
const parseCellAddress = (address: string): { colNumber: number; rowNumber: number } | null => {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match) return null;
  const [, colLabel, rowLabel] = match;
  let colNumber = 0;
  for (const ch of colLabel.toUpperCase()) {
    colNumber = colNumber * 26 + (ch.charCodeAt(0) - 64);
  }
  const rowNumber = Number(rowLabel);
  if (!Number.isFinite(colNumber) || !Number.isFinite(rowNumber)) return null;
  return { colNumber, rowNumber };
};

const normalizeComparableValue = (value: string | number | null): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return String(value);
  return String(value).trim();
};

const sameComparableValue = (a: string | number | null, b: string | number | null): boolean =>
  normalizeComparableValue(a) === normalizeComparableValue(b);
const isEditableEventTarget = (target: EventTarget | null): boolean => {
  if (!(target instanceof Element)) return false;
  if (target instanceof HTMLElement && target.isContentEditable) return true;
  return Boolean(target.closest('input, textarea, select, [contenteditable="true"], [contenteditable="plaintext-only"]'));
};

type CommitNumberInputProps = {
  value: number;
  onCommit: (value: number) => void;
  min?: number;
  max?: number;
  step?: number;
  width?: number;
};

const COMPARE_HEADER_ROW_COUNT = 3;
const EMPTY_MERGE_ROW_OPS = new Map<number, SaveMergeRowOp>();
const EMPTY_MERGE_COL_OPS = new Map<number, SaveMergeColOp>();

const buildCompareRequestSignature = (request: {
  basePath: string;
  oursPath: string;
  theirsPath: string;
  compareMode: ThreeWayDiffRequest['compareMode'];
  primaryKeyCol: number;
  rowSimilarityThreshold: number;
}) =>
  JSON.stringify({
    ...request,
    compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
  });

const cloneMergeSheetsDeep = (sheets: MergeSheetData[]): MergeSheetData[] =>
  sheets.map((sheet) => ({
    ...sheet,
    cells: (sheet.cells ?? []).map((cell) => ({ ...cell })),
    rowsMeta: sheet.rowsMeta ? sheet.rowsMeta.map((meta) => ({ ...meta })) : sheet.rowsMeta,
    columnsMeta: sheet.columnsMeta ? sheet.columnsMeta.map((meta) => ({ ...meta })) : sheet.columnsMeta,
  }));

const cloneResolvedBySheetMap = (source: Map<number, Set<string>>): Map<number, Set<string>> => {
  const next = new Map<number, Set<string>>();
  source.forEach((value, key) => {
    next.set(key, new Set(value));
  });
  return next;
};

const cloneRowOpsBySheetMap = (
  source: Map<number, Map<number, SaveMergeRowOp>>,
): Map<number, Map<number, SaveMergeRowOp>> => {
  const next = new Map<number, Map<number, SaveMergeRowOp>>();
  source.forEach((opsMap, sheetIndex) => {
    const clonedOps = new Map<number, SaveMergeRowOp>();
    opsMap.forEach((op, visualRowNumber) => {
      clonedOps.set(visualRowNumber, {
        ...op,
        values: op.values ? [...op.values] : op.values,
      });
    });
    next.set(sheetIndex, clonedOps);
  });
  return next;
};

const cloneColOpsBySheetMap = (
  source: Map<number, Map<number, SaveMergeColOp>>,
): Map<number, Map<number, SaveMergeColOp>> => {
  const next = new Map<number, Map<number, SaveMergeColOp>>();
  source.forEach((opsMap, sheetIndex) => {
    const clonedOps = new Map<number, SaveMergeColOp>();
    opsMap.forEach((op, alignedColNumber) => {
      clonedOps.set(alignedColNumber, {
        ...op,
        values: op.values ? [...op.values] : op.values,
      });
    });
    next.set(sheetIndex, clonedOps);
  });
  return next;
};

const CommitNumberInput: React.FC<CommitNumberInputProps> = React.memo(
  ({ value, onCommit, min, max, step = 1, width = 60 }) => {
    const [draft, setDraft] = useState<string>(String(value));

    useEffect(() => {
      setDraft(String(value));
    }, [value]);

    const commit = useCallback(() => {
      const trimmed = draft.trim();
      let nextValue = trimmed === '' ? (min ?? 0) : Number(trimmed);
      if (!Number.isFinite(nextValue)) {
        setDraft(String(value));
        return;
      }
      if (typeof min === 'number') nextValue = Math.max(min, nextValue);
      if (typeof max === 'number') nextValue = Math.min(max, nextValue);
      if (step >= 1) nextValue = Math.floor(nextValue);
      setDraft(String(nextValue));
      if (nextValue !== value) {
        onCommit(nextValue);
      }
    }, [draft, max, min, onCommit, step, value]);

    return (
      <input
        type="number"
        min={min}
        max={max}
        step={step}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
            e.currentTarget.blur();
            return;
          }
          if (e.key === 'Escape') {
            e.preventDefault();
            setDraft(String(value));
            e.currentTarget.blur();
          }
        }}
        style={{ width, padding: '2px 6px', boxSizing: 'border-box' }}
      />
    );
  },
);

export const App: React.FC = () => {
  const [mode, setMode] = useState<ViewMode>('diff');

  // 单文件编辑状态
  const [filePath, setFilePath] = useState<string | null>(null);
  const [sheetName, setSheetName] = useState<string | null>(null);
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [rows, setRows] = useState<SheetCell[][]>([]);
  const [changes, setChanges] = useState<Map<string, CellChange>>(new Map());
  const [saving, setSaving] = useState(false);
  // 当前单文件模式下选中的单元格（用于顶部“公式栏”显示）
  const [selectedSingleCell, setSelectedSingleCell] = useState<SheetCell | null>(null);
  // 固定在顶部的首行数，默认 3 行
  const [frozenRowCount, setFrozenRowCount] = useState<number>(3);
  // 固定在左侧的列数（不含最左侧行号列），默认 0 列
  const [frozenColCount, setFrozenColCount] = useState<number>(0);
  // merge/diff 视图中固定在顶部展示的行数，默认 0 行（仅影响显示，不参与比对）
  const [mergeFrozenRowCount, setMergeFrozenRowCount] = useState<number>(0);
  const [mergeFrozenRowDraft, setMergeFrozenRowDraft] = useState<string>('0');
  const [rowSimilarityThreshold, setRowSimilarityThreshold] = useState<number>(0.9);

  // 三方 diff 状态
  const [mergeSheets, setMergeSheets] = useState<MergeSheetData[]>([]);
  const [selectedMergeSheetIndex, setSelectedMergeSheetIndex] = useState<number>(0);
  const [mergeCells, setMergeCells] = useState<MergeCell[]>([]);
  const [mergeRowsMeta, setMergeRowsMeta] = useState<MergeRowMeta[]>([]);
  const [mergeColumnsMeta, setMergeColumnsMeta] = useState<MergeColumnMeta[]>([]);
  const [primaryKeyMode, setPrimaryKeyMode] = useState<PrimaryKeyMode>('auto');
  const [primaryKeyCol, setPrimaryKeyCol] = useState<number>(1);
  const [autoHasPrimaryKey, setAutoHasPrimaryKey] = useState<boolean>(false);
  const [primaryKeyHint, setPrimaryKeyHint] = useState<string>('');
  // 记录“用户已确认合并”的单元格（resolved），按 sheetIndex 分组，key="row:col"（1-based）
  const [resolvedBySheet, setResolvedBySheet] = useState<Map<number, Set<string>>>(new Map());
  const [mergeRowOpsBySheet, setMergeRowOpsBySheet] = useState<Map<number, Map<number, SaveMergeRowOp>>>(new Map());
  const [mergeColOpsBySheet, setMergeColOpsBySheet] = useState<Map<number, Map<number, SaveMergeColOp>>>(new Map());
  const [mergedPreviewMinRows, setMergedPreviewMinRows] = useState<number>(5);
  const [mergedPreviewRows, setMergedPreviewRows] = useState<(string | number | null)[][]>([]);
  const [mergedPreviewRowVisuals, setMergedPreviewRowVisuals] = useState<(number | null)[]>([]);
  const [mergedPreviewAlignedCols, setMergedPreviewAlignedCols] = useState<number[]>([]);
  const [mergeThreeWayRows, setMergeThreeWayRows] = useState<ThreeWayRowResult[]>([]);
  const [showFullTables, setShowFullTables] = useState<boolean>(false);
  const [fullOursRows, setFullOursRows] = useState<(string | number | null)[][]>([]);
  const [fullTheirsRows, setFullTheirsRows] = useState<(string | number | null)[][]>([]);
  const [mergeInfo, setMergeInfo] = useState<{
    basePath: string;
    oursPath: string;
    theirsPath: string;
    sheetName: string;
  } | null>(null);
  const [selectedMergePaths, setSelectedMergePaths] = useState<{
    basePath: string | null;
    oursPath: string | null;
    theirsPath: string | null;
  }>({
    basePath: null,
    oursPath: null,
    theirsPath: null,
  });
  const [mergePathInputs, setMergePathInputs] = useState<{
    basePath: string;
    oursPath: string;
    theirsPath: string;
  }>({
    basePath: '',
    oursPath: '',
    theirsPath: '',
  });
  const [diffLeftWorkbook, setDiffLeftWorkbook] = useState<OpenResult | null>(null);
  const [diffRightWorkbook, setDiffRightWorkbook] = useState<OpenResult | null>(null);
  const [diffPathInputs, setDiffPathInputs] = useState<{
    left: string;
    right: string;
  }>({
    left: '',
    right: '',
  });
  const [diffFileSelectorCollapsed, setDiffFileSelectorCollapsed] = useState<boolean>(false);
  const [mergeFileSelectorCollapsed, setMergeFileSelectorCollapsed] = useState<boolean>(false);
  const [diffAdvancedCollapsed, setDiffAdvancedCollapsed] = useState<boolean>(false);
  const [mergeAdvancedCollapsed, setMergeAdvancedCollapsed] = useState<boolean>(false);
  const [diffAnalyzeInProgress, setDiffAnalyzeInProgress] = useState<boolean>(false);
  const [diffAnalyzeProgress, setDiffAnalyzeProgress] = useState<number>(0);
  const [mergeAnalyzeInProgress, setMergeAnalyzeInProgress] = useState<boolean>(false);
  const [mergeAnalyzeProgress, setMergeAnalyzeProgress] = useState<number>(0);
  const [diffSheets, setDiffSheets] = useState<MergeSheetData[]>([]);
  const [selectedDiffSheetIndex, setSelectedDiffSheetIndex] = useState<number>(0);
  const [diffSelectedCell, setDiffSelectedCell] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [diffLeftChangesBySheet, setDiffLeftChangesBySheet] = useState<Map<string, Map<string, CellChange>>>(
    new Map(),
  );
  const [diffRightChangesBySheet, setDiffRightChangesBySheet] = useState<Map<string, Map<string, CellChange>>>(
    new Map(),
  );
  const [diffLeftRowOpsBySheet, setDiffLeftRowOpsBySheet] = useState<Map<string, SaveMergeRowOp[]>>(
    new Map(),
  );
  const [diffRightRowOpsBySheet, setDiffRightRowOpsBySheet] = useState<Map<string, SaveMergeRowOp[]>>(
    new Map(),
  );
  const [diffUndoStack, setDiffUndoStack] = useState<DiffHistorySnapshot[]>([]);
  const [diffRedoStack, setDiffRedoStack] = useState<DiffHistorySnapshot[]>([]);
  const [diffSavingSide, setDiffSavingSide] = useState<DiffSide | null>(null);
  const [cliInfo, setCliInfo] = useState<CliThreeWayInfo | null>(null);
  const [selectedMergeCell, setSelectedMergeCell] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [mergeUndoStack, setMergeUndoStack] = useState<MergeHistorySnapshot[]>([]);
  const debugSeqRef = useRef(0);
  const cliInfoInitializedRef = useRef(false);
  const lastMergeCompareSignatureRef = useRef<string | null>(null);
  const lastDiffCompareSignatureRef = useRef<string | null>(null);
  const nextDebugRequestId = useCallback(
    (prefix: string) => `${prefix}-${Date.now()}-${++debugSeqRef.current}`,
    [],
  );
  const logRendererDebug = useCallback((event: string, details?: Record<string, unknown>) => {
    try {
      window.excelAPI.debugLog({ source: 'renderer', event, details });
    } catch {
      // 忽略日志失败，避免影响主流程
    }
  }, []);
  const requestedPrimaryKeyCol = useMemo(() => {
    if (primaryKeyMode === 'manual') return Math.max(1, Math.floor(primaryKeyCol || 1));
    if (primaryKeyMode === 'none') return 0;
    return -1;
  }, [primaryKeyMode, primaryKeyCol]);
  const displayPrimaryKeyCol = useMemo(() => {
    const currentMergeSheet = mergeSheets[selectedMergeSheetIndex] ?? null;
    const alignedCol = currentMergeSheet?.primaryKeyAlignedCol;
    return typeof alignedCol === 'number' && alignedCol >= 1 ? alignedCol : undefined;
  }, [mergeSheets, selectedMergeSheetIndex]);
  const compareMode = useMemo(() => (mode === 'diff' ? 'diff' : 'merge'), [mode]);
  const parsedMergeFrozenRowDraft = useMemo(() => {
    const trimmed = mergeFrozenRowDraft.trim();
    if (trimmed === '') return null;
    const parsed = Number(trimmed);
    if (!Number.isFinite(parsed)) return null;
    return Math.max(0, Math.floor(parsed));
  }, [mergeFrozenRowDraft]);
  const canRefreshFrozenRows =
    parsedMergeFrozenRowDraft !== null && parsedMergeFrozenRowDraft !== mergeFrozenRowCount;

  useEffect(() => {
    if (mergeFrozenRowCount >= 0) return;
    setMergeFrozenRowCount(0);
  }, [mergeFrozenRowCount]);

  useEffect(() => {
    setMergeFrozenRowDraft(String(mergeFrozenRowCount));
    logRendererDebug('frozen-rows:applied-value-sync', {
      appliedValue: mergeFrozenRowCount,
    });
  }, [mergeFrozenRowCount, logRendererDebug]);

  const applyMergeFrozenRowDraft = useCallback(() => {
    const requestId = nextDebugRequestId('frozen-rows-refresh');
    if (parsedMergeFrozenRowDraft === null) {
      logRendererDebug('frozen-rows:refresh-invalid', {
        requestId,
        draftValue: mergeFrozenRowDraft,
        appliedValue: mergeFrozenRowCount,
        mode,
      });
      setMergeFrozenRowDraft(String(mergeFrozenRowCount));
      return;
    }
    logRendererDebug('frozen-rows:refresh-click', {
      requestId,
      draftValue: mergeFrozenRowDraft,
      nextValue: parsedMergeFrozenRowDraft,
      appliedValue: mergeFrozenRowCount,
      mode,
    });
    if (parsedMergeFrozenRowDraft !== mergeFrozenRowCount) {
      setMergeFrozenRowCount(parsedMergeFrozenRowDraft);
      return;
    }
    setMergeFrozenRowDraft(String(mergeFrozenRowCount));
  }, [
    logRendererDebug,
    mergeFrozenRowCount,
    mergeFrozenRowDraft,
    mode,
    nextDebugRequestId,
    parsedMergeFrozenRowDraft,
  ]);
  const handleMergeFrozenRowDraftChange = useCallback(
    (nextValue: string, targetMode: ViewMode) => {
      const startedAt = performance.now();
      const trimmed = nextValue.trim();
      let normalizedDraft = '';
      if (trimmed !== '') {
        const parsed = Number(trimmed);
        if (Number.isFinite(parsed)) {
          normalizedDraft = String(Math.max(0, Math.floor(parsed)));
        }
      }
      setMergeFrozenRowDraft(normalizedDraft);
      logRendererDebug('frozen-rows:draft-change', {
        targetMode,
        draftValue: normalizedDraft,
        appliedValue: mergeFrozenRowCount,
      });
      requestAnimationFrame(() => {
        logRendererDebug('frozen-rows:draft-change-painted', {
          targetMode,
          draftValue: normalizedDraft,
          appliedValue: mergeFrozenRowCount,
          elapsedMs: Math.round(performance.now() - startedAt),
        });
      });
    },
    [logRendererDebug, mergeFrozenRowCount],
  );

  const normalizeCompareSheets = useCallback(
    (result: { sheet?: MergeSheetData; sheets?: MergeSheetData[] } | null | undefined) =>
      result?.sheets && result.sheets.length > 0
        ? result.sheets
        : result?.sheet
          ? [result.sheet]
          : [],
    [],
  );
  const buildDefaultResolvedBySheet = useCallback((sheets: MergeSheetData[]) => {
    const resolved = new Map<number, Set<string>>();
    sheets.forEach((sheet, sheetIndex) => {
      const set = new Set<string>();
      (sheet.cells ?? []).forEach((cell) => {
        if (cell.status !== 'conflict') {
          set.add(`${cell.row}:${cell.col}`);
        }
      });
      resolved.set(sheetIndex, set);
    });
    return resolved;
  }, []);
  const captureCurrentMergeSnapshot = useCallback(
    (): MergeHistorySnapshot => ({
      mergeSheets: cloneMergeSheetsDeep(mergeSheets),
      resolvedBySheet: cloneResolvedBySheetMap(resolvedBySheet),
      mergeRowOpsBySheet: cloneRowOpsBySheetMap(mergeRowOpsBySheet),
      mergeColOpsBySheet: cloneColOpsBySheetMap(mergeColOpsBySheet),
      selectedMergeSheetIndex,
      selectedMergeCell: selectedMergeCell ? { ...selectedMergeCell } : null,
    }),
    [
      mergeSheets,
      resolvedBySheet,
      mergeRowOpsBySheet,
      mergeColOpsBySheet,
      selectedMergeSheetIndex,
      selectedMergeCell,
    ],
  );
  const pushMergeUndoSnapshot = useCallback(() => {
    setMergeUndoStack((prev) => {
      const next = [...prev, captureCurrentMergeSnapshot()];
      return next.length > 40 ? next.slice(next.length - 40) : next;
    });
  }, [captureCurrentMergeSnapshot]);
  const resetMergeUndoStack = useCallback(() => {
    setMergeUndoStack([]);
  }, []);
  const handleUndoMergeAction = useCallback(() => {
    if (mergeUndoStack.length === 0) return;
    const snapshot = mergeUndoStack[mergeUndoStack.length - 1];
    setMergeUndoStack((prev) => prev.slice(0, -1));
    const restoredSheets = cloneMergeSheetsDeep(snapshot.mergeSheets);
    const safeSheetIndex = Math.min(
      snapshot.selectedMergeSheetIndex,
      Math.max(0, restoredSheets.length - 1),
    );
    const activeSheet = restoredSheets[safeSheetIndex] ?? null;
    setMergeSheets(restoredSheets);
    setSelectedMergeSheetIndex(safeSheetIndex);
    setMergeCells(activeSheet?.cells ?? []);
    setMergeRowsMeta(activeSheet?.rowsMeta ?? []);
    setMergeColumnsMeta(activeSheet?.columnsMeta ?? []);
    setResolvedBySheet(cloneResolvedBySheetMap(snapshot.resolvedBySheet));
    setMergeRowOpsBySheet(cloneRowOpsBySheetMap(snapshot.mergeRowOpsBySheet));
    setMergeColOpsBySheet(cloneColOpsBySheetMap(snapshot.mergeColOpsBySheet));
    setSelectedMergeCell(snapshot.selectedMergeCell ? { ...snapshot.selectedMergeCell } : null);
    setMergeInfo((prev) =>
      prev
        ? {
            ...prev,
            sheetName: activeSheet?.sheetName ?? prev.sheetName,
          }
        : prev,
    );
  }, [mergeUndoStack]);
  const captureCurrentDiffSnapshot = useCallback(
    (): DiffHistorySnapshot => ({
      leftWorkbook: diffLeftWorkbook,
      rightWorkbook: diffRightWorkbook,
      leftChangesBySheet: diffLeftChangesBySheet,
      rightChangesBySheet: diffRightChangesBySheet,
      leftRowOpsBySheet: diffLeftRowOpsBySheet,
      rightRowOpsBySheet: diffRightRowOpsBySheet,
      diffSheets,
      selectedDiffSheetIndex,
      diffSelectedCell,
    }),
    [
      diffLeftWorkbook,
      diffRightWorkbook,
      diffLeftChangesBySheet,
      diffRightChangesBySheet,
      diffLeftRowOpsBySheet,
      diffRightRowOpsBySheet,
      diffSheets,
      selectedDiffSheetIndex,
      diffSelectedCell,
    ],
  );
  const restoreDiffSnapshot = useCallback((snapshot: DiffHistorySnapshot) => {
    setDiffLeftWorkbook(snapshot.leftWorkbook);
    setDiffRightWorkbook(snapshot.rightWorkbook);
    setDiffLeftChangesBySheet(snapshot.leftChangesBySheet);
    setDiffRightChangesBySheet(snapshot.rightChangesBySheet);
    setDiffLeftRowOpsBySheet(snapshot.leftRowOpsBySheet);
    setDiffRightRowOpsBySheet(snapshot.rightRowOpsBySheet);
    setDiffSheets(snapshot.diffSheets);
    setSelectedDiffSheetIndex(snapshot.selectedDiffSheetIndex);
    setDiffSelectedCell(snapshot.diffSelectedCell);
  }, []);
  const pushDiffHistory = useCallback(
    (snapshot?: DiffHistorySnapshot) => {
      const nextSnapshot = snapshot ?? captureCurrentDiffSnapshot();
      setDiffUndoStack((prev) => {
        const next = [...prev, nextSnapshot];
        return next.length > 30 ? next.slice(next.length - 30) : next;
      });
      setDiffRedoStack([]);
    },
    [captureCurrentDiffSnapshot],
  );
  const resetDiffHistory = useCallback(() => {
    setDiffUndoStack([]);
    setDiffRedoStack([]);
  }, []);
  const handleDiffUndo = useCallback(() => {
    if (diffUndoStack.length === 0) return;
    const snapshot = diffUndoStack[diffUndoStack.length - 1];
    const currentSnapshot = captureCurrentDiffSnapshot();
    setDiffUndoStack((prev) => prev.slice(0, -1));
    setDiffRedoStack((prev) => {
      const next = [...prev, currentSnapshot];
      return next.length > 30 ? next.slice(next.length - 30) : next;
    });
    restoreDiffSnapshot(snapshot);
  }, [captureCurrentDiffSnapshot, diffUndoStack, restoreDiffSnapshot]);
  const handleDiffRedo = useCallback(() => {
    if (diffRedoStack.length === 0) return;
    const snapshot = diffRedoStack[diffRedoStack.length - 1];
    const currentSnapshot = captureCurrentDiffSnapshot();
    setDiffRedoStack((prev) => prev.slice(0, -1));
    setDiffUndoStack((prev) => {
      const next = [...prev, currentSnapshot];
      return next.length > 30 ? next.slice(next.length - 30) : next;
    });
    restoreDiffSnapshot(snapshot);
  }, [captureCurrentDiffSnapshot, diffRedoStack, restoreDiffSnapshot]);

  const applyMergeComparisonResult = useCallback(
    (
      result: { basePath: string; oursPath: string; theirsPath: string; sheet?: MergeSheetData; sheets?: MergeSheetData[] },
      preferredSheetName?: string,
    ) => {
      const allMergeSheets = normalizeCompareSheets(result);
      const preferredIndex = preferredSheetName
        ? allMergeSheets.findIndex((sheet) => sheet.sheetName === preferredSheetName)
        : -1;
      const nextIndex =
        preferredIndex >= 0 ? preferredIndex : Math.min(selectedMergeSheetIndex, Math.max(0, allMergeSheets.length - 1));

      setMode('merge');
      setSelectedSingleCell(null);
      setMergeSheets(allMergeSheets);
      setSelectedMergeSheetIndex(nextIndex);
      setMergeCells(allMergeSheets[nextIndex]?.cells ?? []);
      setMergeRowsMeta(allMergeSheets[nextIndex]?.rowsMeta ?? []);
      setMergeColumnsMeta(allMergeSheets[nextIndex]?.columnsMeta ?? []);
      setAutoHasPrimaryKey(false);
      setPrimaryKeyHint('');
      setResolvedBySheet(buildDefaultResolvedBySheet(allMergeSheets));
      setMergeRowOpsBySheet(new Map());
      setMergeColOpsBySheet(new Map());
      setMergedPreviewRows([]);
      setMergedPreviewRowVisuals([]);
      setMergeInfo({
        basePath: result.basePath,
        oursPath: result.oursPath,
        theirsPath: result.theirsPath,
        sheetName: allMergeSheets[nextIndex]?.sheetName ?? allMergeSheets[0]?.sheetName ?? '',
      });
      setSelectedMergePaths({
        basePath: result.basePath,
        oursPath: result.oursPath,
        theirsPath: result.theirsPath,
      });
      setMergePathInputs({
        basePath: result.basePath,
        oursPath: result.oursPath,
        theirsPath: result.theirsPath,
      });
      setSelectedMergeCell(null);
      resetMergeUndoStack();
    },
    [buildDefaultResolvedBySheet, normalizeCompareSheets, resetMergeUndoStack, selectedMergeSheetIndex],
  );

  const loadMergeComparison = useCallback(
    async (
      paths: {
        basePath: string;
        oursPath: string;
        theirsPath: string;
      },
      preferredSheetName?: string,
    ) => {
      const requestId = nextDebugRequestId('load-merge');
      lastMergeCompareSignatureRef.current = buildCompareRequestSignature({
        basePath: paths.basePath,
        oursPath: paths.oursPath,
        theirsPath: paths.theirsPath,
        compareMode: 'merge',
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      const req: ThreeWayDiffRequest = {
        basePath: paths.basePath,
        oursPath: paths.oursPath,
        theirsPath: paths.theirsPath,
        compareMode: 'merge',
        primaryKeyCol: requestedPrimaryKeyCol,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        rowSimilarityThreshold,
        debugRequestId: requestId,
      };
      const startedAt = performance.now();
      logRendererDebug('loadMergeComparison:start', {
        requestId,
        preferredSheetName: preferredSheetName ?? null,
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      const result = await window.excelAPI.computeThreeWayDiff(req);
      const durationMs = Math.round(performance.now() - startedAt);
      if (!result) {
        logRendererDebug('loadMergeComparison:null', { requestId, durationMs });
        return;
      }
      logRendererDebug('loadMergeComparison:end', {
        requestId,
        durationMs,
        sheetCount: normalizeCompareSheets(result).length,
      });
      applyMergeComparisonResult(result, preferredSheetName);
    },
    [
      applyMergeComparisonResult,
      logRendererDebug,
      requestedPrimaryKeyCol,
      nextDebugRequestId,
      normalizeCompareSheets,
      rowSimilarityThreshold,
    ],
  );

  const loadDiffComparison = useCallback(
    async (leftWorkbook: OpenResult, rightWorkbook: OpenResult, preferredSheetName?: string) => {
      const requestId = nextDebugRequestId('load-diff');
      lastDiffCompareSignatureRef.current = buildCompareRequestSignature({
        basePath: leftWorkbook.filePath,
        oursPath: leftWorkbook.filePath,
        theirsPath: rightWorkbook.filePath,
        compareMode: 'diff',
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      const req: ThreeWayDiffRequest = {
        basePath: leftWorkbook.filePath,
        oursPath: leftWorkbook.filePath,
        theirsPath: rightWorkbook.filePath,
        compareMode: 'diff',
        primaryKeyCol: requestedPrimaryKeyCol,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        rowSimilarityThreshold,
        debugRequestId: requestId,
      };
      const startedAt = performance.now();
      logRendererDebug('loadDiffComparison:start', {
        requestId,
        preferredSheetName: preferredSheetName ?? null,
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      const result = await window.excelAPI.computeThreeWayDiff(req);
      const durationMs = Math.round(performance.now() - startedAt);
      if (!result) {
        logRendererDebug('loadDiffComparison:null', { requestId, durationMs });
        return;
      }

      const allDiffSheets = normalizeCompareSheets(result);
      const preferredIndex = preferredSheetName
        ? allDiffSheets.findIndex((sheet) => sheet.sheetName === preferredSheetName)
        : -1;
      const nextIndex =
        preferredIndex >= 0 ? preferredIndex : Math.min(selectedDiffSheetIndex, Math.max(0, allDiffSheets.length - 1));

      setMode('diff');
      setDiffLeftWorkbook(leftWorkbook);
      setDiffRightWorkbook(rightWorkbook);
      setDiffPathInputs({
        left: leftWorkbook.filePath,
        right: rightWorkbook.filePath,
      });
      setDiffSheets(allDiffSheets);
      setSelectedDiffSheetIndex(nextIndex);
      setDiffSelectedCell(null);
      resetDiffHistory();
      logRendererDebug('loadDiffComparison:end', {
        requestId,
        durationMs,
        sheetCount: allDiffSheets.length,
      });
    },
    [
      logRendererDebug,
      nextDebugRequestId,
      normalizeCompareSheets,
      requestedPrimaryKeyCol,
      resetDiffHistory,
      rowSimilarityThreshold,
      selectedDiffSheetIndex,
    ],
  );

  const handlePickDiffWorkbook = useCallback(
    async (side: DiffSide) => {
      const result: OpenResult | null = await window.excelAPI.openFile();
      if (!result) return;
      setDiffPathInputs((prev) => ({
        ...prev,
        [side]: result.filePath,
      }));
      setMode('diff');
    },
    [],
  );

  const handleLoadDiffFromInputs = useCallback(async () => {
    if (diffAnalyzeInProgress) return;
    const rawLeft = diffPathInputs.left.trim();
    const rawRight = diffPathInputs.right.trim();
    if (!rawLeft || !rawRight) {
      alert('请先填写左侧和右侧 Excel 文件路径。');
      return;
    }
    setMode('diff');
    setDiffAnalyzeInProgress(true);
    setDiffAnalyzeProgress(5);

    let animatedProgress = 5;
    const timer = window.setInterval(() => {
      animatedProgress = Math.min(92, animatedProgress + Math.random() * 4 + 1);
      setDiffAnalyzeProgress(animatedProgress);
    }, 160);

    let success = false;
    try {
      const leftWorkbook = await window.excelAPI.loadWorkbook(rawLeft);
      if (!leftWorkbook) {
        alert(`无法打开左侧 Excel：${rawLeft}`);
        return;
      }
      setDiffAnalyzeProgress((prev) => Math.max(prev, 30));

      const rightWorkbook = await window.excelAPI.loadWorkbook(rawRight);
      if (!rightWorkbook) {
        alert(`无法打开右侧 Excel：${rawRight}`);
        return;
      }
      setDiffAnalyzeProgress((prev) => Math.max(prev, 55));

      setDiffPathInputs({
        left: leftWorkbook.filePath,
        right: rightWorkbook.filePath,
      });
      await loadDiffComparison(leftWorkbook, rightWorkbook, diffSheets[selectedDiffSheetIndex]?.sheetName);
      setDiffAnalyzeProgress(100);
      success = true;
    } finally {
      window.clearInterval(timer);
      setDiffAnalyzeInProgress(false);
      if (!success) {
        setDiffAnalyzeProgress(0);
      }
    }
  }, [diffAnalyzeInProgress, diffPathInputs, diffSheets, selectedDiffSheetIndex, loadDiffComparison]);

  const handlePickMergeWorkbook = useCallback(
    async (role: 'basePath' | 'oursPath' | 'theirsPath') => {
      const result: OpenResult | null = await window.excelAPI.openFile();
      if (!result) return;
      setMergePathInputs((prev) => ({
        ...prev,
        [role]: result.filePath,
      }));
      const nextPaths = {
        ...selectedMergePaths,
        [role]: result.filePath,
      };
      setMode('merge');
      setSelectedMergePaths(nextPaths);
    },
    [selectedMergePaths],
  );

  const handleLoadMergeFromInputs = useCallback(async () => {
    if (mergeAnalyzeInProgress) return;
    const rawBase = mergePathInputs.basePath.trim();
    const rawOurs = mergePathInputs.oursPath.trim();
    const rawTheirs = mergePathInputs.theirsPath.trim();
    if (!rawBase || !rawOurs || !rawTheirs) {
      alert('请先填写 base / ours / theirs 的文件路径。');
      return;
    }

    setMode('merge');
    setMergeAnalyzeInProgress(true);
    setMergeAnalyzeProgress(5);

    let animatedProgress = 5;
    const timer = window.setInterval(() => {
      animatedProgress = Math.min(92, animatedProgress + Math.random() * 4 + 1);
      setMergeAnalyzeProgress(animatedProgress);
    }, 160);

    let success = false;
    try {
      const baseWorkbook = await window.excelAPI.loadWorkbook(rawBase);
      if (!baseWorkbook) {
        alert(`无法打开 base Excel：${rawBase}`);
        return;
      }
      setMergeAnalyzeProgress((prev) => Math.max(prev, 20));

      const oursWorkbook = await window.excelAPI.loadWorkbook(rawOurs);
      if (!oursWorkbook) {
        alert(`无法打开 ours Excel：${rawOurs}`);
        return;
      }
      setMergeAnalyzeProgress((prev) => Math.max(prev, 35));

      const theirsWorkbook = await window.excelAPI.loadWorkbook(rawTheirs);
      if (!theirsWorkbook) {
        alert(`无法打开 theirs Excel：${rawTheirs}`);
        return;
      }
      setMergeAnalyzeProgress((prev) => Math.max(prev, 50));

      const normalizedPaths = {
        basePath: baseWorkbook.filePath,
        oursPath: oursWorkbook.filePath,
        theirsPath: theirsWorkbook.filePath,
      };
      setSelectedMergePaths(normalizedPaths);
      setMergePathInputs(normalizedPaths);
      await loadMergeComparison(normalizedPaths, mergeInfo?.sheetName);
      setMergeAnalyzeProgress(100);
      success = true;
    } finally {
      window.clearInterval(timer);
      setMergeAnalyzeInProgress(false);
      if (!success) {
        setMergeAnalyzeProgress(0);
      }
    }
  }, [mergeAnalyzeInProgress, mergePathInputs, loadMergeComparison, mergeInfo?.sheetName]);

  useEffect(() => {
    if (cliInfoInitializedRef.current) return;
    cliInfoInitializedRef.current = true;
    (async () => {
      try {
        const info = await window.excelAPI.getCliThreeWayInfo();
        if (!info) return;
        setCliInfo(info);
        if (info.mode === 'merge') {
          await loadMergeComparison(
            {
              basePath: info.basePath,
              oursPath: info.oursPath,
              theirsPath: info.theirsPath,
            },
          );
          return;
        }
        const [leftWorkbook, rightWorkbook] = await Promise.all([
          window.excelAPI.loadWorkbook(info.oursPath),
          window.excelAPI.loadWorkbook(info.theirsPath),
        ]);
        if (!leftWorkbook || !rightWorkbook) return;
        await loadDiffComparison(leftWorkbook, rightWorkbook);
      } catch {
        // 忽略错误，保持交互式模式可用
      }
    })();
  }, [loadDiffComparison, loadMergeComparison]);
  useEffect(() => {
    if (mode !== 'diff') return;
    const handleWindowKeyDown = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      if (isEditableEventTarget(e.target) || isEditableEventTarget(document.activeElement)) return;
      const key = e.key.toLowerCase();
      if (key === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleDiffRedo();
        else handleDiffUndo();
        return;
      }
      if (key === 'y') {
        e.preventDefault();
        handleDiffRedo();
      }
    };
    window.addEventListener('keydown', handleWindowKeyDown);
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown);
    };
  }, [handleDiffRedo, handleDiffUndo, mode]);
  useEffect(() => {
    if (mode !== 'merge') return;
    const handleWindowKeyDown = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
      if (isEditableEventTarget(e.target) || isEditableEventTarget(document.activeElement)) return;
      const key = e.key.toLowerCase();
      if (key === 'z') {
        e.preventDefault();
        handleUndoMergeAction();
      }
    };
    window.addEventListener('keydown', handleWindowKeyDown);
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown);
    };
  }, [handleUndoMergeAction, mode]);

  // 当主键设置变化时，重新计算 merge 结果（避免重开文件）
  useEffect(() => {
    if (mode !== 'merge' || !mergeInfo) return;
    const compareSignature = buildCompareRequestSignature({
      basePath: mergeInfo.basePath,
      oursPath: mergeInfo.oursPath,
      theirsPath: mergeInfo.theirsPath,
      compareMode,
      primaryKeyCol: requestedPrimaryKeyCol,
      rowSimilarityThreshold,
    });
    if (lastMergeCompareSignatureRef.current === compareSignature) {
      logRendererDebug('mergeEffect:compute-skip', {
        reason: 'signature-unchanged',
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      return;
    }
    lastMergeCompareSignatureRef.current = compareSignature;
    let cancelled = false;
    (async () => {
      const requestId = nextDebugRequestId('merge-effect');
      const req: ThreeWayDiffRequest = {
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        compareMode,
        primaryKeyCol: requestedPrimaryKeyCol,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        rowSimilarityThreshold,
        debugRequestId: requestId,
      };
      const startedAt = performance.now();
      logRendererDebug('mergeEffect:compute-start', {
        requestId,
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      try {
        const result = await window.excelAPI.computeThreeWayDiff(req);
        const durationMs = Math.round(performance.now() - startedAt);
        if (!result) {
          logRendererDebug('mergeEffect:compute-null', { requestId, durationMs });
          return;
        }
        if (cancelled) {
          logRendererDebug('mergeEffect:compute-cancelled-after-result', { requestId, durationMs });
          return;
        }
        const allMergeSheets =
          result.sheets && result.sheets.length > 0
            ? result.sheets
            : result.sheet
              ? [result.sheet]
              : [];
        const nextIndex = Math.min(selectedMergeSheetIndex, Math.max(0, allMergeSheets.length - 1));
        setMergeSheets(allMergeSheets);
        setSelectedMergeSheetIndex(nextIndex);
        setMergeCells(allMergeSheets[nextIndex]?.cells ?? []);
        setMergeRowsMeta(allMergeSheets[nextIndex]?.rowsMeta ?? []);
        setMergeColumnsMeta(allMergeSheets[nextIndex]?.columnsMeta ?? []);
        setResolvedBySheet(buildDefaultResolvedBySheet(allMergeSheets));
        setMergeRowOpsBySheet(new Map());
        setMergeColOpsBySheet(new Map());
        setMergedPreviewRows([]);
        setMergedPreviewRowVisuals([]);
        setSelectedMergeCell(null);
        resetMergeUndoStack();
        logRendererDebug('mergeEffect:compute-end', {
          requestId,
          durationMs,
          sheetCount: allMergeSheets.length,
          activeSheet: allMergeSheets[nextIndex]?.sheetName ?? null,
          diffCellCount: allMergeSheets[nextIndex]?.cells?.length ?? 0,
        });
      } catch (error) {
        logRendererDebug('mergeEffect:compute-error', {
          requestId,
          message: error instanceof Error ? error.message : String(error),
        });
        console.error(error);
      }
    })();
    return () => {
      cancelled = true;
      logRendererDebug('mergeEffect:compute-cancel-request');
    };
  }, [
    logRendererDebug,
    requestedPrimaryKeyCol,
    rowSimilarityThreshold,
    mergeInfo?.basePath,
    mergeInfo?.oursPath,
    mergeInfo?.theirsPath,
    mode,
    compareMode,
    nextDebugRequestId,
    buildDefaultResolvedBySheet,
    resetMergeUndoStack,
  ]);

  // 当主键设置变化时，重新计算 diff 结果（避免重开文件）
  useEffect(() => {
    if (mode !== 'diff' || !diffLeftWorkbook || !diffRightWorkbook) return;
    const compareSignature = buildCompareRequestSignature({
      basePath: diffLeftWorkbook.filePath,
      oursPath: diffLeftWorkbook.filePath,
      theirsPath: diffRightWorkbook.filePath,
      compareMode: 'diff',
      primaryKeyCol: requestedPrimaryKeyCol,
      rowSimilarityThreshold,
    });
    if (lastDiffCompareSignatureRef.current === compareSignature) {
      logRendererDebug('diffEffect:compute-skip', {
        reason: 'signature-unchanged',
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      return;
    }
    lastDiffCompareSignatureRef.current = compareSignature;
    let cancelled = false;
    (async () => {
      const requestId = nextDebugRequestId('diff-effect');
      const req: ThreeWayDiffRequest = {
        basePath: diffLeftWorkbook.filePath,
        oursPath: diffLeftWorkbook.filePath,
        theirsPath: diffRightWorkbook.filePath,
        compareMode: 'diff',
        primaryKeyCol: requestedPrimaryKeyCol,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        rowSimilarityThreshold,
        debugRequestId: requestId,
      };
      const startedAt = performance.now();
      logRendererDebug('diffEffect:compute-start', {
        requestId,
        leftPath: diffLeftWorkbook.filePath,
        rightPath: diffRightWorkbook.filePath,
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
        primaryKeyCol: requestedPrimaryKeyCol,
        rowSimilarityThreshold,
      });
      try {
        const result = await window.excelAPI.computeThreeWayDiff(req);
        const durationMs = Math.round(performance.now() - startedAt);
        if (!result) {
          logRendererDebug('diffEffect:compute-null', { requestId, durationMs });
          return;
        }
        if (cancelled) {
          logRendererDebug('diffEffect:compute-cancelled-after-result', { requestId, durationMs });
          return;
        }
        const allDiffSheets =
          result.sheets && result.sheets.length > 0
            ? result.sheets
            : result.sheet
              ? [result.sheet]
              : [];
        const nextIndex = Math.min(selectedDiffSheetIndex, Math.max(0, allDiffSheets.length - 1));
        setDiffSheets(allDiffSheets);
        setSelectedDiffSheetIndex(nextIndex);
        setDiffSelectedCell(null);
        resetDiffHistory();
        logRendererDebug('diffEffect:compute-end', {
          requestId,
          durationMs,
          sheetCount: allDiffSheets.length,
          activeSheet: allDiffSheets[nextIndex]?.sheetName ?? null,
          diffCellCount: allDiffSheets[nextIndex]?.cells?.length ?? 0,
        });
      } catch (error) {
        logRendererDebug('diffEffect:compute-error', {
          requestId,
          message: error instanceof Error ? error.message : String(error),
        });
        console.error(error);
      }
    })();
    return () => {
      cancelled = true;
      logRendererDebug('diffEffect:compute-cancel-request', {
        leftPath: diffLeftWorkbook.filePath,
        rightPath: diffRightWorkbook.filePath,
      });
    };
  }, [
    logRendererDebug,
    nextDebugRequestId,
    requestedPrimaryKeyCol,
    rowSimilarityThreshold,
    diffLeftWorkbook?.filePath,
    diffRightWorkbook?.filePath,
    mode,
    resetDiffHistory,
  ]);


  /**
   * 单文件编辑模式下，当用户修改某个输入框时：
   * - 更新内存中的 rows；
   * - 在 changes Map 中记录此单元格修改，供后续一次性保存。
   */
  const handleCellChange = useCallback(
    (address: string, newValue: string) => {
      setRows((prev) =>
        prev.map((row) =>
          row.map((cell) =>
            cell.address === address
              ? {
                  ...cell,
                  value: newValue === '' ? null : newValue,
                }
              : cell,
          ),
        ),
      );

      setChanges((prev) => {
        const next = new Map(prev);
        next.set(address, {
          address,
          newValue: newValue === '' ? null : newValue,
        });
        return next;
      });
    },
    [],
  );
  const updateWorkbookForDiffRowDelete = useCallback(
    (
      workbook: OpenResult | null,
      targetSheetName: string,
      targetRowNumber: number,
    ): OpenResult | null => {
      if (!workbook) return workbook;
      const nextSheets = workbook.sheets.map((sheet) => {
        if (sheet.sheetName !== targetSheetName) return sheet;
        const remainingRows = sheet.rows.filter((_, rowIndex) => rowIndex !== targetRowNumber - 1);
        const nextRows = remainingRows.map((row, rowIndex) =>
          row.map((cell, colIndex) => ({
            ...cell,
            address: makeAddress(colIndex + 1, rowIndex + 1),
            row: rowIndex + 1,
            col: colIndex + 1,
          })),
        );
        return {
          ...sheet,
          rows: nextRows,
        };
      });
      return {
        ...workbook,
        sheet: nextSheets[0] ?? workbook.sheet,
        sheets: nextSheets,
      };
    },
    [],
  );
  const updateDiffRowOpsAfterDeletedRow = useCallback(
    (
      setter: React.Dispatch<React.SetStateAction<Map<string, SaveMergeRowOp[]>>>,
      targetSheetName: string,
      targetRowNumber: number,
      alignedRowNumber: number,
      options: {
        removeInsertedOp: boolean;
        removeVisualRow: boolean;
        appendDeleteOp?: SaveMergeRowOp | null;
      },
    ) => {
      setter((prev) => {
        const currentOps = prev.get(targetSheetName) ?? [];
        let nextOps = currentOps.filter(
          (op) => !(options.removeInsertedOp && op.action === 'insert' && op.visualRowNumber === alignedRowNumber),
        );
        if (options.appendDeleteOp) {
          nextOps = [...nextOps, options.appendDeleteOp];
        }
        nextOps = nextOps
          .map((op) => {
            let nextOp = op;
            if (op.targetRowNumber > targetRowNumber) {
              nextOp = {
                ...nextOp,
                targetRowNumber: op.targetRowNumber - 1,
              };
            }
            if (options.removeVisualRow && nextOp.visualRowNumber != null && nextOp.visualRowNumber > alignedRowNumber) {
              nextOp = {
                ...nextOp,
                visualRowNumber: nextOp.visualRowNumber - 1,
              };
            }
            return nextOp;
          })
          .sort((a, b) => a.targetRowNumber - b.targetRowNumber || (a.visualRowNumber ?? 0) - (b.visualRowNumber ?? 0));
        const next = new Map(prev);
        if (nextOps.length === 0) next.delete(targetSheetName);
        else next.set(targetSheetName, nextOps);
        return next;
      });
    },
    [],
  );
  const shiftDiffChangesByInsertedRow = useCallback(
    (
      setter: React.Dispatch<React.SetStateAction<Map<string, Map<string, CellChange>>>>,
      targetSheetName: string,
      targetRowNumber: number,
    ) => {
      setter((prev) => {
        const sheetChanges = prev.get(targetSheetName);
        if (!sheetChanges || sheetChanges.size === 0) return prev;
        const shifted = new Map<string, CellChange>();
        sheetChanges.forEach((change) => {
          const parsed = parseCellAddress(change.address);
          if (!parsed || parsed.rowNumber < targetRowNumber) {
            shifted.set(change.address, change);
            return;
          }
          const nextAddress = makeAddress(parsed.colNumber, parsed.rowNumber + 1);
          shifted.set(nextAddress, {
            ...change,
            address: nextAddress,
          });
        });
        const next = new Map(prev);
        next.set(targetSheetName, shifted);
        return next;
      });
    },
    [],
  );
  const shiftDiffChangesByDeletedRow = useCallback(
    (
      setter: React.Dispatch<React.SetStateAction<Map<string, Map<string, CellChange>>>>,
      targetSheetName: string,
      targetRowNumber: number,
    ) => {
      setter((prev) => {
        const sheetChanges = prev.get(targetSheetName);
        if (!sheetChanges || sheetChanges.size === 0) return prev;
        const shifted = new Map<string, CellChange>();
        sheetChanges.forEach((change) => {
          const parsed = parseCellAddress(change.address);
          if (!parsed) {
            shifted.set(change.address, change);
            return;
          }
          if (parsed.rowNumber === targetRowNumber) {
            return;
          }
          if (parsed.rowNumber < targetRowNumber) {
            shifted.set(change.address, change);
            return;
          }
          const nextAddress = makeAddress(parsed.colNumber, parsed.rowNumber - 1);
          shifted.set(nextAddress, {
            ...change,
            address: nextAddress,
          });
        });
        const next = new Map(prev);
        if (shifted.size === 0) next.delete(targetSheetName);
        else next.set(targetSheetName, shifted);
        return next;
      });
    },
    [],
  );
  const shiftDiffRowOpsByInsertedRow = useCallback(
    (
      setter: React.Dispatch<React.SetStateAction<Map<string, SaveMergeRowOp[]>>>,
      targetSheetName: string,
      targetRowNumber: number,
    ) => {
      setter((prev) => {
        const ops = prev.get(targetSheetName);
        if (!ops || ops.length === 0) return prev;
        const next = new Map(prev);
        next.set(
          targetSheetName,
          ops.map((op) =>
            op.targetRowNumber >= targetRowNumber
              ? {
                  ...op,
                  targetRowNumber: op.targetRowNumber + 1,
                }
              : op,
          ),
        );
        return next;
      });
    },
    [],
  );
  const appendDiffRowOp = useCallback(
    (
      setter: React.Dispatch<React.SetStateAction<Map<string, SaveMergeRowOp[]>>>,
      targetSheetName: string,
      rowOp: SaveMergeRowOp,
    ) => {
      setter((prev) => {
        const next = new Map(prev);
        const ops = [...(next.get(targetSheetName) ?? []), rowOp].sort(
          (a, b) => a.targetRowNumber - b.targetRowNumber || (a.visualRowNumber ?? 0) - (b.visualRowNumber ?? 0),
        );
        next.set(targetSheetName, ops);
        return next;
      });
    },
    [],
  );

  /**
   * 将单文件编辑模式下所有修改过的单元格一次性写回原 Excel。
   */
  const handleSave = useCallback(async () => {
    if (!filePath || changes.size === 0) return;
    setSaving(true);
    try {
      const changeList = Array.from(changes.values());
      await window.excelAPI.saveChanges({
        changes: changeList,
        sheetName: sheetName ?? undefined,
        sheetIndex: selectedSheetIndex,
      });
      setChanges(new Map());
      // 不需要刷新格式，只要值正确写回即可
    } catch (e) {
      alert(`保存失败：${(e as any)?.message ?? String(e)}`);
    } finally {
      setSaving(false);
    }
  }, [changes, filePath, sheetName, selectedSheetIndex]);

  const getSheetFromWorkbook = useCallback((workbook: OpenResult | null, targetSheetName: string | null) => {
    if (!workbook || !targetSheetName) return null;
    return workbook.sheets.find((sheet) => sheet.sheetName === targetSheetName) ?? null;
  }, []);
  const currentDiffSheet = useMemo(
    () => diffSheets[selectedDiffSheetIndex] ?? null,
    [diffSheets, selectedDiffSheetIndex],
  );
  const currentMergeSheet = useMemo(
    () => mergeSheets[selectedMergeSheetIndex] ?? null,
    [mergeSheets, selectedMergeSheetIndex],
  );
  const currentDiffSheetName = currentDiffSheet?.sheetName ?? null;
  const currentDiffRowsMeta = useMemo(
    () => currentDiffSheet?.rowsMeta ?? [],
    [currentDiffSheet],
  );
  const currentDiffColumnsMeta = useMemo(
    () => currentDiffSheet?.columnsMeta ?? [],
    [currentDiffSheet],
  );
  const activePrimaryKeySheet = useMemo(
    () => (mode === 'diff' ? currentDiffSheet : mode === 'merge' ? currentMergeSheet : null),
    [mode, currentDiffSheet, currentMergeSheet],
  );
  const activePrimaryKeyAlignedCol = activePrimaryKeySheet?.primaryKeyAlignedCol ?? null;
  const activePrimaryKeyOursCol = activePrimaryKeySheet?.primaryKeyOursCol ?? null;
  const activePrimaryKeySource = activePrimaryKeySheet?.primaryKeySource ?? 'none';
  const currentDiffLeftRows = useMemo(
    () => getSheetFromWorkbook(diffLeftWorkbook, currentDiffSheetName)?.rows ?? [],
    [diffLeftWorkbook, currentDiffSheetName, getSheetFromWorkbook],
  );
  const currentDiffRightRows = useMemo(
    () => getSheetFromWorkbook(diffRightWorkbook, currentDiffSheetName)?.rows ?? [],
    [diffRightWorkbook, currentDiffSheetName, getSheetFromWorkbook],
  );
  const activePrimaryKeyColText = useMemo(() => {
    if (activePrimaryKeyOursCol && activePrimaryKeyOursCol >= 1) {
      return `第 ${activePrimaryKeyOursCol} 列（${colNumberToLabel(activePrimaryKeyOursCol)}）`;
    }
    if (activePrimaryKeyAlignedCol && activePrimaryKeyAlignedCol >= 1) {
      return `对齐后第 ${activePrimaryKeyAlignedCol} 列`;
    }
    return '';
  }, [activePrimaryKeyAlignedCol, activePrimaryKeyOursCol]);
  const autoPrimaryKeySourceText = useMemo(() => {
    switch (activePrimaryKeySource) {
      case 'auto-implicit':
        return '强匹配';
      case 'auto-header':
        return '表头命中';
      case 'auto-weak':
        return '弱匹配';
      default:
        return '';
    }
  }, [activePrimaryKeySource]);
  const autoPrimaryKeyDisplayText = useMemo(() => {
    if (!activePrimaryKeyColText || activePrimaryKeySource === 'none') return '未识别到主键';
    return autoPrimaryKeySourceText ? `${activePrimaryKeyColText}（${autoPrimaryKeySourceText}）` : activePrimaryKeyColText;
  }, [activePrimaryKeyColText, activePrimaryKeySource, autoPrimaryKeySourceText]);
  useEffect(() => {
    if (mode !== 'diff' && mode !== 'merge') {
      setAutoHasPrimaryKey(false);
      setPrimaryKeyHint('');
      return;
    }
    if (primaryKeyMode === 'none') {
      setAutoHasPrimaryKey(false);
      setPrimaryKeyHint('无主键：固定使用序列 / 内容对齐');
      return;
    }
    if (primaryKeyMode === 'manual') {
      const manualCol = Math.max(1, Math.floor(primaryKeyCol || 1));
      const manualColText = `第 ${manualCol} 列（${colNumberToLabel(manualCol)}）`;
      const manualApplied = activePrimaryKeySource === 'manual' && !!activePrimaryKeyColText;
      setAutoHasPrimaryKey(manualApplied);
      setPrimaryKeyHint(
        manualApplied
          ? `手动指定：当前使用 ${activePrimaryKeyColText || manualColText} 作为主键`
          : `手动指定：已请求 ${manualColText}，但当前工作表无法稳定使用该列，已退回无主键对齐`,
      );
      return;
    }
    const autoDetected = activePrimaryKeySource !== 'none' && !!activePrimaryKeyColText;
    setAutoHasPrimaryKey(autoDetected);
    setPrimaryKeyHint(
      autoDetected
        ? `自动识别：当前识别到 ${autoPrimaryKeyDisplayText}`
        : '自动识别：未找到稳定主键，当前回退为无主键对齐',
    );
  }, [mode, primaryKeyMode, primaryKeyCol, activePrimaryKeySource, activePrimaryKeyColText, autoPrimaryKeyDisplayText]);
  const countSheetChanges = useCallback(
    (changesBySheet: Map<string, Map<string, CellChange>>) =>
      Array.from(changesBySheet.values()).reduce((sum, sheetChanges) => sum + sheetChanges.size, 0),
    [],
  );
  const diffLeftChangeCount = useMemo(
    () => countSheetChanges(diffLeftChangesBySheet),
    [countSheetChanges, diffLeftChangesBySheet],
  );
  const diffRightChangeCount = useMemo(
    () => countSheetChanges(diffRightChangesBySheet),
    [countSheetChanges, diffRightChangesBySheet],
  );
  const countSheetRowOps = useCallback(
    (rowOpsBySheet: Map<string, SaveMergeRowOp[]>) =>
      Array.from(rowOpsBySheet.values()).reduce((sum, ops) => sum + ops.length, 0),
    [],
  );
  const diffLeftPendingCount = useMemo(
    () => diffLeftChangeCount + countSheetRowOps(diffLeftRowOpsBySheet),
    [countSheetRowOps, diffLeftChangeCount, diffLeftRowOpsBySheet],
  );
  const diffRightPendingCount = useMemo(
    () => diffRightChangeCount + countSheetRowOps(diffRightRowOpsBySheet),
    [countSheetRowOps, diffRightChangeCount, diffRightRowOpsBySheet],
  );
  const updateWorkbookForDiffChange = useCallback(
    (
      workbook: OpenResult | null,
      targetSheetName: string,
      cell: DiffCellData,
      nextValue: string | number | null,
    ): OpenResult | null => {
      if (!workbook || !cell.sourceRowNumber || !cell.sourceColNumber) return workbook;
      const nextSheets = workbook.sheets.map((sheet) => {
        if (sheet.sheetName !== targetSheetName) return sheet;
        const nextRows = sheet.rows.map((row, rowIndex) => {
          if (rowIndex !== cell.sourceRowNumber! - 1) return row;
          const nextRow = [...row];
          while (nextRow.length < cell.sourceColNumber!) {
            const colNumber = nextRow.length + 1;
            nextRow.push({
              address: makeAddress(colNumber, cell.sourceRowNumber!),
              row: cell.sourceRowNumber!,
              col: colNumber,
              value: null,
            });
          }
          const existingCell = nextRow[cell.sourceColNumber! - 1];
          nextRow[cell.sourceColNumber! - 1] = {
            ...(existingCell ?? {
              address: makeAddress(cell.sourceColNumber!, cell.sourceRowNumber!),
              row: cell.sourceRowNumber!,
              col: cell.sourceColNumber!,
            }),
            value: nextValue,
          };
          return nextRow;
        });
        return {
          ...sheet,
          rows: nextRows,
        };
      });
      return {
        ...workbook,
        sheet: nextSheets[0] ?? workbook.sheet,
        sheets: nextSheets,
      };
    },
    [],
  );
  const updateWorkbookForDiffRowInsert = useCallback(
    (
      workbook: OpenResult | null,
      targetSheetName: string,
      targetRowNumber: number,
      rowValues: (string | number | null)[],
    ): OpenResult | null => {
      if (!workbook) return workbook;
      const nextSheets = workbook.sheets.map((sheet) => {
        if (sheet.sheetName !== targetSheetName) return sheet;
        const maxColCount = Math.max(
          rowValues.length,
          sheet.rows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
        );
        const normalizeRow = (
          row: SheetCell[] | undefined,
          rowNumber: number,
          fallbackValues?: (string | number | null)[],
        ): SheetCell[] =>
          Array.from({ length: maxColCount }, (_, idx) => ({
            address: makeAddress(idx + 1, rowNumber),
            row: rowNumber,
            col: idx + 1,
            value:
              row?.[idx]?.value ??
              (fallbackValues && idx < fallbackValues.length ? fallbackValues[idx] ?? null : null),
          }));
        const before = sheet.rows
          .slice(0, Math.max(0, targetRowNumber - 1))
          .map((row, idx) => normalizeRow(row, idx + 1));
        const inserted = normalizeRow(undefined, targetRowNumber, rowValues);
        const after = sheet.rows
          .slice(Math.max(0, targetRowNumber - 1))
          .map((row, idx) => normalizeRow(row, targetRowNumber + idx + 1));
        return {
          ...sheet,
          rows: [...before, inserted, ...after],
        };
      });
      return {
        ...workbook,
        sheet: nextSheets[0] ?? workbook.sheet,
        sheets: nextSheets,
      };
    },
    [],
  );
  const updateDiffChangesBySheet = useCallback(
    (
      setter: React.Dispatch<React.SetStateAction<Map<string, Map<string, CellChange>>>>,
      targetSheetName: string,
      address: string,
      nextValue: string | number | null,
    ) => {
      setter((prev) => {
        const next = new Map(prev);
        const sheetChanges = new Map(next.get(targetSheetName) ?? new Map<string, CellChange>());
        sheetChanges.set(address, {
          address,
          newValue: nextValue,
        });
        next.set(targetSheetName, sheetChanges);
        return next;
      });
    },
    [],
  );
  const computeDiffInsertTargetRowNumber = useCallback(
    (side: DiffSide, alignedRowNumber: number) => {
      const rowKey = side === 'left' ? 'oursRowNumber' : 'theirsRowNumber';
      const rowsMeta = [...currentDiffRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      const idx = rowsMeta.findIndex((row) => row.visualRowNumber === alignedRowNumber);
      if (idx >= 0) {
        for (let i = idx - 1; i >= 0; i -= 1) {
          const rowNumber = rowsMeta[i][rowKey];
          if (rowNumber) return rowNumber + 1;
        }
        for (let i = idx + 1; i < rowsMeta.length; i += 1) {
          const rowNumber = rowsMeta[i][rowKey];
          if (rowNumber) return rowNumber;
        }
      }
      const fallbackRows = side === 'left' ? currentDiffLeftRows : currentDiffRightRows;
      return fallbackRows.length + 1;
    },
    [currentDiffLeftRows, currentDiffRightRows, currentDiffRowsMeta],
  );
  const applyDiffInsertedRowMeta = useCallback(
    (side: DiffSide, alignedRowNumber: number, targetRowNumber: number, copiedWholeRow: boolean) => {
      const rowKey = side === 'left' ? 'oursRowNumber' : 'theirsRowNumber';
      const statusKey = side === 'left' ? 'oursStatus' : 'theirsStatus';
      const oppositeStatusKey = side === 'left' ? 'theirsStatus' : 'oursStatus';
      setDiffSheets((prev) =>
        prev.map((sheet, sheetIndex) => {
          if (sheetIndex !== selectedDiffSheetIndex) return sheet;
          const nextRowsMeta = (sheet.rowsMeta ?? []).map((row) => {
            const nextRow: MergeRowMeta = { ...row };
            const sideRowNumber = nextRow[rowKey];
            if (sideRowNumber != null && sideRowNumber >= targetRowNumber) {
              (nextRow as any)[rowKey] = sideRowNumber + 1;
            }
            if (row.visualRowNumber === alignedRowNumber) {
              (nextRow as any)[rowKey] = targetRowNumber;
              (nextRow as any)[statusKey] = copiedWholeRow ? 'unchanged' : 'modified';
              if (copiedWholeRow) {
                (nextRow as any)[oppositeStatusKey] = 'unchanged';
              }
            }
            return nextRow;
          });
          return {
            ...sheet,
            rowsMeta: nextRowsMeta,
          };
        }),
      );
    },
    [selectedDiffSheetIndex],
  );
  const applyDiffDeletedRowMeta = useCallback(
    (side: DiffSide, alignedRowNumber: number, targetRowNumber: number, removeVisualRow: boolean) => {
      const rowKey = side === 'left' ? 'oursRowNumber' : 'theirsRowNumber';
      const statusKey = side === 'left' ? 'oursStatus' : 'theirsStatus';
      setDiffSheets((prev) =>
        prev.map((sheet, sheetIndex) => {
          if (sheetIndex !== selectedDiffSheetIndex) return sheet;
          if (!sheet.rowsMeta || sheet.rowsMeta.length === 0) return sheet;
          const nextRowsMeta = sheet.rowsMeta.flatMap((row) => {
            if (row.visualRowNumber === alignedRowNumber) {
              if (removeVisualRow) return [];
              return [
                {
                  ...row,
                  [rowKey]: null,
                  [statusKey]: 'deleted',
                } as MergeRowMeta,
              ];
            }
            const nextRow: MergeRowMeta = { ...row };
            const sideRowNumber = nextRow[rowKey];
            if (sideRowNumber != null && sideRowNumber > targetRowNumber) {
              (nextRow as any)[rowKey] = sideRowNumber - 1;
            }
            if (removeVisualRow && row.visualRowNumber > alignedRowNumber) {
              nextRow.visualRowNumber = row.visualRowNumber - 1;
            }
            return [nextRow];
          });
          return {
            ...sheet,
            rowsMeta: nextRowsMeta,
          };
        }),
      );
    },
    [selectedDiffSheetIndex],
  );
  const stageDiffInsertedRow = useCallback(
    (
      side: DiffSide,
      alignedRowNumber: number,
      rowValues: (string | number | null)[],
      copiedWholeRow: boolean,
      recordHistory = true,
    ): number | null => {
      if (!currentDiffSheetName) return null;
      if (recordHistory) {
        pushDiffHistory();
      }
      const targetRowNumber = computeDiffInsertTargetRowNumber(side, alignedRowNumber);
      const rowOp: SaveMergeRowOp = {
        sheetName: currentDiffSheetName,
        action: 'insert',
        targetRowNumber,
        values: rowValues,
        visualRowNumber: alignedRowNumber,
      };
      if (side === 'left') {
        shiftDiffChangesByInsertedRow(setDiffLeftChangesBySheet, currentDiffSheetName, targetRowNumber);
        shiftDiffRowOpsByInsertedRow(setDiffLeftRowOpsBySheet, currentDiffSheetName, targetRowNumber);
        appendDiffRowOp(setDiffLeftRowOpsBySheet, currentDiffSheetName, rowOp);
        setDiffLeftWorkbook((prev) =>
          updateWorkbookForDiffRowInsert(prev, currentDiffSheetName, targetRowNumber, rowValues),
        );
      } else {
        shiftDiffChangesByInsertedRow(setDiffRightChangesBySheet, currentDiffSheetName, targetRowNumber);
        shiftDiffRowOpsByInsertedRow(setDiffRightRowOpsBySheet, currentDiffSheetName, targetRowNumber);
        appendDiffRowOp(setDiffRightRowOpsBySheet, currentDiffSheetName, rowOp);
        setDiffRightWorkbook((prev) =>
          updateWorkbookForDiffRowInsert(prev, currentDiffSheetName, targetRowNumber, rowValues),
        );
      }
      applyDiffInsertedRowMeta(side, alignedRowNumber, targetRowNumber, copiedWholeRow);
      return targetRowNumber;
    },
    [
      appendDiffRowOp,
      applyDiffInsertedRowMeta,
      computeDiffInsertTargetRowNumber,
      currentDiffSheetName,
      pushDiffHistory,
      shiftDiffChangesByInsertedRow,
      shiftDiffRowOpsByInsertedRow,
      updateWorkbookForDiffRowInsert,
    ],
  );
  const handleDiffCellChange = useCallback(
    (
      side: DiffSide,
      cell: DiffCellData,
      newValue: string,
      options?: {
        recordHistory?: boolean;
      },
    ) => {
      if (!currentDiffSheetName || !cell.sourceRowNumber || !cell.sourceColNumber) return;
      const currentValue = cell.value == null ? '' : String(cell.value);
      if (currentValue === newValue) return;
      if (options?.recordHistory !== false) {
        pushDiffHistory();
      }
      const targetAddress = cell.address ?? makeAddress(cell.sourceColNumber, cell.sourceRowNumber);
      const nextValue = newValue === '' ? null : newValue;
      if (side === 'left') {
        setDiffLeftWorkbook((prev) => updateWorkbookForDiffChange(prev, currentDiffSheetName, cell, nextValue));
        updateDiffChangesBySheet(setDiffLeftChangesBySheet, currentDiffSheetName, targetAddress, nextValue);
        return;
      }
      setDiffRightWorkbook((prev) => updateWorkbookForDiffChange(prev, currentDiffSheetName, cell, nextValue));
      updateDiffChangesBySheet(setDiffRightChangesBySheet, currentDiffSheetName, targetAddress, nextValue);
    },
    [
      currentDiffSheetName,
      pushDiffHistory,
      updateDiffChangesBySheet,
      updateWorkbookForDiffChange,
    ],
  );
  const handleApplyDiffOtherCell = useCallback(
    (side: DiffSide, cell: DiffCellData) => {
      if (sameComparableValue(cell.value, cell.otherValue)) {
        alert('另一边相同位置的单元格和值当前一致，没有需要复制的内容。');
        return;
      }
      if (!cell.sourceRowNumber) {
        if (!cell.sourceColNumber) {
          alert('当前这一侧缺少对应列，无法把另一边的值复制过来。');
          return;
        }
        const targetRows = side === 'left' ? currentDiffLeftRows : currentDiffRightRows;
        const maxColCount = Math.max(
          cell.sourceColNumber,
          targetRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
        );
        const rowValues = Array(maxColCount).fill(null) as (string | number | null)[];
        rowValues[cell.sourceColNumber - 1] = cell.otherValue ?? null;
        const insertedRowNumber = stageDiffInsertedRow(side, cell.alignedRowNumber, rowValues, false);
        if (insertedRowNumber == null) {
          alert('当前工作表不存在，无法插入新行。');
        }
        return;
      }
      handleDiffCellChange(side, cell, cell.otherValue == null ? '' : String(cell.otherValue));
    },
    [currentDiffLeftRows, currentDiffRightRows, handleDiffCellChange, stageDiffInsertedRow],
  );
  const handleApplyDiffOtherRow = useCallback(
    (side: DiffSide, cell: DiffCellData) => {
      const rowMeta =
        currentDiffRowsMeta.find((row) => row.visualRowNumber === cell.alignedRowNumber) ?? {
          oursRowNumber: currentDiffLeftRows[cell.alignedRowNumber - 1] ? cell.alignedRowNumber : null,
          theirsRowNumber: currentDiffRightRows[cell.alignedRowNumber - 1] ? cell.alignedRowNumber : null,
        };
      const targetRowNumber = side === 'left' ? rowMeta.oursRowNumber ?? null : rowMeta.theirsRowNumber ?? null;
      const sourceRowNumber = side === 'left' ? rowMeta.theirsRowNumber ?? null : rowMeta.oursRowNumber ?? null;
      const effectiveColumns =
        currentDiffColumnsMeta.length > 0
          ? currentDiffColumnsMeta
          : Array.from(
              {
                length: Math.max(
                  currentDiffLeftRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
                  currentDiffRightRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
                ),
              },
              (_, idx) => ({
                col: idx + 1,
                oursCol: idx + 1,
                theirsCol: idx + 1,
              }),
            );
      const targetRows = side === 'left' ? currentDiffLeftRows : currentDiffRightRows;
      const sourceRows = side === 'left' ? currentDiffRightRows : currentDiffLeftRows;
      if (!sourceRowNumber) {
        alert('另一边也没有对应行，无法复制整行。');
        return;
      }
      if (!targetRowNumber) {
        const maxTargetColCount = Math.max(
          targetRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
          effectiveColumns.reduce((max, columnMeta) => {
            const targetColNumber = side === 'left' ? columnMeta.oursCol ?? 0 : columnMeta.theirsCol ?? 0;
            return Math.max(max, targetColNumber);
          }, 0),
        );
        const rowValues = Array(maxTargetColCount).fill(null) as (string | number | null)[];
        effectiveColumns.forEach((columnMeta) => {
          const targetColNumber = side === 'left' ? columnMeta.oursCol ?? null : columnMeta.theirsCol ?? null;
          if (!targetColNumber) return;
          const sourceColNumber = side === 'left' ? columnMeta.theirsCol ?? null : columnMeta.oursCol ?? null;
          const sourceValue =
            sourceColNumber ? sourceRows[sourceRowNumber - 1]?.[sourceColNumber - 1]?.value ?? null : null;
          rowValues[targetColNumber - 1] = sourceValue;
        });
        const insertedRowNumber = stageDiffInsertedRow(side, cell.alignedRowNumber, rowValues, true);
        if (insertedRowNumber == null) {
          alert('当前工作表不存在，无法插入新行。');
        }
        return;
      }
      const changedCells: Array<{
        columnMeta: {
          col: number;
        };
        targetColNumber: number;
        targetValue: string | number | null;
        sourceValue: string | number | null;
        targetAddress: string | null;
      }> = [];
      effectiveColumns.forEach((columnMeta) => {
        const targetColNumber = side === 'left' ? columnMeta.oursCol ?? null : columnMeta.theirsCol ?? null;
        if (!targetColNumber) return;
        const targetSheetCell = targetRows[targetRowNumber - 1]?.[targetColNumber - 1];
        const sourceColNumber = side === 'left' ? columnMeta.theirsCol ?? null : columnMeta.oursCol ?? null;
        const sourceValue =
          sourceRowNumber && sourceColNumber
            ? sourceRows[sourceRowNumber - 1]?.[sourceColNumber - 1]?.value ?? null
            : null;
        const targetValue = targetSheetCell?.value ?? null;
        if (sameComparableValue(targetValue, sourceValue)) return;
        changedCells.push({
          columnMeta,
          targetColNumber,
          targetValue,
          sourceValue,
          targetAddress:
            targetSheetCell?.address ?? (targetRowNumber ? makeAddress(targetColNumber, targetRowNumber) : null),
        });
      });
      if (changedCells.length === 0) {
        alert('另一边整行与当前行一致，没有需要复制的内容。');
        return;
      }
      pushDiffHistory();
      changedCells.forEach(({ columnMeta, targetColNumber, targetValue, sourceValue, targetAddress }) => {
        handleDiffCellChange(
          side,
          {
            alignedRowNumber: cell.alignedRowNumber,
            alignedColNumber: columnMeta.col,
            address: targetAddress,
            value: targetValue,
            otherValue: sourceValue,
            sourceRowNumber: targetRowNumber,
            sourceColNumber: targetColNumber,
            isDifferent: !sameComparableValue(targetValue, sourceValue),
          },
          sourceValue == null ? '' : String(sourceValue),
          { recordHistory: false },
        );
      });
    },
    [
      currentDiffColumnsMeta,
      currentDiffLeftRows,
      currentDiffRightRows,
      currentDiffRowsMeta,
      handleDiffCellChange,
      pushDiffHistory,
      stageDiffInsertedRow,
    ],
  );
  const handleDeleteDiffRow = useCallback(
    (side: DiffSide, cell: DiffCellData) => {
      if (!currentDiffSheetName) return;
      const alignedRowNumber = cell.alignedRowNumber;
      const rowMeta =
        currentDiffRowsMeta.find((row) => row.visualRowNumber === alignedRowNumber) ?? {
          visualRowNumber: alignedRowNumber,
          oursRowNumber: currentDiffLeftRows[alignedRowNumber - 1] ? alignedRowNumber : null,
          theirsRowNumber: currentDiffRightRows[alignedRowNumber - 1] ? alignedRowNumber : null,
        };
      const targetRowNumber = side === 'left' ? rowMeta.oursRowNumber ?? null : rowMeta.theirsRowNumber ?? null;
      if (!targetRowNumber) {
        alert(`当前${side === 'left' ? '左侧' : '右侧'}这一行不存在，无法删除。`);
        return;
      }
      const currentRowOps =
        (side === 'left' ? diffLeftRowOpsBySheet : diffRightRowOpsBySheet).get(currentDiffSheetName) ?? [];
      const existingInsertOp =
        currentRowOps.find((op) => op.action === 'insert' && op.visualRowNumber === alignedRowNumber) ?? null;
      const removeVisualRow = side === 'left' ? rowMeta.theirsRowNumber == null : rowMeta.oursRowNumber == null;
      const deleteOp =
        existingInsertOp == null
          ? {
              sheetName: currentDiffSheetName,
              action: 'delete' as const,
              targetRowNumber,
              visualRowNumber: alignedRowNumber,
            }
          : null;
      pushDiffHistory();
      if (side === 'left') {
        shiftDiffChangesByDeletedRow(setDiffLeftChangesBySheet, currentDiffSheetName, targetRowNumber);
        updateDiffRowOpsAfterDeletedRow(
          setDiffLeftRowOpsBySheet,
          currentDiffSheetName,
          targetRowNumber,
          alignedRowNumber,
          {
            removeInsertedOp: existingInsertOp != null,
            removeVisualRow,
            appendDeleteOp: deleteOp,
          },
        );
        setDiffLeftWorkbook((prev) => updateWorkbookForDiffRowDelete(prev, currentDiffSheetName, targetRowNumber));
      } else {
        shiftDiffChangesByDeletedRow(setDiffRightChangesBySheet, currentDiffSheetName, targetRowNumber);
        updateDiffRowOpsAfterDeletedRow(
          setDiffRightRowOpsBySheet,
          currentDiffSheetName,
          targetRowNumber,
          alignedRowNumber,
          {
            removeInsertedOp: existingInsertOp != null,
            removeVisualRow,
            appendDeleteOp: deleteOp,
          },
        );
        setDiffRightWorkbook((prev) => updateWorkbookForDiffRowDelete(prev, currentDiffSheetName, targetRowNumber));
      }
      applyDiffDeletedRowMeta(side, alignedRowNumber, targetRowNumber, removeVisualRow);
      if (removeVisualRow) {
        setDiffSelectedCell((prev) => {
          if (!prev) return prev;
          const selectedRowNumber = prev.rowIndex + 1;
          if (selectedRowNumber === alignedRowNumber) return null;
          if (selectedRowNumber > alignedRowNumber) {
            return {
              ...prev,
              rowIndex: prev.rowIndex - 1,
            };
          }
          return prev;
        });
      }
    },
    [
      applyDiffDeletedRowMeta,
      currentDiffLeftRows,
      currentDiffRightRows,
      currentDiffRowsMeta,
      currentDiffSheetName,
      diffLeftRowOpsBySheet,
      diffRightRowOpsBySheet,
      pushDiffHistory,
      shiftDiffChangesByDeletedRow,
      updateDiffRowOpsAfterDeletedRow,
      updateWorkbookForDiffRowDelete,
    ],
  );
  const handleSaveDiffSide = useCallback(
    async (side: DiffSide) => {
      const workbook = side === 'left' ? diffLeftWorkbook : diffRightWorkbook;
      const otherWorkbook = side === 'left' ? diffRightWorkbook : diffLeftWorkbook;
      const changesBySheet = side === 'left' ? diffLeftChangesBySheet : diffRightChangesBySheet;
      const rowOpsBySheet = side === 'left' ? diffLeftRowOpsBySheet : diffRightRowOpsBySheet;
      if (!workbook || (changesBySheet.size === 0 && rowOpsBySheet.size === 0)) return;
      setDiffSavingSide(side);
      try {
        const targetSheetNames = new Set<string>([
          ...Array.from(changesBySheet.keys()),
          ...Array.from(rowOpsBySheet.keys()),
        ]);
        for (const targetSheetName of targetSheetNames) {
          const sheetChanges = changesBySheet.get(targetSheetName) ?? new Map<string, CellChange>();
          const rowOps = rowOpsBySheet.get(targetSheetName) ?? [];
          if (sheetChanges.size === 0 && rowOps.length === 0) continue;
          await window.excelAPI.saveChanges({
            filePath: workbook.filePath,
            sheetName: targetSheetName,
            changes: Array.from(sheetChanges.values()),
            rowOps,
          });
        }
        const reloaded = await window.excelAPI.loadWorkbook(workbook.filePath);
        const nextWorkbook = reloaded ?? workbook;
        if (side === 'left') {
          setDiffLeftChangesBySheet(new Map());
          setDiffLeftRowOpsBySheet(new Map());
          if (otherWorkbook) {
            await loadDiffComparison(nextWorkbook, otherWorkbook, currentDiffSheetName ?? undefined);
          } else {
            setDiffLeftWorkbook(nextWorkbook);
          }
        } else {
          setDiffRightChangesBySheet(new Map());
          setDiffRightRowOpsBySheet(new Map());
          if (otherWorkbook) {
            await loadDiffComparison(otherWorkbook, nextWorkbook, currentDiffSheetName ?? undefined);
          } else {
            setDiffRightWorkbook(nextWorkbook);
          }
        }
      } catch (e) {
        alert(`保存 ${side === 'left' ? '左侧' : '右侧'} Excel 失败：${(e as any)?.message ?? String(e)}`);
      } finally {
        setDiffSavingSide(null);
      }
    },
    [
      currentDiffSheetName,
      diffLeftChangesBySheet,
      diffLeftRowOpsBySheet,
      diffLeftWorkbook,
      diffRightChangesBySheet,
      diffRightRowOpsBySheet,
      diffRightWorkbook,
      loadDiffComparison,
    ],
  );
  const hasData = useMemo(() => rows.length > 0, [rows]);
  const hasMergeData = useMemo(() => mergeCells.length > 0, [mergeCells]);
  const diffRemainingCount = useMemo(() => {
    const rowMetas =
      currentDiffRowsMeta.length > 0
        ? currentDiffRowsMeta
        : Array.from({ length: Math.max(currentDiffLeftRows.length, currentDiffRightRows.length) }, (_, idx) => ({
            visualRowNumber: idx + 1,
            oursRowNumber: currentDiffLeftRows[idx] ? idx + 1 : null,
            theirsRowNumber: currentDiffRightRows[idx] ? idx + 1 : null,
          }));
    const columnMetas =
      currentDiffColumnsMeta.length > 0
        ? currentDiffColumnsMeta
        : Array.from(
            {
              length: Math.max(
                currentDiffLeftRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
                currentDiffRightRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0),
              ),
            },
            (_, idx) => ({
              col: idx + 1,
              oursCol: idx + 1,
              theirsCol: idx + 1,
            }),
          );

    let remaining = 0;
    rowMetas.forEach((rowMeta) => {
      columnMetas.forEach((columnMeta) => {
        const leftValue =
          rowMeta.oursRowNumber && columnMeta.oursCol
            ? currentDiffLeftRows[rowMeta.oursRowNumber - 1]?.[columnMeta.oursCol - 1]?.value ?? null
            : null;
        const rightValue =
          rowMeta.theirsRowNumber && columnMeta.theirsCol
            ? currentDiffRightRows[rowMeta.theirsRowNumber - 1]?.[columnMeta.theirsCol - 1]?.value ?? null
            : null;
        if (!sameComparableValue(leftValue, rightValue)) {
          remaining += 1;
        }
      });
    });
    return remaining;
  }, [
    currentDiffColumnsMeta,
    currentDiffLeftRows,
    currentDiffRightRows,
    currentDiffRowsMeta,
  ]);
  const mergeRemainingCount = useMemo(() => {
    const resolvedSet = resolvedBySheet.get(selectedMergeSheetIndex) ?? new Set<string>();
    return mergeCells.reduce((count, cell) => {
      if (cell.status !== 'conflict') return count;
      return resolvedSet.has(`${cell.row}:${cell.col}`) ? count : count + 1;
    }, 0);
  }, [mergeCells, resolvedBySheet, selectedMergeSheetIndex]);
  const mergeCellKeySet = useMemo(
    () => new Set(mergeCells.map((c) => `${c.row}:${c.col}`)),
    [mergeCells],
  );
  useEffect(() => {
    if (mode !== 'merge' || !mergeInfo || !showFullTables) {
      setFullOursRows([]);
      setFullTheirsRows([]);
      return;
    }
    let cancelled = false;
    (async () => {
      const [oursSheet, theirsSheet] = await Promise.all([
        window.excelAPI.getSheetData({
          path: mergeInfo.oursPath,
          sheetName: mergeInfo.sheetName,
          sheetIndex: selectedMergeSheetIndex,
        }),
        window.excelAPI.getSheetData({
          path: mergeInfo.theirsPath,
          sheetName: mergeInfo.sheetName,
          sheetIndex: selectedMergeSheetIndex,
        }),
      ]);
      if (cancelled) return;
      const oursRows = (oursSheet?.rows ?? []).map((row: SheetCell[]) =>
        row.map((c: SheetCell) => c.value ?? null),
      );
      const theirsRows = (theirsSheet?.rows ?? []).map((row: SheetCell[]) =>
        row.map((c: SheetCell) => c.value ?? null),
      );
      setFullOursRows(oursRows);
      setFullTheirsRows(theirsRows);
    })();
    return () => {
      cancelled = true;
      logRendererDebug('mergedPreview:getThreeWayRows-cancel-request', {
        sheetName: mergeInfo.sheetName,
      });
    };
  }, [
    mode,
    showFullTables,
    mergeInfo?.oursPath,
    mergeInfo?.theirsPath,
    mergeInfo?.sheetName,
    selectedMergeSheetIndex,
  ]);

  const mergeCellsByRow = useMemo(() => {
    const m = new Map<number, MergeCell[]>();
    mergeCells.forEach((cell) => {
      if (!m.has(cell.row)) m.set(cell.row, []);
      m.get(cell.row)!.push(cell);
    });
    return m;
  }, [mergeCells]);

  // 顶部"公式栏"当前要展示的单元格信息
  const selectedMergeCellData = useMemo(() => {
    if (mode !== 'merge' || !selectedMergeCell) return null;
    const rowNumber = selectedMergeCell.rowIndex + 1;
    const colNumber = selectedMergeCell.colIndex + 1;
    const rowCells = mergeCellsByRow.get(rowNumber);
    const hit = rowCells?.find((c) => c.col === colNumber) ?? null;
    if (hit) return hit;
    const keyCol =
      typeof displayPrimaryKeyCol === 'number' && displayPrimaryKeyCol >= 1
        ? Math.floor(displayPrimaryKeyCol)
        : -1;
    if (keyCol > 0 && colNumber === keyCol) {
      const meta = mergeRowsMeta.find((m) => m.visualRowNumber === rowNumber);
      if (!meta) return null;
      const value = meta.key ?? null;
      const addressRow = meta.oursRowNumber ?? meta.baseRowNumber ?? meta.theirsRowNumber ?? rowNumber;
      return {
        address: makeAddress(colNumber, addressRow),
        row: rowNumber,
        col: colNumber,
        baseValue: value,
        oursValue: value,
        theirsValue: value,
        status: 'unchanged',
        mergedValue: value,
      };
    }
    return null;
  }, [mode, selectedMergeCell, mergeCellsByRow, displayPrimaryKeyCol, mergeRowsMeta]);


  const mergedPath = useMemo(() => {
    if (!mergeInfo) return null;
    if (cliInfo?.mode === 'merge') {
      return cliInfo.mergedPath ?? mergeInfo.oursPath;
    }
    if (cliInfo?.mode === 'diff') {
      return mergeInfo.oursPath;
    }
    return null;
  }, [mergeInfo, cliInfo]);
  const currentRowOps = useMemo(
    () => mergeRowOpsBySheet.get(selectedMergeSheetIndex) ?? EMPTY_MERGE_ROW_OPS,
    [mergeRowOpsBySheet, selectedMergeSheetIndex],
  );
  const currentColOps = useMemo(
    () => mergeColOpsBySheet.get(selectedMergeSheetIndex) ?? EMPTY_MERGE_COL_OPS,
    [mergeColOpsBySheet, selectedMergeSheetIndex],
  );
  useEffect(() => {
    if (mode !== 'merge' || !mergeInfo) {
      setMergedPreviewRows([]);
      setMergedPreviewRowVisuals([]);
      setMergedPreviewAlignedCols([]);
      setMergeThreeWayRows([]);
      return;
    }
    let cancelled = false;
    (async () => {
      const requestId = nextDebugRequestId('merged-preview');
      const metas = [...mergeRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      const minRows = Math.max(1, Math.floor(mergedPreviewMinRows));
      if (metas.length === 0) {
        if (!cancelled) {
          setMergedPreviewRows(Array.from({ length: minRows }, () => []));
          setMergedPreviewRowVisuals(Array.from({ length: minRows }, () => null));
          setMergedPreviewAlignedCols([]);
          setMergeThreeWayRows([]);
        }
        return;
      }
      const startedAt = performance.now();
      logRendererDebug('mergedPreview:getThreeWayRows-start', {
        requestId,
        sheetName: mergeInfo.sheetName,
        rowCount: metas.length,
        compareHeaderRowCount: COMPARE_HEADER_ROW_COUNT,
      });
      const rowsReq = metas.map((m) => ({
        rowNumber: m.visualRowNumber,
        baseRowNumber: m.baseRowNumber,
        oursRowNumber: m.oursRowNumber,
        theirsRowNumber: m.theirsRowNumber,
      }));
      const result = await window.excelAPI.getThreeWayRows({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        compareMode,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        debugRequestId: requestId,
        rows: rowsReq,
      });
      const durationMs = Math.round(performance.now() - startedAt);
      if (!result) {
        logRendererDebug('mergedPreview:getThreeWayRows-null', { requestId, durationMs });
        return;
      }
      if (cancelled) {
        logRendererDebug('mergedPreview:getThreeWayRows-cancelled-after-result', { requestId, durationMs });
        return;
      }
      setMergeThreeWayRows(result.rows ?? []);
      const rawColCount = result.colCount ?? 0;
      // Build effective column list considering col ops
      const deletedAlignedCols = new Set<number>();
      const insertedAlignedCols: number[] = [];
      currentColOps.forEach((op, alignedCol) => {
        if (op.action === 'delete') deletedAlignedCols.add(alignedCol);
        else if (op.action === 'insert') insertedAlignedCols.push(alignedCol);
      });
      // Map aligned col -> ours col for non-deleted columns
      // IMPORTANT: 只包含 ours 模板中存在的列（oursCol 非空），
      // theirs-only 列只有在用户显式选择"插入"后才通过下方 insert 逻辑加入，
      // 否则 merged 预览中会出现重复列。
      const effectiveColMap: { alignedCol: number; oursCol: number | null }[] = [];
      for (let c = 1; c <= rawColCount; c += 1) {
        if (deletedAlignedCols.has(c)) continue;
        const meta = mergeColumnsMeta.find((m) => m.col === c);
        if (!meta?.oursCol) continue;
        effectiveColMap.push({ alignedCol: c, oursCol: meta.oursCol });
      }
      // Add inserted columns (theirs-only)
      insertedAlignedCols.sort((a, b) => a - b);
      // IMPORTANT: 先收集所有插入位置，然后从后往前插入，避免索引错乱
      const insertions: Array<{ idx: number; col: number }> = [];
      for (const ac of insertedAlignedCols) {
        const meta = mergeColumnsMeta.find((m) => m.col === ac);
        if (meta && !meta.oursCol && meta.theirsCol) {
          // Find insertion position
          let insertIdx = effectiveColMap.length;
          for (let i = 0; i < effectiveColMap.length; i += 1) {
            if (effectiveColMap[i].alignedCol > ac) {
              insertIdx = i;
              break;
            }
          }
          insertions.push({ idx: insertIdx, col: ac });
        }
      }
      // 从后往前插入，避免每次 splice 改变后续索引
      // 同一位置的多个插入按 col 降序处理，确保最终顺序正确
      insertions.sort((a, b) => b.idx - a.idx || b.col - a.col);
      for (const ins of insertions) {
        effectiveColMap.splice(ins.idx, 0, { alignedCol: ins.col, oursCol: null });
      }
      const colCount = effectiveColMap.length;
      const mergedRows: (string | number | null)[][] = [];
      const mergedVisuals: (number | null)[] = [];
      result.rows.forEach((rowRes: any, idx: number) => {
        const meta = metas[idx];
        const visualRowNumber = meta?.visualRowNumber ?? rowRes.rowNumber ?? idx + 1;
        const op = currentRowOps.get(visualRowNumber);
        const oursMissing = !meta?.oursRowNumber;
        if (oursMissing && op?.action !== 'insert') return;
        if (!oursMissing && op?.action === 'delete') return;
        // Build merged row based on effective columns
        const mergedRow: (string | number | null)[] = [];
        for (let i = 0; i < effectiveColMap.length; i += 1) {
          const colInfo = effectiveColMap[i];
          const alignedCol = colInfo.alignedCol;
          const colMeta = mergeColumnsMeta.find((m) => m.col === alignedCol);
          // Check if there's a diff cell override
          const diffCell = (mergeCellsByRow.get(visualRowNumber) ?? []).find((c) => c.col === alignedCol);
          if (diffCell) {
            mergedRow.push(diffCell.mergedValue ?? null);
            continue;
          }
          // Otherwise get from ours/theirs raw data
          if (op?.action === 'insert' && op.values) {
            // IMPORTANT: 用 alignedCol 索引而非循环索引 i——effectiveColMap 会跳过已删除的列和未插入的 theirs-only 列，
            // 但 op.values 始终按原始 aligned 列顺序排列，所以必须用 alignedCol - 1 取值。
            mergedRow.push(op.values[alignedCol - 1] ?? null);
          } else if (colMeta?.oursCol && rowRes.ours) {
            mergedRow.push(rowRes.ours[alignedCol - 1] ?? null);
          } else if (colMeta?.theirsCol && rowRes.theirs) {
            // For theirs-only columns that are being inserted
            mergedRow.push(rowRes.theirs[alignedCol - 1] ?? null);
          } else {
            mergedRow.push(null);
          }
        }
        mergedRows.push(mergedRow);
        mergedVisuals.push(visualRowNumber);
      });
      while (mergedRows.length < minRows) {
        mergedRows.push(Array(colCount).fill(null));
        mergedVisuals.push(null);
      }
      if (cancelled) return;
      setMergedPreviewAlignedCols(effectiveColMap.map((item) => item.alignedCol));
      setMergedPreviewRows(mergedRows);
      setMergedPreviewRowVisuals(mergedVisuals);
      logRendererDebug('mergedPreview:getThreeWayRows-end', {
        requestId,
        durationMs,
        rowCount: mergedRows.length,
        colCount,
      });
    })();
    return () => {
      cancelled = true;
      logRendererDebug('mergedPreview:getThreeWayRows-cancel-request', {
        sheetName: mergeInfo.sheetName,
      });
    };
  }, [
    logRendererDebug,
    mode,
    mergeInfo,
    mergeRowsMeta,
    mergeCellsByRow,
    mergedPreviewMinRows,
    selectedMergeSheetIndex,
    currentRowOps,
    currentColOps,
    mergeColumnsMeta,
    compareMode,
    nextDebugRequestId,
  ]);

  const selectedDiffCellData = useMemo(() => {
    if (mode !== 'diff' || !diffSelectedCell) return null;
    const visualRowNumber = diffSelectedCell.rowIndex + 1;
    const alignedColNumber = diffSelectedCell.colIndex + 1;
    const rowMeta = currentDiffRowsMeta.find((row) => row.visualRowNumber === visualRowNumber);
    const columnMeta = currentDiffColumnsMeta.find((column) => column.col === alignedColNumber);
    const leftRowNumber = rowMeta?.oursRowNumber ?? null;
    const rightRowNumber = rowMeta?.theirsRowNumber ?? null;
    const leftColNumber = columnMeta?.oursCol ?? null;
    const rightColNumber = columnMeta?.theirsCol ?? null;
    const leftCell =
      leftRowNumber && leftColNumber ? currentDiffLeftRows[leftRowNumber - 1]?.[leftColNumber - 1] ?? null : null;
    const rightCell =
      rightRowNumber && rightColNumber ? currentDiffRightRows[rightRowNumber - 1]?.[rightColNumber - 1] ?? null : null;
    return {
      address:
        leftCell?.address ??
        rightCell?.address ??
        makeAddress(alignedColNumber, leftRowNumber ?? rightRowNumber ?? visualRowNumber),
      leftValue: leftCell?.value ?? null,
      rightValue: rightCell?.value ?? null,
    };
  }, [
    mode,
    diffSelectedCell,
    currentDiffColumnsMeta,
    currentDiffLeftRows,
    currentDiffRightRows,
    currentDiffRowsMeta,
  ]);

  // 顶部“公式栏”当前要展示的单元格坐标和值（diff / merge 共用）
  let currentCellAddress = '';
  let currentCellValue = '';

  if (mode === 'diff' && selectedDiffCellData) {
    currentCellAddress = selectedDiffCellData.address;
    currentCellValue = '';
  } else if (mode === 'merge' && selectedMergeCellData) {
    currentCellAddress = selectedMergeCellData.address;
    // merge 模式下不再用一个“当前值”展示；此字段保留给 diff 模式信息栏占位
    currentCellValue = '';
  }

  const handleSelectMergeCell = useCallback((rowIndex: number, colIndex: number) => {
    setSelectedMergeCell({ rowIndex, colIndex });
  }, []);
  const updateRowOpForSheet = useCallback(
    (sheetIndex: number, visualRowNumber: number, op: SaveMergeRowOp | null) => {
      setMergeRowOpsBySheet((prev) => {
        const next = new Map(prev);
        const sheetOps = new Map(next.get(sheetIndex) ?? new Map<number, SaveMergeRowOp>());
        if (op) sheetOps.set(visualRowNumber, op);
        else sheetOps.delete(visualRowNumber);
        if (sheetOps.size === 0) next.delete(sheetIndex);
        else next.set(sheetIndex, sheetOps);
        return next;
      });
    },
    [],
  );
  const updateColOpForSheet = useCallback(
    (sheetIndex: number, alignedColNumber: number, op: SaveMergeColOp | null) => {
      setMergeColOpsBySheet((prev) => {
        const next = new Map(prev);
        const sheetOps = new Map(next.get(sheetIndex) ?? new Map<number, SaveMergeColOp>());
        if (op) sheetOps.set(alignedColNumber, op);
        else sheetOps.delete(alignedColNumber);
        if (sheetOps.size === 0) next.delete(sheetIndex);
        else next.set(sheetIndex, sheetOps);
        return next;
      });
    },
    [],
  );
  const computeInsertTargetColNumber = useCallback(
    (alignedColNumber: number) => {
      if (!mergeColumnsMeta || mergeColumnsMeta.length === 0) return 1;
      const metaMap = new Map<number, MergeColumnMeta>();
      mergeColumnsMeta.forEach((m) => metaMap.set(m.col, m));
      for (let c = alignedColNumber - 1; c >= 1; c -= 1) {
        const meta = metaMap.get(c);
        if (meta?.oursCol) return meta.oursCol + 1;
      }
      for (let c = alignedColNumber + 1; c <= metaMap.size; c += 1) {
        const meta = metaMap.get(c);
        if (meta?.oursCol) return meta.oursCol;
      }
      return 1;
    },
    [mergeColumnsMeta],
  );
  const computeInsertTargetRowNumber = useCallback(
    (visualRowNumber: number) => {
      const list = [...mergeRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      const idx = list.findIndex((m) => m.visualRowNumber === visualRowNumber);
      if (idx < 0) return 1;
      for (let i = idx - 1; i >= 0; i -= 1) {
        const r = list[i].oursRowNumber;
        if (r) return r + 1;
      }
      for (let i = idx + 1; i < list.length; i += 1) {
        const r = list[i].oursRowNumber;
        if (r) return r;
      }
      return 1;
    },
    [mergeRowsMeta],
  );
  const buildMergedRowValues = useCallback(
    async (visualRowNumber: number, rowMeta: MergeRowMeta) => {
      if (!mergeInfo) return null;
      const result = await window.excelAPI.getThreeWayRow({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        compareMode,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        rowNumber: visualRowNumber,
        baseRowNumber: rowMeta.baseRowNumber ?? null,
        oursRowNumber: rowMeta.oursRowNumber ?? null,
        theirsRowNumber: rowMeta.theirsRowNumber ?? null,
      });
      if (!result) return null;
      const colCount = result.colCount;
      let baseRow: (string | number | null)[] = [];
      if (rowMeta.oursRowNumber) baseRow = result.ours;
      else if (rowMeta.baseRowNumber) baseRow = result.base;
      else if (rowMeta.theirsRowNumber) baseRow = result.theirs;
      else baseRow = Array(colCount).fill(null);
      const mergedRow = baseRow.slice(0, colCount);
      if (mergedRow.length < colCount) {
        mergedRow.push(...Array(colCount - mergedRow.length).fill(null));
      }
      const diffCells = mergeCellsByRow.get(visualRowNumber) ?? [];
      diffCells.forEach((cell) => {
        if (cell.col >= 1 && cell.col <= colCount) {
          mergedRow[cell.col - 1] = cell.mergedValue ?? null;
        }
      });
      return mergedRow;
    },
    [mergeInfo, selectedMergeSheetIndex, mergeCellsByRow, compareMode],
  );

  /**
   * merge 模式下，在右侧详情中点击“用 base / ours / theirs”按钮时：
   * - 更新 mergeSheets 中对应单元格的 mergedValue；
   * - 同步更新当前正在展示的 mergeRows；
   *   这样列表与详情都能立即反映最新选择。
   */
  const markResolvedKeys = useCallback(
    (sheetIndex: number, keys: string[]) => {
      if (keys.length === 0) return;
      setResolvedBySheet((prev) => {
        const next = new Map(prev);
        const current = next.get(sheetIndex) ?? new Set<string>();
        const merged = new Set(current);
        keys.forEach((k) => merged.add(k));
        next.set(sheetIndex, merged);
        return next;
      });
    },
    [],
  );
  const handleResolveMergeCell = useCallback(
    (rowNumber: number, colNumber: number) => {
      const key = `${rowNumber}:${colNumber}`;
      if (!mergeCellKeySet.has(key)) return;
      pushMergeUndoSnapshot();
      markResolvedKeys(selectedMergeSheetIndex, [key]);
    },
    [mergeCellKeySet, markResolvedKeys, pushMergeUndoSnapshot, selectedMergeSheetIndex],
  );

  const handleApplyMergeChoice = useCallback(
    (source: 'base' | 'ours' | 'theirs') => {
      if (!selectedMergeCell) return;

      const { rowIndex, colIndex } = selectedMergeCell;
      const key = `${rowIndex + 1}:${colIndex + 1}`;
      if (!mergeCellKeySet.has(key)) return;
      pushMergeUndoSnapshot();
      // 只标记用户显式操作过的单元格
      markResolvedKeys(selectedMergeSheetIndex, [key]);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.row - 1 !== rowIndex || cell.col - 1 !== colIndex) return cell;
            let value: string | number | null;
            if (source === 'base') value = cell.baseValue;
            else if (source === 'ours') value = cell.oursValue;
            else value = cell.theirsValue;
            return { ...cell, mergedValue: value };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      // 同步当前视图的 cells
      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.row - 1 !== rowIndex || cell.col - 1 !== colIndex) return cell;
          let value: string | number | null;
          if (source === 'base') value = cell.baseValue;
          else if (source === 'ours') value = cell.oursValue;
          else value = cell.theirsValue;
          return { ...cell, mergedValue: value };
        }),
      );
    },
    [selectedMergeCell, selectedMergeSheetIndex, markResolvedKeys, mergeCellKeySet, pushMergeUndoSnapshot],
  );

  const handleApplyMergeRowChoice = useCallback(
    async (rowNumber: number, source: 'ours' | 'theirs') => {
      pushMergeUndoSnapshot();
      const valueFrom = (cell: MergeCell) => (source === 'ours' ? cell.oursValue : cell.theirsValue);

      // 标记这一行所有差异单元格为 resolved
      const keys = mergeCells
        .filter((c) => c.row === rowNumber)
        .map((c) => `${c.row}:${c.col}`);
      markResolvedKeys(selectedMergeSheetIndex, keys);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.row !== rowNumber) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      // 同步当前视图的 cells
      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.row !== rowNumber) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );
      const rowMeta = mergeRowsMeta.find((m) => m.visualRowNumber === rowNumber);
      if (!rowMeta || !mergeInfo) return;
      const oursRowNumber = rowMeta.oursRowNumber ?? null;
      const theirsRowNumber = rowMeta.theirsRowNumber ?? null;
      let op: SaveMergeRowOp | null = null;
      if (!oursRowNumber && theirsRowNumber) {
        if (source === 'theirs') {
          const values = await buildMergedRowValues(rowNumber, rowMeta);
          if (values) {
            op = {
              sheetName: mergeInfo.sheetName,
              action: 'insert',
              targetRowNumber: computeInsertTargetRowNumber(rowNumber),
              values,
              visualRowNumber: rowNumber,
            };
          }
        }
      } else if (oursRowNumber && !theirsRowNumber) {
        if (source === 'theirs') {
          op = {
            sheetName: mergeInfo.sheetName,
            action: 'delete',
            targetRowNumber: oursRowNumber,
            visualRowNumber: rowNumber,
          };
        }
      }
      if (op || currentRowOps.has(rowNumber)) {
        updateRowOpForSheet(selectedMergeSheetIndex, rowNumber, op);
      }
    },
    [
      selectedMergeSheetIndex,
      mergeCells,
      markResolvedKeys,
      mergeRowsMeta,
      mergeInfo,
      buildMergedRowValues,
      computeInsertTargetRowNumber,
      updateRowOpForSheet,
      currentRowOps,
      pushMergeUndoSnapshot,
    ],
  );
  const handleDeleteMergeRow = useCallback(
    (rowNumber: number) => {
      pushMergeUndoSnapshot();
      const keys = mergeCells
        .filter((c) => c.row === rowNumber)
        .map((c) => `${c.row}:${c.col}`);
      markResolvedKeys(selectedMergeSheetIndex, keys);
      const rowMeta = mergeRowsMeta.find((m) => m.visualRowNumber === rowNumber);
      const existingOp = currentRowOps.get(rowNumber) ?? null;
      if (!rowMeta || !mergeInfo) {
        if (existingOp) {
          updateRowOpForSheet(selectedMergeSheetIndex, rowNumber, null);
        }
        return;
      }
      if (rowMeta.oursRowNumber) {
        updateRowOpForSheet(selectedMergeSheetIndex, rowNumber, {
          sheetName: mergeInfo.sheetName,
          action: 'delete',
          targetRowNumber: rowMeta.oursRowNumber,
          visualRowNumber: rowNumber,
        });
        return;
      }
      if (existingOp) {
        updateRowOpForSheet(selectedMergeSheetIndex, rowNumber, null);
      }
    },
    [
      currentRowOps,
      markResolvedKeys,
      mergeCells,
      mergeInfo,
      mergeRowsMeta,
      pushMergeUndoSnapshot,
      selectedMergeSheetIndex,
      updateRowOpForSheet,
    ],
  );

  const handleApplyMergeCellChoice = useCallback(
    (rowNumber: number, colNumber: number, source: 'ours' | 'theirs') => {
      const valueFrom = (cell: MergeCell) => (source === 'ours' ? cell.oursValue : cell.theirsValue);
      const key = `${rowNumber}:${colNumber}`;
      if (!mergeCellKeySet.has(key)) return;

      pushMergeUndoSnapshot();
      markResolvedKeys(selectedMergeSheetIndex, [`${rowNumber}:${colNumber}`]);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.row !== rowNumber || cell.col !== colNumber) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.row !== rowNumber || cell.col !== colNumber) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );
    },
    [selectedMergeSheetIndex, markResolvedKeys, mergeCellKeySet, pushMergeUndoSnapshot],
  );

  const buildMergedColumnValues = useCallback(
    async (colNumber: number) => {
      if (!mergeInfo) return null;
      // Get all rows for this sheet to build column values
      const metas = [...mergeRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      if (metas.length === 0) return [];
      
      const rowsReq = metas.map((m) => ({
        rowNumber: m.visualRowNumber,
        baseRowNumber: m.baseRowNumber,
        oursRowNumber: m.oursRowNumber,
        theirsRowNumber: m.theirsRowNumber,
      }));
      const result = await window.excelAPI.getThreeWayRows({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        compareMode,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        frozenRowCount: COMPARE_HEADER_ROW_COUNT,
        rows: rowsReq,
      });
      if (!result || !result.rows) return [];
      
      // Extract column values from result
      // IMPORTANT: 不要过滤任何行，必须收集所有行的值
      // 因为保存时列操作在行操作之前，那些行还没被删除
      const columnValues: (string | number | null)[] = [];
      result.rows.forEach((rowRes: any) => {
        const visualRowNumber = rowRes.rowNumber ?? 0;
        
        // Get value from aligned column (colNumber is 1-based aligned col)
        const diffCell = (mergeCellsByRow.get(visualRowNumber) ?? []).find((c) => c.col === colNumber);
        if (diffCell) {
          columnValues.push(diffCell.mergedValue ?? null);
        } else if (rowRes.theirs && colNumber >= 1 && colNumber <= rowRes.theirs.length) {
          columnValues.push(rowRes.theirs[colNumber - 1] ?? null);
        } else {
          columnValues.push(null);
        }
      });
      return columnValues;
    },
    [mergeInfo, selectedMergeSheetIndex, mergeRowsMeta, mergeCellsByRow, compareMode],
  );

  const handleApplyMergeColumnChoice = useCallback(
    async (colNumber: number, source: 'ours' | 'theirs') => {
      pushMergeUndoSnapshot();
      const valueFrom = (cell: MergeCell) => (source === 'theirs' ? cell.theirsValue : cell.oursValue);
      const keys = mergeCells.filter((c) => c.col === colNumber).map((c) => `${c.row}:${c.col}`);
      markResolvedKeys(selectedMergeSheetIndex, keys);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.col !== colNumber) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.col !== colNumber) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );

      if (!mergeInfo) return;
      const meta = mergeColumnsMeta.find((c) => c.col === colNumber);
      // theirs-only column -> insert
      const canInsert = source === 'theirs' && meta && !meta.oursCol && meta.theirsCol;
      // ours-only column but user chooses theirs (which is empty) -> delete
      const canDelete = source === 'theirs' && meta && meta.oursCol && !meta.theirsCol;
      if (canInsert) {
        const targetColNumber = computeInsertTargetColNumber(colNumber);
        const values = await buildMergedColumnValues(colNumber);
        if (values) {
          const op: SaveMergeColOp = {
            sheetName: mergeInfo.sheetName,
            action: 'insert',
            targetColNumber,
            alignedColNumber: colNumber,
            source,
            values,
          };
          updateColOpForSheet(selectedMergeSheetIndex, colNumber, op);
        }
      } else if (canDelete && meta.oursCol) {
        const op: SaveMergeColOp = {
          sheetName: mergeInfo.sheetName,
          action: 'delete',
          targetColNumber: meta.oursCol,
          alignedColNumber: colNumber,
          source,
        };
        updateColOpForSheet(selectedMergeSheetIndex, colNumber, op);
      } else if (currentColOps.has(colNumber)) {
        // Clear any existing op if user changes mind
        updateColOpForSheet(selectedMergeSheetIndex, colNumber, null);
      }
    },
    [
      mergeCells,
      markResolvedKeys,
      selectedMergeSheetIndex,
      mergeInfo,
      mergeColumnsMeta,
      computeInsertTargetColNumber,
      buildMergedColumnValues,
      updateColOpForSheet,
      currentColOps,
      pushMergeUndoSnapshot,
    ],
  );

  const handleApplyMergeCellsChoice = useCallback(
    (keys: { rowNumber: number; colNumber: number }[], source: 'base' | 'ours' | 'theirs') => {
      if (!keys.length) return;
      const valueFrom = (cell: MergeCell) =>
        source === 'base' ? cell.baseValue : source === 'ours' ? cell.oursValue : cell.theirsValue;
      const filtered = keys.filter((k) => mergeCellKeySet.has(`${k.rowNumber}:${k.colNumber}`));
      if (!filtered.length) return;
      pushMergeUndoSnapshot();
      const keySet = new Set(filtered.map((k) => `${k.rowNumber}:${k.colNumber}`));
      markResolvedKeys(selectedMergeSheetIndex, Array.from(keySet));

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            const k = `${cell.row}:${cell.col}`;
            if (!keySet.has(k)) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      setMergeCells((prev) =>
        prev.map((cell) => {
          const k = `${cell.row}:${cell.col}`;
          if (!keySet.has(k)) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );
    },
    [selectedMergeSheetIndex, markResolvedKeys, mergeCellKeySet, pushMergeUndoSnapshot],
  );

  /**
   * merge 模式下，将所有工作表的 mergedValue 写回一个目标 Excel 文件。
   *
   * 为了避免误操作，这里会先统计所有发生变化的单元格，
   * 构造一个预览字符串通过 window.confirm 让用户二次确认。
   */
  const handleSaveMergeToFile = useCallback(async () => {
    if (!mergeInfo || mergeSheets.length === 0) return;

    // 生成本次合并的概要信息：mergeSheets.cells 本身就是差异单元格列表
    const changedCells: { sheetName: string; address: string; ours: any; theirs: any; merged: any }[] = [];
    let skippedCells = 0;
    mergeSheets.forEach((sheet) => {
      const rowMetaMap = new Map<number, MergeRowMeta>();
      (sheet.rowsMeta ?? []).forEach((m) => rowMetaMap.set(m.visualRowNumber, m));
      const hasRowMeta = (sheet.rowsMeta ?? []).length > 0;
      sheet.cells.forEach((cell: MergeCell) => {
        const meta = rowMetaMap.get(cell.row);
        const targetRowNumber = meta?.oursRowNumber ?? null;
        const targetColNumber = cell.oursCol ?? null;
        if (hasRowMeta && !targetRowNumber) {
          skippedCells += 1;
          return;
        }
        if (!targetColNumber) {
          skippedCells += 1;
          return;
        }
        const address = targetRowNumber ? makeAddress(targetColNumber, targetRowNumber) : makeAddress(targetColNumber, cell.row);
        changedCells.push({
          sheetName: sheet.sheetName,
          address,
          ours: cell.oursValue,
          theirs: cell.theirsValue,
          merged: cell.mergedValue,
        });
      });
    });

    const formatVal = (v: any) => (v === null || v === undefined ? '' : String(v));

    const maxLines = 100;
    const lines = changedCells.slice(0, maxLines).map((c) =>
      `[${c.sheetName}] 单元格 ${c.address}: ours="${formatVal(c.ours)}"  |  theirs="${formatVal(
        c.theirs,
      )}"  |  合并="${formatVal(c.merged)}"`,
    );

    if (changedCells.length > maxLines) {
      lines.push(`…… 还有 ${changedCells.length - maxLines} 个单元格未展示`);
    }
    if (skippedCells > 0) {
      lines.push(`（提示：有 ${skippedCells} 个单元格因 ours 侧缺少对应行/列而未写入）`);
    }

    const preview =
      `本次合并将影响 ${changedCells.length} 个单元格（覆盖所有工作表）：` +
      (lines.length ? `\n\n${lines.join('\n')}` : '\n(无差异单元格——仅写回了当前值)') +
      '\n\n注意：保存时会将所有工作表的合并结果一并写入目标 Excel 文件。' +
      '\n\n确认要将以上结果写入 Excel 文件吗？';

    const confirmed = window.confirm(preview);
    if (!confirmed) return;

    const cells = mergeSheets.flatMap((sheet: MergeSheetData) => {
      const rowMetaMap = new Map<number, MergeRowMeta>();
      (sheet.rowsMeta ?? []).forEach((m) => rowMetaMap.set(m.visualRowNumber, m));
      const hasRowMeta = (sheet.rowsMeta ?? []).length > 0;
      return sheet.cells
        .map((cell: MergeCell) => {
          const meta = rowMetaMap.get(cell.row);
          const targetRowNumber = meta?.oursRowNumber ?? null;
          if (hasRowMeta && !targetRowNumber) return null;
          const targetColNumber = cell.oursCol ?? null;
          if (!targetColNumber) return null;
          const address = targetRowNumber ? makeAddress(targetColNumber, targetRowNumber) : makeAddress(targetColNumber, cell.row);
          return {
            sheetName: sheet.sheetName,
            address,
            value: cell.mergedValue,
          };
        })
        .filter(Boolean) as { sheetName: string; address: string; value: string | number | null }[];
    });
    // 构建 aligned → physical 列映射：考虑列删除和列插入后，物理工作表的列布局
    const buildPhysicalColMap = (
      colsMeta: MergeColumnMeta[],
      colOpsMap: Map<number, SaveMergeColOp>,
    ): number[] => {
      const rawColCount = colsMeta.reduce((m, c) => Math.max(m, c.col), 0);
      const deletedCols = new Set<number>();
      const insertedCols: number[] = [];
      colOpsMap.forEach((op, ac) => {
        if (op.action === 'delete') deletedCols.add(ac);
        else if (op.action === 'insert') insertedCols.push(ac);
      });
      const map: number[] = [];
      for (let c = 1; c <= rawColCount; c += 1) {
        if (deletedCols.has(c)) continue;
        const m = colsMeta.find((cm) => cm.col === c);
        if (!m?.oursCol) continue;
        map.push(c);
      }
      insertedCols.sort((a, b) => a - b);
      const ins: Array<{ idx: number; col: number }> = [];
      for (const ac of insertedCols) {
        const m = colsMeta.find((cm) => cm.col === ac);
        if (m && !m.oursCol && m.theirsCol) {
          let insertIdx = map.length;
          for (let k = 0; k < map.length; k += 1) {
            if (map[k] > ac) { insertIdx = k; break; }
          }
          ins.push({ idx: insertIdx, col: ac });
        }
      }
      ins.sort((a, b) => b.idx - a.idx || b.col - a.col);
      for (const entry of ins) map.splice(entry.idx, 0, entry.col);
      return map;
    };

    const rowOps = Array.from(mergeRowOpsBySheet.entries()).flatMap(([sheetIndex, opsMap]) => {
      const sheet = mergeSheets[sheetIndex];
      const sheetName = sheet?.sheetName ?? mergeInfo.sheetName;
      const colsMeta = sheet?.columnsMeta ?? [];
      const colOpsForSheet = mergeColOpsBySheet.get(sheetIndex) ?? new Map<number, SaveMergeColOp>();
      // 将 row op 的 values 从 aligned 列空间重映射到物理列空间，
      // 跳过 theirs-only 列（除非用户选择了插入）和已删除的列。
      const colMap = colsMeta.length > 0 ? buildPhysicalColMap(colsMeta, colOpsForSheet) : null;
      return Array.from(opsMap.values()).map((op) => ({
        ...op,
        sheetName: op.sheetName || sheetName,
        values: op.values && colMap
          ? colMap.map((ac) => op.values![ac - 1] ?? null)
          : op.values,
      }));
    });
    const buildMergedColumnValues = (sheet: MergeSheetData, alignedColNumber: number) => {
      const rowsMeta = sheet.rowsMeta ?? [];
      const rowMetaMap = new Map<number, MergeRowMeta>();
      rowsMeta.forEach((m) => rowMetaMap.set(m.visualRowNumber, m));
      const maxRow = rowsMeta.reduce((m, r) => Math.max(m, r.oursRowNumber ?? 0), 0);
      const values: (string | number | null)[] = Array(maxRow).fill(null);
      sheet.cells.forEach((cell) => {
        if (cell.col !== alignedColNumber) return;
        const meta = rowMetaMap.get(cell.row);
        if (!meta?.oursRowNumber) return;
        values[meta.oursRowNumber - 1] = cell.mergedValue ?? null;
      });
      return values;
    };
    const colOps = Array.from(mergeColOpsBySheet.entries()).flatMap(([sheetIndex, opsMap]) => {
      const sheet = mergeSheets[sheetIndex];
      const sheetName = sheet?.sheetName ?? mergeInfo.sheetName;
      return Array.from(opsMap.values()).map((op) => ({
        ...op,
        sheetName: op.sheetName || sheetName,
        values: sheet && op.alignedColNumber ? buildMergedColumnValues(sheet, op.alignedColNumber) : op.values,
      }));
    });

    const payload: SaveMergeRequest = {
      templatePath: mergeInfo.oursPath,
      cells,
      rowOps,
      colOps,
      basePath: mergeInfo.basePath,
      oursPath: mergeInfo.oursPath,
      theirsPath: mergeInfo.theirsPath,
    };

    try {
      const result = await window.excelAPI.saveMergeResult(payload);
      if (!result.success || result.cancelled) {
        const msg = result.errorMessage ?? '未知错误，可能是目标文件被占用或没有写入权限。';
        alert(`保存合并结果失败：${msg}`);
        return;
      }

      alert(`合并结果已保存到: ${result.filePath ?? ''}`);
    } catch (e) {
      alert(`保存合并结果失败：${String(e)}`);
    }
  }, [mergeInfo, mergeSheets, mergeRowOpsBySheet, mergeColOpsBySheet]);
  const mergedPreviewScrollToCell = useMemo(() => {
    if (!selectedMergeCell) return null;
    const visualRowNumber = selectedMergeCell.rowIndex + 1;
    const rowIndex = mergedPreviewRowVisuals.indexOf(visualRowNumber);
    if (rowIndex < 0) return null;
    const alignedCol = selectedMergeCell.colIndex + 1;
    const colIndex = mergedPreviewAlignedCols.indexOf(alignedCol);
    if (colIndex < 0) return null;
    return { rowIndex, colIndex };
  }, [mergedPreviewAlignedCols, selectedMergeCell, mergedPreviewRowVisuals]);
  const renderMergedPreviewRowHeader = (rowIndex: number) => {
    const visual = mergedPreviewRowVisuals[rowIndex];
    return visual == null ? '' : visual;
  };
  const renderMergedPreviewHeaderCell = (colIndex: number) =>
    colNumberToLabel(mergedPreviewAlignedCols[colIndex] ?? colIndex + 1);
  const renderMergedPreviewCell = (cell: string | number | null, ctx: any) => {
    const visualRowNumber = mergedPreviewRowVisuals[ctx.rowIndex];
    const alignedCol = mergedPreviewAlignedCols[ctx.colIndex] ?? ctx.colIndex + 1;
    const mergeCell = visualRowNumber == null ? null : (mergeCellsByRow.get(visualRowNumber) ?? []).find((item) => item.col === alignedCol) ?? null;
    const resolved = visualRowNumber == null ? false : (resolvedBySheet.get(selectedMergeSheetIndex)?.has(`${visualRowNumber}:${alignedCol}`) ?? false);
    const value = cell == null ? '' : String(cell);
    return (
      <div
        onMouseDown={() => {
          if (visualRowNumber == null) return;
          setSelectedMergeCell({ rowIndex: visualRowNumber - 1, colIndex: alignedCol - 1 });
        }}
        title={
          mergeCell
            ? `${value}\n状态: ${mergeCell.status}\n${resolved ? '已确认' : '待确认'}`
            : value
        }
        style={{
          width: '100%',
          height: '100%',
          boxSizing: 'border-box',
          backgroundColor: 'transparent',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          cursor: 'pointer',
          userSelect: 'none',
        }}
      >
        {value}
      </div>
    );
  };
  const getMergedPreviewCellStyle = (_cell: any, ctx: any): React.CSSProperties => {
    const style: React.CSSProperties = {};
    const visualRowNumber = mergedPreviewRowVisuals[ctx.rowIndex];
    const alignedCol = mergedPreviewAlignedCols[ctx.colIndex] ?? ctx.colIndex + 1;
    const mergeCell =
      visualRowNumber == null ? null : (mergeCellsByRow.get(visualRowNumber) ?? []).find((item) => item.col === alignedCol) ?? null;
    const resolved =
      visualRowNumber == null
        ? false
        : (resolvedBySheet.get(selectedMergeSheetIndex)?.has(`${visualRowNumber}:${alignedCol}`) ?? false);
    if (ctx.isFrozenRow || ctx.isFrozenCol) {
      style.backgroundColor = '#f5f5f5';
    }
    if (mergeCell) {
      if (resolved) {
        style.backgroundColor = '#f0f0f0';
      } else if (mergeCell.status === 'conflict') {
        style.backgroundColor = '#fff3e0';
        style.boxShadow = 'inset 0 0 0 2px #ff8a00';
      } else if (mergeCell.status === 'both-changed-same') {
        style.backgroundColor = '#f5f5f5';
      } else if (mergeCell.status === 'ours-changed') {
        style.backgroundColor = '#e9f8e9';
      } else if (mergeCell.status === 'theirs-changed') {
        style.backgroundColor = '#ffe9e9';
      }
    }
    if (selectedMergeCell) {
      if (
        visualRowNumber === selectedMergeCell.rowIndex + 1 &&
        alignedCol === selectedMergeCell.colIndex + 1
      ) {
        style.outline = '2px solid #ff8000';
        style.outlineOffset = '-2px';
        style.position = 'relative';
        style.zIndex = 6;
      }
    }
    return style;
  };
  const mergedPreviewSafeRows = useMemo(() => {
    const minRows = Math.max(1, Math.floor(mergedPreviewMinRows));
    if (!mergedPreviewRows || mergedPreviewRows.length === 0) {
      return Array.from({ length: minRows }, () => [null]);
    }
    const first = mergedPreviewRows[0];
    const hasCols = Array.isArray(first) && first.length > 0;
    if (!hasCols) {
      return mergedPreviewRows.map(() => [null]);
    }
    return mergedPreviewRows;
  }, [mergedPreviewRows, mergedPreviewMinRows]);
  const mergedPreviewMinHeight = Math.max(mergedPreviewMinRows, 1) * 24 + 28;
  const primaryKeyHintColor =
    primaryKeyMode === 'manual' && activePrimaryKeySource === 'none'
      ? '#b00020'
      : primaryKeyMode === 'auto' && activePrimaryKeySource === 'none'
        ? '#8d6e00'
        : '#666';
  const primaryKeyControl = (
    <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 8, flexWrap: 'wrap' }}>
      <span>主键策略:</span>
      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12 }}>
        <input type="radio" name="primaryKeyMode" checked={primaryKeyMode === 'auto'} onChange={() => setPrimaryKeyMode('auto')} />
        自动识别
      </label>
      {primaryKeyMode === 'auto' && (
        <>
          <span>Auto 结果:</span>
          <input
            readOnly
            value={autoPrimaryKeyDisplayText}
            style={{ width: 220, padding: '2px 6px', boxSizing: 'border-box', color: '#333' }}
          />
        </>
      )}
      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12 }}>
        <input
          type="radio"
          name="primaryKeyMode"
          checked={primaryKeyMode === 'manual'}
          onChange={() => setPrimaryKeyMode('manual')}
        />
        手动指定
      </label>
      <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12 }}>
        <input type="radio" name="primaryKeyMode" checked={primaryKeyMode === 'none'} onChange={() => setPrimaryKeyMode('none')} />
        无主键
      </label>
      {primaryKeyMode === 'manual' && (
        <>
          <span>主键列:</span>
          <CommitNumberInput
            value={Math.max(1, Math.floor(primaryKeyCol || 1))}
            min={1}
            onCommit={(value) => setPrimaryKeyCol(value)}
          />
          <span style={{ fontSize: 12, color: '#666' }}>（1=A 列，2=B 列…）</span>
        </>
      )}
      <span style={{ fontSize: 12, color: primaryKeyHintColor }}>
        {primaryKeyHint || '主键用于稳定“同一业务行”的身份；调整后会重新比较。'}
      </span>
    </div>
  );

  return (
    <div
      style={{
        padding: 16,
        fontFamily: 'sans-serif',
        height: '100vh',
        boxSizing: 'border-box',
        display: 'flex',
        flexDirection: 'column',
        overflow: 'hidden',
      }}
    >
      <div style={{ marginBottom: 12, display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
        <button onClick={() => setMode('diff')} disabled={mode === 'diff'}>
          Diff 模式
        </button>
        <button onClick={() => setMode('merge')} disabled={mode === 'merge'}>
          Merge 模式
        </button>
        {mode === 'diff' && (
          <>
            <button
              onClick={() => handleSaveDiffSide('left')}
              disabled={!diffLeftWorkbook || diffLeftPendingCount === 0 || diffSavingSide !== null}
            >
              {diffSavingSide === 'left' ? '保存左侧中…' : `保存左侧 (${diffLeftPendingCount})`}
            </button>
            <button
              onClick={() => handleSaveDiffSide('right')}
              disabled={!diffRightWorkbook || diffRightPendingCount === 0 || diffSavingSide !== null}
            >
              {diffSavingSide === 'right' ? '保存右侧中…' : `保存右侧 (${diffRightPendingCount})`}
            </button>
            <span style={{ fontSize: 12, color: '#666' }}>
              默认主流程：选择左右两个 Excel，左右并排对比，双击单元格可直接编辑。
            </span>
          </>
        )}
        {mode === 'merge' && hasMergeData && mergeInfo && (
          <>
            <button onClick={handleSaveMergeToFile} style={{ marginLeft: 8 }}>
              {cliInfo?.mode === 'merge'
                ? '将合并结果写回 Git 合并文件（MERGED，解决冲突）'
                : cliInfo?.mode === 'diff'
                ? '将合并结果覆盖 ours（当前分支）文件'
                : '保存合并结果为新的 Excel 文件（以 ours 为格式模板）'}
            </button>
            <span style={{ marginLeft: 8, fontSize: 12, color: '#666' }}>
              {cliInfo
                ? '（本次操作会将所有工作表的合并结果写入 Git 传入的目标文件，保存后回到 Git 执行 git add 即可完成冲突解决）'
                : '（注意：保存时会将所有工作表的合并结果一并写入目标文件）'}
            </span>
          </>
        )}
      </div>

      {mode === 'diff' && (
        <div
          style={{
            marginBottom: 12,
            padding: 12,
            border: '1px solid #dcdcdc',
            borderRadius: 10,
            display: 'grid',
            gap: 8,
            flexShrink: 0,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
            <div style={{ fontSize: 12, color: '#555' }}>文件选择（支持直接粘贴路径）</div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <div
                  style={{
                    width: 160,
                    height: 8,
                    borderRadius: 999,
                    border: '1px solid #cbd5e1',
                    backgroundColor: '#eef2f7',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    style={{
                      width: `${Math.max(0, Math.min(100, diffAnalyzeProgress))}%`,
                      height: '100%',
                      backgroundColor: diffAnalyzeInProgress ? '#2563eb' : '#94a3b8',
                      transition: 'width 140ms linear',
                    }}
                  />
                </div>
                <span style={{ fontSize: 11, color: '#64748b', width: 52, textAlign: 'right' }}>
                  {diffAnalyzeInProgress ? `${Math.round(diffAnalyzeProgress)}%` : '待加载'}
                </span>
              </div>
              <button type="button" onClick={() => void handleLoadDiffFromInputs()} disabled={diffAnalyzeInProgress}>
                {diffAnalyzeInProgress ? '分析中...' : '加载'}
              </button>
              <button type="button" onClick={() => setDiffFileSelectorCollapsed((prev) => !prev)}>
                {diffFileSelectorCollapsed ? '展开' : '收起'}
              </button>
            </div>
          </div>
          {!diffFileSelectorCollapsed && (
            <>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <span style={{ width: 72, fontSize: 12, color: '#555' }}>左侧文件</span>
                <input
                  value={diffPathInputs.left}
                  placeholder="粘贴左侧 Excel 路径"
                  onChange={(e) =>
                    setDiffPathInputs((prev) => ({
                      ...prev,
                      left: e.target.value,
                    }))
                  }
                  style={{ flex: 1, minWidth: 260, padding: '4px 6px', boxSizing: 'border-box' }}
                />
                <button type="button" onClick={() => handlePickDiffWorkbook('left')}>
                  选择左侧
                </button>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <span style={{ width: 72, fontSize: 12, color: '#555' }}>右侧文件</span>
                <input
                  value={diffPathInputs.right}
                  placeholder="粘贴右侧 Excel 路径"
                  onChange={(e) =>
                    setDiffPathInputs((prev) => ({
                      ...prev,
                      right: e.target.value,
                    }))
                  }
                  style={{ flex: 1, minWidth: 260, padding: '4px 6px', boxSizing: 'border-box' }}
                />
                <button type="button" onClick={() => handlePickDiffWorkbook('right')}>
                  选择右侧
                </button>
              </div>
            </>
          )}
        </div>
      )}

      {mode === 'merge' && (
        <div
          style={{
            marginBottom: 12,
            padding: 12,
            border: '1px solid #dcdcdc',
            borderRadius: 10,
            display: 'grid',
            gap: 8,
            flexShrink: 0,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
            <div style={{ fontSize: 12, color: '#555' }}>文件选择（支持直接粘贴路径）</div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <div
                  style={{
                    width: 160,
                    height: 8,
                    borderRadius: 999,
                    border: '1px solid #cbd5e1',
                    backgroundColor: '#eef2f7',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    style={{
                      width: `${Math.max(0, Math.min(100, mergeAnalyzeProgress))}%`,
                      height: '100%',
                      backgroundColor: mergeAnalyzeInProgress ? '#2563eb' : '#94a3b8',
                      transition: 'width 140ms linear',
                    }}
                  />
                </div>
                <span style={{ fontSize: 11, color: '#64748b', width: 52, textAlign: 'right' }}>
                  {mergeAnalyzeInProgress ? `${Math.round(mergeAnalyzeProgress)}%` : '待加载'}
                </span>
              </div>
              <button type="button" onClick={() => void handleLoadMergeFromInputs()} disabled={mergeAnalyzeInProgress}>
                {mergeAnalyzeInProgress ? '分析中...' : '加载'}
              </button>
              <button type="button" onClick={() => setMergeFileSelectorCollapsed((prev) => !prev)}>
                {mergeFileSelectorCollapsed ? '展开' : '收起'}
              </button>
            </div>
          </div>
          {!mergeFileSelectorCollapsed &&
            ([
              ['basePath', 'base', '粘贴 base Excel 路径'],
              ['oursPath', 'ours', '粘贴 ours Excel 路径'],
              ['theirsPath', 'theirs', '粘贴 theirs Excel 路径'],
            ] as const).map(([role, label, placeholder]) => (
              <div key={role} style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <span style={{ width: 72, fontSize: 12, color: '#555' }}>{label}</span>
                <input
                  value={mergePathInputs[role]}
                  placeholder={placeholder}
                  onChange={(e) =>
                    setMergePathInputs((prev) => ({
                      ...prev,
                      [role]: e.target.value,
                    }))
                  }
                  style={{ flex: 1, minWidth: 260, padding: '4px 6px', boxSizing: 'border-box' }}
                />
                <button type="button" onClick={() => handlePickMergeWorkbook(role)}>
                  {`选择 ${label}`}
                </button>
              </div>
            ))}
        </div>
      )}

      {mode === 'diff' && (
        <div
          style={{
            marginBottom: 12,
            padding: 12,
            border: '1px solid #dcdcdc',
            borderRadius: 10,
            display: 'grid',
            gap: 8,
            flexShrink: 0,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
            <div style={{ fontSize: 12, color: '#555' }}>高级参数</div>
            <button type="button" onClick={() => setDiffAdvancedCollapsed((prev) => !prev)}>
              {diffAdvancedCollapsed ? '展开' : '收起'}
            </button>
          </div>
          {!diffAdvancedCollapsed && (
            <>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span>diff 视图冻结行数:</span>
                  <input
                    type="number"
                    min={0}
                    value={mergeFrozenRowDraft}
                    onChange={(e) => handleMergeFrozenRowDraftChange(e.target.value, 'diff')}
                    style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
                  />
                  <button type="button" onClick={applyMergeFrozenRowDraft} disabled={!canRefreshFrozenRows}>
                    刷新
                  </button>
                </div>
                <span style={{ fontSize: 12, color: '#666' }}>当前剩余差异: {diffRemainingCount}</span>
                <span style={{ fontSize: 12, color: '#666' }}>左侧未保存修改: {diffLeftPendingCount}</span>
                <span style={{ fontSize: 12, color: '#666' }}>右侧未保存修改: {diffRightPendingCount}</span>
              </div>
              {primaryKeyControl}
            </>
          )}
        </div>
      )}

      {mode === 'merge' && (
        <div
          style={{
            marginBottom: 12,
            padding: 12,
            border: '1px solid #dcdcdc',
            borderRadius: 10,
            display: 'grid',
            gap: 8,
            flexShrink: 0,
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
            <div style={{ fontSize: 12, color: '#555' }}>高级参数</div>
            <button type="button" onClick={() => setMergeAdvancedCollapsed((prev) => !prev)}>
              {mergeAdvancedCollapsed ? '展开' : '收起'}
            </button>
          </div>
          {!mergeAdvancedCollapsed && (
            <>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span>merge 视图冻结行数:</span>
                  <input
                    type="number"
                    min={0}
                    value={mergeFrozenRowDraft}
                    onChange={(e) => handleMergeFrozenRowDraftChange(e.target.value, 'merge')}
                    style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
                  />
                  <button type="button" onClick={applyMergeFrozenRowDraft} disabled={!canRefreshFrozenRows}>
                    刷新
                  </button>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span>行相似度阈值:</span>
                  <CommitNumberInput
                    value={rowSimilarityThreshold}
                    min={0}
                    max={1}
                    step={0.01}
                    onCommit={(value) => setRowSimilarityThreshold(value)}
                  />
                  <span style={{ fontSize: 12, color: '#666' }}>（0~1，越大越严格）</span>
                </div>
                <span style={{ fontSize: 12, color: '#666' }}>当前未解决冲突: {mergeRemainingCount}</span>
                <span style={{ fontSize: 12, color: '#666' }}>本表差异单元格: {mergeCells.length}</span>
              </div>
              {primaryKeyControl}
            </>
          )}
        </div>
      )}


      {/* 主内容：表格 / 三方 Merge，占用剩余空间，由内部自己滚动 */}
      <div
        style={{
          flex: 1,
          minHeight: 0,
          overflow: 'hidden',
          display: 'flex',
          flexDirection: 'column',
        }}
      >

      {mode === 'diff' && (
        <div style={{ flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
          <div style={{ marginBottom: 8 }}>
            <div style={{ marginTop: 4, fontSize: 12, display: 'flex', gap: 12, flexWrap: 'wrap', color: '#666' }}>
              <span>当前剩余差异: {diffRemainingCount}</span>
              <span>左侧未保存修改: {diffLeftPendingCount}</span>
              <span>右侧未保存修改: {diffRightPendingCount}</span>
              <span>无对应行/列的位置仅作占位显示，不能直接编辑。</span>
            </div>
            <div
              style={{
                display: 'flex',
                alignItems: 'flex-start',
                gap: 12,
                marginTop: 8,
                flexWrap: 'wrap',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <span style={{ fontSize: 12 }}>单元格地址</span>
                <input
                  readOnly
                  value={currentCellAddress}
                  placeholder="例如 A1"
                  style={{ width: 90, padding: '2px 6px', boxSizing: 'border-box' }}
                />
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <span style={{ fontSize: 12, whiteSpace: 'nowrap' }}>left</span>
                <input
                  readOnly
                  value={selectedDiffCellData?.leftValue == null ? '' : String(selectedDiffCellData.leftValue)}
                  style={{ width: 260, padding: '2px 6px', boxSizing: 'border-box' }}
                />
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <span style={{ fontSize: 12, whiteSpace: 'nowrap' }}>right</span>
                <input
                  readOnly
                  value={selectedDiffCellData?.rightValue == null ? '' : String(selectedDiffCellData.rightValue)}
                  style={{ width: 260, padding: '2px 6px', boxSizing: 'border-box' }}
                />
              </div>
            </div>
          </div>
          <div style={{ flex: 1, minHeight: 0 }}>
            {diffLeftWorkbook && diffRightWorkbook ? (
              currentDiffSheet ? (
                <DiffSideBySide
                  leftPath={diffLeftWorkbook.filePath}
                  rightPath={diffRightWorkbook.filePath}
                  leftRows={currentDiffLeftRows}
                  rightRows={currentDiffRightRows}
                  rowsMeta={currentDiffRowsMeta}
                  columnsMeta={currentDiffColumnsMeta}
                  frozenRowCount={mergeFrozenRowCount}
                  selected={diffSelectedCell}
                  onSelectCell={(rowIndex, colIndex) => setDiffSelectedCell({ rowIndex, colIndex })}
                  onCellChange={handleDiffCellChange}
                  onApplyOtherSideCell={handleApplyDiffOtherCell}
                  onApplyOtherSideRow={handleApplyDiffOtherRow}
                  onDeleteRow={handleDeleteDiffRow}
                />
              ) : (
                <div>没有可对比的工作表（左右文件没有同名工作表）。</div>
              )
            ) : (
              <div>请先在上方选择左右两个 Excel 文件。</div>
            )}
          </div>
          <div
            style={{
              marginTop: 8,
              paddingTop: 8,
              borderTop: '1px solid #eee',
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              flexWrap: 'wrap',
              flexShrink: 0,
            }}
          >
            <span style={{ fontSize: 12, color: '#555' }}>工作表</span>
            <div style={{ display: 'inline-flex', borderBottom: '1px solid #ccc', gap: 4, flexWrap: 'wrap' }}>
              {diffSheets.map((sheet, idx) => {
                const isActive = idx === selectedDiffSheetIndex;
                return (
                  <button
                    key={sheet.sheetName || idx}
                    type="button"
                    onClick={() => {
                      setSelectedDiffSheetIndex(idx);
                      setDiffSelectedCell(null);
                    }}
                    style={{
                      padding: '2px 8px',
                      fontSize: 12,
                      borderRadius: '4px 4px 0 0',
                      border: '1px solid #ccc',
                      borderBottom: isActive ? '2px solid white' : '1px solid #ccc',
                      backgroundColor: isActive ? '#ffffff' : '#f5f5f5',
                      cursor: 'pointer',
                    }}
                  >
                    {sheet.sheetName || `Sheet${idx + 1}`}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {mode === 'merge' && (
        <div style={{ flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
          {mergeInfo && mergeSheets.length === 0 ? (
            <div>没有可对比的工作表（base / ours / theirs 中没有任何“同名工作表”的交集）。</div>
          ) : mergeInfo ? (
            <div style={{ flex: 1, minHeight: 0, display: 'flex', gap: 8 }}>
              <div style={{ flex: 1, minWidth: 0, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
                <div style={{ flex: 1, minHeight: 0 }}>
                  <MergeWorkbench
                    cells={mergeCells}
                    rowsMeta={mergeRowsMeta}
                    columnsMeta={mergeColumnsMeta}
                    sourceRows={mergeThreeWayRows}
                    layoutMode="grids-only"
                    selected={selectedMergeCell}
                    onSelectCell={handleSelectMergeCell}
                    onApplySelectedCellChoice={handleApplyMergeChoice}
                    onApplyCellsChoice={handleApplyMergeCellsChoice}
                    onResolveCell={handleResolveMergeCell}
                    onApplyRowChoice={handleApplyMergeRowChoice}
                    onDeleteRow={handleDeleteMergeRow}
                    resolvedCellKeys={resolvedBySheet.get(selectedMergeSheetIndex)}
                    frozenRowCount={mergeFrozenRowCount}
                    primaryKeyCol={displayPrimaryKeyCol}
                    sheetName={mergeInfo?.sheetName ?? ''}
                    basePath={mergeInfo?.basePath ?? null}
                    oursPath={mergeInfo?.oursPath ?? null}
                    theirsPath={mergeInfo?.theirsPath ?? null}
                    mergedPath={mergedPath}
                    remainingCount={mergeRemainingCount}
                    canUndo={mergeUndoStack.length > 0}
                    onUndo={handleUndoMergeAction}
                  />
                </div>
                <div style={{ marginTop: 8, fontSize: 12, color: '#666', flexShrink: 0 }}>
                  merged 结果已整合到工作台主区域（第四栏），冲突位置会直接高亮并可一键处理。
                </div>
                {mergeSheets.length > 0 && (
                  <div
                    style={{
                      marginTop: 8,
                      paddingTop: 8,
                      borderTop: '1px solid #eee',
                      display: 'flex',
                      alignItems: 'center',
                      gap: 8,
                      flexWrap: 'wrap',
                      flexShrink: 0,
                    }}
                  >
                    <span style={{ fontSize: 12, color: '#555' }}>工作表</span>
                    <div style={{ display: 'inline-flex', borderBottom: '1px solid #ccc', gap: 4, flexWrap: 'wrap' }}>
                      {mergeSheets.map((s, idx) => {
                        const isActive = idx === selectedMergeSheetIndex;
                        const hasDiff =
                          typeof s.hasExactDiff === 'boolean' ? s.hasExactDiff : (s.cells?.length ?? 0) > 0;
                        return (
                          <button
                            key={s.sheetName || idx}
                            type="button"
                            onClick={() => {
                              setSelectedMergeSheetIndex(idx);
                              const sheet = mergeSheets[idx];
                              setMergeInfo((prev) =>
                                prev
                                  ? {
                                      ...prev,
                                      sheetName: sheet?.sheetName ?? prev.sheetName,
                                    }
                                  : prev,
                              );
                              setMergeCells(sheet?.cells ?? []);
                              setMergeRowsMeta(sheet?.rowsMeta ?? []);
                              setMergeColumnsMeta(sheet?.columnsMeta ?? []);
                              setResolvedBySheet((prev) => {
                                if (prev.has(idx)) return prev;
                                const next = new Map(prev);
                                const resolved = new Set<string>();
                                (sheet?.cells ?? []).forEach((cell) => {
                                  if (cell.status !== 'conflict') {
                                    resolved.add(`${cell.row}:${cell.col}`);
                                  }
                                });
                                next.set(idx, resolved);
                                return next;
                              });
                              setSelectedMergeCell(null);
                            }}
                            style={{
                              padding: '2px 8px',
                              fontSize: 12,
                              borderRadius: '4px 4px 0 0',
                              border: '1px solid #ccc',
                              borderBottom: isActive ? '2px solid white' : '1px solid #ccc',
                              backgroundColor: isActive ? '#ffffff' : '#f5f5f5',
                              cursor: 'pointer',
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: 6,
                            }}
                          >
                            {hasDiff && (
                              <span
                                title="该工作表有内容变动"
                                style={{
                                  width: 8,
                                  height: 8,
                                  backgroundColor: '#d32f2f',
                                  borderRadius: 2,
                                  display: 'inline-block',
                                }}
                              />
                            )}
                            {s.sheetName || `Sheet${idx + 1}`}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
              <div style={{ width: 460, minWidth: 420, minHeight: 0, display: 'flex' }}>
                <MergeWorkbench
                  cells={mergeCells}
                  rowsMeta={mergeRowsMeta}
                  columnsMeta={mergeColumnsMeta}
                  sourceRows={mergeThreeWayRows}
                  layoutMode="panel-only"
                  selected={selectedMergeCell}
                  onSelectCell={handleSelectMergeCell}
                  onApplySelectedCellChoice={handleApplyMergeChoice}
                  onApplyCellsChoice={handleApplyMergeCellsChoice}
                  onResolveCell={handleResolveMergeCell}
                  onApplyRowChoice={handleApplyMergeRowChoice}
                  onDeleteRow={handleDeleteMergeRow}
                  resolvedCellKeys={resolvedBySheet.get(selectedMergeSheetIndex)}
                  frozenRowCount={mergeFrozenRowCount}
                  primaryKeyCol={displayPrimaryKeyCol}
                  sheetName={mergeInfo?.sheetName ?? ''}
                  basePath={mergeInfo?.basePath ?? null}
                  oursPath={mergeInfo?.oursPath ?? null}
                  theirsPath={mergeInfo?.theirsPath ?? null}
                  mergedPath={mergedPath}
                  remainingCount={mergeRemainingCount}
                  canUndo={mergeUndoStack.length > 0}
                  onUndo={handleUndoMergeAction}
                />
              </div>
            </div>
          ) : (
            <div style={{ marginBottom: 12 }}>请先在上方选择 base / ours / theirs 三个 Excel 文件。</div>
          )}
        </div>
      )}
      </div>
    </div>
  );
};

