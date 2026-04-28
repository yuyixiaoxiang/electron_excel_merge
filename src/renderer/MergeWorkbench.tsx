import React, { useEffect, useMemo, useRef, useState } from 'react';
import type {
  MergeCell,
  MergeColumnMeta,
  MergeRowMeta,
  RowStatus,
  ThreeWayCompareMode,
  ThreeWayRowResult,
} from '../main/preload';
import { VirtualGrid, VirtualGridRenderCtx } from './VirtualGrid';

type SourceSide = 'base' | 'ours' | 'theirs' | 'merged';

type MergeWorkbenchCell = {
  key: string;
  address: string;
  rowNumber: number;
  displayRowNumber: number;
  colNumber: number;
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  mergedValue: string | number | null;
  status: MergeCell['status'] | 'unchanged';
  resolved: boolean;
  isDiffCell: boolean;
  isContextCell: boolean;
  formulaControlled?: boolean;
  sharedControlled?: boolean;
  sharedControlMasterSheetName?: string | null;
  sharedControlIsMaster?: boolean;
};

export interface MergeWorkbenchProps {
  cells: MergeCell[];
  rowsMeta: MergeRowMeta[];
  columnsMeta?: MergeColumnMeta[];
  sourceRows: ThreeWayRowResult[];
  compareMode?: ThreeWayCompareMode;
  layoutMode?: 'full' | 'grids-only' | 'panel-only';
  selected?: { rowIndex: number; colIndex: number } | null;
  onSelectCell?: (rowIndex: number, colIndex: number) => void;
  onApplySelectedCellChoice?: (source: 'base' | 'ours' | 'theirs') => void;
  onApplyCellsChoice?: (
    keys: Array<{ rowNumber: number; colNumber: number }>,
    source: 'base' | 'ours' | 'theirs',
  ) => void;
  onResolveCell?: (rowNumber: number, colNumber: number) => void;
  onApplyRowChoice?: (rowNumber: number, source: 'ours' | 'theirs') => void;
  onApplyColumnChoice?: (colNumber: number, source: 'ours' | 'theirs') => void;
  onDeleteRow?: (rowNumber: number) => void;
  resolvedCellKeys?: Set<string>;
  frozenRowCount?: number;
  primaryKeyCol?: number;
  sheetName?: string;
  basePath?: string | null;
  oursPath?: string | null;
  theirsPath?: string | null;
  mergedPath?: string | null;
  fullBaseRows?: (string | number | null)[][];
  fullOursRows?: (string | number | null)[][];
  fullTheirsRows?: (string | number | null)[][];
  mergedPreviewRows?: (string | number | null)[][];
  mergedPreviewRowVisuals?: (number | null)[];
  mergedPreviewAlignedCols?: number[];
  onSaveMergeResult?: () => void;
  saveMergeResultLabel?: string;
  remainingCount: number;
  canUndo?: boolean;
  onUndo?: () => void;
  canJumpToPreviousConflict?: boolean;
  onJumpToPreviousConflict?: () => void;
  canJumpToNextConflict?: boolean;
  onJumpToNextConflict?: () => void;
  showTheirsChangedReviewFallback?: boolean;
}

const DEFAULT_FROZEN_HEADER_ROWS = 3;
const DEFAULT_COL_WIDTH = 108;
const GRID_ROW_HEIGHT = 24;
const SYNC_SCROLL_TOLERANCE_PX = 1;
const SYNC_SCROLL_EVENT_TTL_MS = 160;
const ROW_HEADER_WIDTH = 62;
const FROZEN_BG = '#f2f4f7';
const BASE_BG = '#fff9e8';
const OURS_BG = '#ecf8ec';
const THEIRS_BG = '#fff0f0';
const MERGED_BG = '#e8efff';
const RESOLVED_BG = '#f5f5f5';
const CONFLICT_OUTLINE = '#ff8a00';
const PENDING_REVIEW_BG = '#fff1e6';
const PENDING_REVIEW_OUTLINE = '#ea580c';
const AUTO_MERGED_BG = '#b9d7ff';
const AUTO_MERGED_SUBTLE_BG = '#dbeafe';
const FORMULA_BG = '#e5e7eb';
const FORMULA_BORDER = '#9ca3af';
const FORMULA_TEXT = '#4b5563';

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
const makeCellAddress = (colNumber: number, rowNumber: number): string =>
  `${colNumberToLabel(colNumber)}${rowNumber}`;

const getRowStatusLabel = (status: RowStatus | undefined): string => {
  switch (status) {
    case 'added':
      return '新增';
    case 'deleted':
      return '缺失';
    case 'modified':
      return '改动';
    case 'ambiguous':
      return '歧义';
    case 'unchanged':
    default:
      return '稳定';
  }
};

const getCellStatusLabel = (status: MergeCell['status'] | 'unchanged'): string => {
  switch (status) {
    case 'ours-changed':
      return '自动并入';
    case 'theirs-changed':
      return '自动并入';
    case 'both-changed-same':
      return '自动并入';
    case 'conflict':
      return '冲突';
    case 'unchanged':
    default:
      return '无差异';
  }
};

const getDisplayedCellStatusLabel = (cell: MergeWorkbenchCell | null | undefined): string => {
  const mode = getProtectedCellMode(cell);
  if (mode === 'formula') return '公式控制';
  if (mode === 'shared') return '共享控制';
  if (cell?.isContextCell) return '无差异（上下文）';
  return getCellStatusLabel(cell?.status ?? 'unchanged');
};

const truncatePath = (value?: string | null): string => {
  if (!value) return '';
  if (value.length <= 48) return value;
  return `…${value.slice(-47)}`;
};

const getPanelAccent = (side: SourceSide) => {
  switch (side) {
    case 'base':
      return '#c58a00';
    case 'ours':
      return '#2e7d32';
    case 'theirs':
      return '#c62828';
    case 'merged':
      return '#1d4ed8';
  }
};

const getPanelBackground = (
  status: MergeWorkbenchCell['status'],
  side: SourceSide,
  resolved: boolean,
  isPendingReview: boolean,
  isAutoMerged: boolean,
) => {
  if (side === 'merged') {
    if (status === 'conflict' && !resolved) return '#fff1e6';
    if (isPendingReview) return PENDING_REVIEW_BG;
    if (isAutoMerged && !resolved) return AUTO_MERGED_BG;
    if (isAutoMerged || status !== 'unchanged') return AUTO_MERGED_SUBTLE_BG;
    return 'white';
  }
  if (resolved && status !== 'unchanged') return RESOLVED_BG;
  if (status === 'unchanged') return 'white';
  if (side === 'base') return BASE_BG;
  if (status === 'ours-changed') return side === 'ours' ? '#d4f8d4' : 'white';
  if (status === 'theirs-changed') return side === 'theirs' ? '#ffd6d6' : 'white';
  if (status === 'both-changed-same') return '#fafafa';
  return side === 'ours' ? '#d4f8d4' : side === 'theirs' ? '#ffc8c8' : BASE_BG;
};

const getProtectedCellHint = (cell: MergeWorkbenchCell | null | undefined): string | null => {
  const mode = getProtectedCellMode(cell);
  if (mode === 'formula') return '公式控制位：这个位置不能直接编辑，保存时会保留模板里的公式。';
  if (mode === 'shared') {
    const masterSheetName = getSharedControlMasterSheetName(cell);
    return masterSheetName
      ? `共享控制位：这个位置由 ${masterSheetName} sheet 统一控制，不能单独编辑。`
      : '共享控制位：这个位置在多个工作表里同步变化，不能单独编辑。';
  }
  return null;
};

const getDefaultMergedValue = (
  cell: MergeCell | undefined,
  row: ThreeWayRowResult | undefined,
  colNumber: number,
  resolved: boolean,
) => {
  if (cell?.mergedValue !== null && cell?.mergedValue !== undefined) {
    return cell.mergedValue;
  }
  if (resolved) {
    return cell?.mergedValue ?? null;
  }
  if (!row) return cell?.mergedValue ?? null;
  const oursValue = row.ours[colNumber - 1] ?? null;
  const theirsValue = row.theirs[colNumber - 1] ?? null;
  const baseValue = row.base[colNumber - 1] ?? null;
  if (oursValue !== null && oursValue !== undefined) return oursValue;
  if (theirsValue !== null && theirsValue !== undefined) return theirsValue;
  return baseValue;
};

const describeCellDecision = (
  cell: MergeWorkbenchCell,
  rowMeta?: MergeRowMeta,
  compareMode: ThreeWayCompareMode = 'merge',
) => {
  const prefix =
    rowMeta && (rowMeta.oursStatus === 'ambiguous' || rowMeta.theirsStatus === 'ambiguous')
      ? '该行对齐存在歧义，先看清三侧原始值再决定。'
      : '';
  const isSimpleMergeMode = compareMode === 'simple-merge';
  if (isFormulaControlledCell(cell)) {
    return `${
      prefix
    }这个位置由模板公式控制，不能直接采用 ${
      isSimpleMergeMode ? 'ours / theirs' : 'base / ours / theirs'
    } 的文本结果；保存时会保留 ours 模板里的公式。`.trim();
  }
  if (isSharedControlledCell(cell)) {
    const masterSheetName = getSharedControlMasterSheetName(cell);
    return `${
      prefix
    }这个位置属于共享控制位，不能单独采用 ${
      isSimpleMergeMode ? 'ours / theirs' : 'base / ours / theirs'
    }；${
      masterSheetName ? `请去 ${masterSheetName} sheet 的主位修改。` : '需要跟随共享主位一起变化。'
    }`.trim();
  }
  switch (cell.status) {
    case 'ours-changed':
      return `${prefix}这个位置归类为自动并入：ours 相对 base 发生变化，theirs 保持与 base 一致；如果你认可当前分支改动，直接采用当前 merged 结果即可。`.trim();
    case 'theirs-changed':
      return `${prefix}这个位置归类为自动并入：theirs 相对 base 发生变化，ours 保持与 base 一致；如果你确认要把对方改动并入结果，优先采用 theirs。`.trim();
    case 'both-changed-same':
      return `${prefix}这个位置归类为自动并入：ours 和 theirs 都相对 base 改了，但结果相同；系统已经自动给出同一 merged 值，你只需要确认。`.trim();
    case 'conflict':
      return isSimpleMergeMode
        ? `${prefix}这个位置归类为冲突：ours 和 theirs 的内容不同；这是人工决策点，需要你在 ours / theirs 之间做选择。`.trim()
        : `${prefix}这个位置归类为冲突：ours 和 theirs 都相对 base 改了，而且结果不同；这是人工决策点，需要你在 base / ours / theirs 之间做选择。`.trim();
    case 'unchanged':
    default:
      if (cell.isContextCell) {
        return rowMeta && (rowMeta.oursStatus === 'ambiguous' || rowMeta.theirsStatus === 'ambiguous')
          ? '当前单元格本身没有值差异；之所以显示在这里，是因为同列其他位置有差异，同时该行对齐也存在歧义。'
          : '当前单元格本身没有差异；之所以显示在这里，是因为同列其他位置有差异，这里只是保留给你做上下文对照。';
      }
      return rowMeta && (rowMeta.oursStatus === 'ambiguous' || rowMeta.theirsStatus === 'ambiguous')
        ? '当前单元格本身没有值冲突，但所在行的对齐并不稳定。'
        : isSimpleMergeMode
          ? '当前单元格在两侧没有形成需要人工处理的差异。'
          : '当前单元格在三侧没有形成需要人工处理的差异。';
  }
};

const describeRowDecision = (rowMeta?: MergeRowMeta, compareMode: ThreeWayCompareMode = 'merge') => {
  if (!rowMeta) return '当前没有选中行。';
  if (!rowMeta.baseRowNumber && rowMeta.oursRowNumber && !rowMeta.theirsRowNumber) {
    return '这是 ours 独有的新增行；如果不想保留它，可以删除该行。';
  }
  if (!rowMeta.baseRowNumber && !rowMeta.oursRowNumber && rowMeta.theirsRowNumber) {
    return '这是 theirs 独有的新增行；如果需要把它并入结果，可以采用 theirs 整行。';
  }
  if (rowMeta.baseRowNumber && rowMeta.oursRowNumber && !rowMeta.theirsRowNumber) {
    return 'theirs 缺少这行；如果你采用 theirs 行，结果会删除 ours 中对应的物理行。';
  }
  if (rowMeta.baseRowNumber && !rowMeta.oursRowNumber && rowMeta.theirsRowNumber) {
    return 'ours 侧缺少这行；如果你采用 theirs 行，结果会插入一行到 ours 模板中。';
  }
  if (rowMeta.oursStatus === 'ambiguous' || rowMeta.theirsStatus === 'ambiguous') {
    const oursSim = typeof rowMeta.oursSimilarity === 'number' ? rowMeta.oursSimilarity.toFixed(2) : '-';
    const theirsSim = typeof rowMeta.theirsSimilarity === 'number' ? rowMeta.theirsSimilarity.toFixed(2) : '-';
    return `该行是按内容相似度对齐的：ours=${oursSim}，theirs=${theirsSim}。如果值看起来不可信，优先按整行重新选择。`;
  }
  return compareMode === 'simple-merge'
    ? '这行已经完成两侧对齐；只要 ours 和 theirs 有不同，就需要你在 merged 中做选择。'
    : '这行已经完成三方对齐；一般只需要对其中的冲突单元格做选择。';
};

const makeValueText = (value: string | number | null) => (value == null ? '∅' : String(value));

const isFormulaControlledCell = (
  cell: Pick<MergeCell, 'formulaControlled'> | Pick<MergeWorkbenchCell, 'formulaControlled'> | null | undefined,
): boolean => cell?.formulaControlled === true;

const isSharedControlledCell = (
  cell: Pick<MergeCell, 'sharedControlled'> | Pick<MergeWorkbenchCell, 'sharedControlled'> | null | undefined,
): boolean => cell?.sharedControlled === true;

const getSharedControlMasterSheetName = (
  cell:
    | Pick<MergeCell, 'sharedControlMasterSheetName'>
    | Pick<MergeWorkbenchCell, 'sharedControlMasterSheetName'>
    | null
    | undefined,
): string | null => {
  const value = cell?.sharedControlMasterSheetName;
  return typeof value === 'string' && value.trim() ? value.trim() : null;
};

const getProtectedCellMode = (
  cell:
    | Pick<MergeCell, 'formulaControlled' | 'sharedControlled'>
    | Pick<MergeWorkbenchCell, 'formulaControlled' | 'sharedControlled'>
    | null
    | undefined,
): 'formula' | 'shared' | null => {
  if (isFormulaControlledCell(cell)) return 'formula';
  if (isSharedControlledCell(cell)) return 'shared';
  return null;
};

const isProtectedCell = (
  cell:
    | Pick<MergeCell, 'formulaControlled' | 'sharedControlled'>
    | Pick<MergeWorkbenchCell, 'formulaControlled' | 'sharedControlled'>
    | null
    | undefined,
): boolean => getProtectedCellMode(cell) !== null;

const isAttentionBucketCell = (
  cell: Pick<MergeCell, 'row' | 'col' | 'status' | 'formulaControlled' | 'sharedControlled'>,
  resolvedKeySet: Set<string>,
  includeTheirsChanged: boolean,
): boolean => {
  if (isProtectedCell(cell)) return false;
  if (resolvedKeySet.has(`${cell.row}:${cell.col}`)) return false;
  return cell.status === 'conflict';
};

const getTonePalette = (
  tone: 'base' | 'ours' | 'theirs' | 'merged' | 'danger' | 'neutral',
): { backgroundColor: string; borderColor: string; color: string } =>
  tone === 'base'
    ? { backgroundColor: BASE_BG, borderColor: '#d6b25e', color: '#8a5a00' }
    : tone === 'ours'
      ? { backgroundColor: OURS_BG, borderColor: '#7bbf85', color: '#166534' }
      : tone === 'theirs'
        ? { backgroundColor: THEIRS_BG, borderColor: '#f1a5a5', color: '#b91c1c' }
        : tone === 'merged'
          ? { backgroundColor: MERGED_BG, borderColor: '#93c5fd', color: '#1d4ed8' }
          : tone === 'danger'
            ? { backgroundColor: '#fff1f2', borderColor: '#fda4af', color: '#b42318' }
            : { backgroundColor: '#f8fafc', borderColor: '#cbd5e1', color: '#334155' };

const getQuickActionButtonStyle = (
  tone: 'base' | 'ours' | 'theirs' | 'merged' | 'danger' | 'neutral',
  disabled?: boolean,
): React.CSSProperties => {
  const palette = getTonePalette(tone);

  return {
    border: `1px solid ${palette.borderColor}`,
    backgroundColor: disabled ? '#f8fafc' : palette.backgroundColor,
    color: disabled ? '#94a3b8' : palette.color,
    borderRadius: 8,
    padding: '7px 10px',
    fontSize: 12,
    fontWeight: 600,
    cursor: disabled ? 'not-allowed' : 'pointer',
    opacity: disabled ? 0.7 : 1,
  };
};

const getCurrentValueCardPalette = (
  side: SourceSide,
  cell: MergeWorkbenchCell | null | undefined,
): { backgroundColor: string; borderColor: string; color: string; valueColor: string } => {
  if (isProtectedCell(cell)) {
    return {
      backgroundColor: FORMULA_BG,
      borderColor: FORMULA_BORDER,
      color: FORMULA_TEXT,
      valueColor: FORMULA_TEXT,
    };
  }
  const palette = getTonePalette(side);
  return {
    backgroundColor: palette.backgroundColor,
    borderColor: palette.borderColor,
    color: palette.color,
    valueColor: '#111827',
  };
};

const MergeWorkbenchComponent: React.FC<MergeWorkbenchProps> = ({
  cells,
  rowsMeta,
  columnsMeta,
  sourceRows,
  compareMode = 'merge',
  layoutMode = 'full',
  selected,
  onSelectCell,
  onApplySelectedCellChoice,
  onApplyCellsChoice,
  onResolveCell,
  onApplyRowChoice,
  onApplyColumnChoice,
  onDeleteRow,
  resolvedCellKeys,
  frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS,
  primaryKeyCol,
  sheetName,
  basePath,
  oursPath,
  theirsPath,
  mergedPath,
  fullBaseRows,
  fullOursRows,
  fullTheirsRows,
  mergedPreviewRows,
  mergedPreviewRowVisuals,
  mergedPreviewAlignedCols,
  onSaveMergeResult,
  saveMergeResultLabel,
  remainingCount,
  canUndo = false,
  onUndo,
  canJumpToPreviousConflict,
  onJumpToPreviousConflict,
  canJumpToNextConflict,
  onJumpToNextConflict,
  showTheirsChangedReviewFallback = false,
}) => {
  const baseScrollRef = useRef<HTMLDivElement | null>(null);
  const oursScrollRef = useRef<HTMLDivElement | null>(null);
  const theirsScrollRef = useRef<HTMLDivElement | null>(null);
  const mergedScrollRef = useRef<HTMLDivElement | null>(null);
  const leftPaneRef = useRef<HTMLDivElement | null>(null);
  const isSyncingXRef = useRef(false);
  const isSyncingYRef = useRef(false);
  const pendingSyncedScrollYRef = useRef<
    Partial<Record<SourceSide, { value: number; expiresAt: number }>>
  >({});
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [showDiffRowsOnly, setShowDiffRowsOnly] = useState(false);
  const [showConflictRowsOnly, setShowConflictRowsOnly] = useState(false);
  const [topPanelRatio, setTopPanelRatio] = useState(62);
  const isSimpleMergeMode = compareMode === 'simple-merge';

  useEffect(() => {
    if (!isSimpleMergeMode || !showDiffRowsOnly) return;
    setShowDiffRowsOnly(false);
  }, [isSimpleMergeMode, showDiffRowsOnly]);

  const getScrollRatio = (current: number, max: number): number => {
    if (!Number.isFinite(current) || !Number.isFinite(max) || max <= 0) return 0;
    return Math.min(1, Math.max(0, current / max));
  };

  const applyScrollByRatio = (element: HTMLDivElement, axis: 'x' | 'y', ratio: number) => {
    const maxScroll =
      axis === 'x'
        ? Math.max(0, element.scrollWidth - element.clientWidth)
        : Math.max(0, element.scrollHeight - element.clientHeight);
    const nextValue = maxScroll <= 0 ? 0 : ratio * maxScroll;
    if (axis === 'x') {
      element.scrollLeft = nextValue;
      return;
    }
    element.scrollTop = nextValue;
  };

  const shouldIgnoreSyncedScrollY = (side: SourceSide, scrollTop: number) => {
    const marker = pendingSyncedScrollYRef.current[side];
    if (!marker) return false;
    if (marker.expiresAt < performance.now()) {
      delete pendingSyncedScrollYRef.current[side];
      return false;
    }
    if (Math.abs(marker.value - scrollTop) <= SYNC_SCROLL_TOLERANCE_PX) {
      delete pendingSyncedScrollYRef.current[side];
      return true;
    }
    delete pendingSyncedScrollYRef.current[side];
    return false;
  };

  const setSyncedScrollTop = (side: SourceSide, element: HTMLDivElement, nextScrollTop: number) => {
    if (Math.abs(element.scrollTop - nextScrollTop) <= SYNC_SCROLL_TOLERANCE_PX) return;
    element.scrollTop = nextScrollTop;
    pendingSyncedScrollYRef.current[side] = {
      value: element.scrollTop,
      expiresAt: performance.now() + SYNC_SCROLL_EVENT_TTL_MS,
    };
  };

  const syncScrollX = (from: SourceSide, scrollLeft: number) => {
    const targets = [
      { side: 'base' as const, ref: baseScrollRef },
      { side: 'ours' as const, ref: oursScrollRef },
      { side: 'theirs' as const, ref: theirsScrollRef },
      { side: 'merged' as const, ref: mergedScrollRef },
    ];
    const sourceElement = targets.find((target) => target.side === from)?.ref.current ?? null;
    if (isSyncingXRef.current) return;
    isSyncingXRef.current = true;
    const sourceMaxScroll = sourceElement
      ? Math.max(0, sourceElement.scrollWidth - sourceElement.clientWidth)
      : 0;
    const ratio = getScrollRatio(scrollLeft, sourceMaxScroll);
    targets.forEach((target) => {
      if (target.side === from || !target.ref.current) return;
      applyScrollByRatio(target.ref.current, 'x', ratio);
    });
    requestAnimationFrame(() => {
      isSyncingXRef.current = false;
    });
  };

  const syncScrollY = (from: SourceSide, scrollTop: number) => {
    const targets = [
      { side: 'base' as const, ref: baseScrollRef },
      { side: 'ours' as const, ref: oursScrollRef },
      { side: 'theirs' as const, ref: theirsScrollRef },
      { side: 'merged' as const, ref: mergedScrollRef },
    ];
    const sourceElement = targets.find((target) => target.side === from)?.ref.current ?? null;
    if (shouldIgnoreSyncedScrollY(from, scrollTop)) return;
    if (isSyncingYRef.current) return;
    isSyncingYRef.current = true;
    const sourceMaxScroll = sourceElement
      ? Math.max(0, sourceElement.scrollHeight - sourceElement.clientHeight)
      : 0;
    const ratio = getScrollRatio(scrollTop, sourceMaxScroll);
    const safeFrozenRowCount = Math.max(0, Math.floor(frozenRowCount));
    const sourceRows = panelRowsBySide[from] ?? [];
    const sourceTopRowIndex = Math.min(
      Math.max(0, safeFrozenRowCount + Math.floor(scrollTop / GRID_ROW_HEIGHT)),
      Math.max(0, sourceRows.length - 1),
    );
    const sourceDisplayRowNumber = sourceRows[sourceTopRowIndex]?.[0]?.displayRowNumber ?? null;
    targets.forEach((target) => {
      if (target.side === from || !target.ref.current) return;
      const targetRowIndex =
        sourceDisplayRowNumber != null ? displayRowIndexBySide[target.side].get(sourceDisplayRowNumber) : undefined;
      if (typeof targetRowIndex === 'number') {
        setSyncedScrollTop(
          target.side,
          target.ref.current,
          targetRowIndex <= safeFrozenRowCount ? 0 : (targetRowIndex - safeFrozenRowCount) * GRID_ROW_HEIGHT,
        );
        return;
      }
      const beforeScrollTop = target.ref.current.scrollTop;
      applyScrollByRatio(target.ref.current, 'y', ratio);
      if (Math.abs(target.ref.current.scrollTop - beforeScrollTop) > SYNC_SCROLL_TOLERANCE_PX) {
        pendingSyncedScrollYRef.current[target.side] = {
          value: target.ref.current.scrollTop,
          expiresAt: performance.now() + SYNC_SCROLL_EVENT_TTL_MS,
        };
      }
    });
    requestAnimationFrame(() => {
      isSyncingYRef.current = false;
    });
  };

  const startResizePanels = (event: React.MouseEvent<HTMLDivElement>) => {
    event.preventDefault();
    const pane = leftPaneRef.current;
    if (!pane) return;
    const rect = pane.getBoundingClientRect();
    const startY = event.clientY;
    const startRatio = topPanelRatio;
    const minRatio = 30;
    const maxRatio = 80;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaY = moveEvent.clientY - startY;
      const deltaRatio = (deltaY / Math.max(rect.height, 1)) * 100;
      const nextRatio = Math.min(maxRatio, Math.max(minRatio, startRatio + deltaRatio));
      setTopPanelRatio(nextRatio);
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  const rowsMetaMap = useMemo(() => {
    const map = new Map<number, MergeRowMeta>();
    rowsMeta.forEach((row) => map.set(row.visualRowNumber, row));
    return map;
  }, [rowsMeta]);

  const sourceRowMap = useMemo(() => {
    const map = new Map<number, ThreeWayRowResult>();
    sourceRows.forEach((row) => {
      if (typeof row.rowNumber === 'number') {
        map.set(row.rowNumber, row);
      }
    });
    return map;
  }, [sourceRows]);

  const mergeCellMap = useMemo(() => {
    const map = new Map<string, MergeCell>();
    cells.forEach((cell) => map.set(`${cell.row}:${cell.col}`, cell));
    return map;
  }, [cells]);
  const columnsMetaMap = useMemo(() => {
    const map = new Map<number, MergeColumnMeta>();
    (columnsMeta ?? []).forEach((meta) => map.set(meta.col, meta));
    return map;
  }, [columnsMeta]);

  const resolvedKeySet = resolvedCellKeys ?? new Set<string>();
  const normalizedPrimaryKeyCol =
    typeof primaryKeyCol === 'number' && primaryKeyCol >= 1 ? Math.floor(primaryKeyCol) : null;
  const mergedPreviewAlignedColSet = useMemo(
    () => new Set((mergedPreviewAlignedCols ?? []).filter((value): value is number => typeof value === 'number' && value >= 1)),
    [mergedPreviewAlignedCols],
  );

  const diffColumns = useMemo(() => {
    const diffCols = new Set<number>();
    cells.forEach((cell) => diffCols.add(cell.col));
    if (normalizedPrimaryKeyCol) diffCols.add(normalizedPrimaryKeyCol);
    return Array.from(diffCols)
      .filter((col) => mergedPreviewAlignedColSet.size === 0 || mergedPreviewAlignedColSet.has(col))
      .sort((a, b) => a - b);
  }, [cells, mergedPreviewAlignedColSet, normalizedPrimaryKeyCol]);

  const allColumns = useMemo(() => {
    if ((mergedPreviewAlignedCols ?? []).length > 0) {
      return Array.from(new Set(mergedPreviewAlignedCols ?? [])).sort((a, b) => a - b);
    }
    if (columnsMeta && columnsMeta.length > 0) {
      return Array.from(new Set(columnsMeta.map((meta) => meta.col))).sort((a, b) => a - b);
    }
    const colCountFromSource = sourceRows[0]?.colCount ?? 0;
    if (colCountFromSource > 0) {
      return Array.from({ length: colCountFromSource }, (_, idx) => idx + 1);
    }
    return diffColumns;
  }, [columnsMeta, diffColumns, mergedPreviewAlignedCols, sourceRows]);

  const displayColumns = useMemo(() => {
    if (!normalizedPrimaryKeyCol || allColumns.includes(normalizedPrimaryKeyCol)) return allColumns;
    return [normalizedPrimaryKeyCol, ...allColumns].sort((a, b) => a - b);
  }, [allColumns, normalizedPrimaryKeyCol]);

  const allRowNumbers = useMemo(() => {
    if (rowsMeta.length > 0) {
      return rowsMeta.map((row) => row.visualRowNumber);
    }
    const rowNumbers = new Set<number>();
    cells.forEach((cell) => rowNumbers.add(cell.row));
    return Array.from(rowNumbers).sort((a, b) => a - b);
  }, [cells, rowsMeta]);

  const diffRowNumberSet = useMemo(() => {
    const rows = new Set<number>();
    cells.forEach((cell) => {
      if (cell.status !== 'unchanged') {
        rows.add(cell.row);
      }
    });
    return rows;
  }, [cells]);
  const conflictCells = useMemo(
    () => cells.filter((cell) => cell.status === 'conflict').sort((a, b) => a.row - b.row || a.col - b.col),
    [cells],
  );
  const conflictRowNumberSet = useMemo(() => {
    const rows = new Set<number>();
    conflictCells.forEach((cell) => {
      if (cell.status === 'conflict') {
        rows.add(cell.row);
      }
    });
    return rows;
  }, [conflictCells]);
  const effectiveShowDiffRowsOnly = !isSimpleMergeMode && showDiffRowsOnly && diffRowNumberSet.size > 0;
  const effectiveShowConflictRowsOnly = showConflictRowsOnly && conflictRowNumberSet.size > 0;
  const diffRowFilterFallback = !isSimpleMergeMode && showDiffRowsOnly && !effectiveShowDiffRowsOnly;
  const conflictRowFilterFallback = showConflictRowsOnly && !effectiveShowConflictRowsOnly;
  const displayRowNumbers = useMemo(() => {
    let rows = allRowNumbers;
    if (effectiveShowDiffRowsOnly) {
      rows = rows.filter((rowNumber) => diffRowNumberSet.has(rowNumber));
    }
    if (effectiveShowConflictRowsOnly) {
      rows = rows.filter((rowNumber) => conflictRowNumberSet.has(rowNumber));
    }
    return rows;
  }, [
    allRowNumbers,
    conflictRowNumberSet,
    diffRowNumberSet,
    effectiveShowConflictRowsOnly,
    effectiveShowDiffRowsOnly,
  ]);
  const displayRowNumberSet = useMemo(() => new Set(displayRowNumbers), [displayRowNumbers]);

  useEffect(() => {
    if (displayColumns.length === 0) return;
    setColumnWidths((prev) => {
      if (prev.length === displayColumns.length) return prev;
      return Array(displayColumns.length).fill(DEFAULT_COL_WIDTH);
    });
  }, [displayColumns.length]);

  const gridRows = useMemo<MergeWorkbenchCell[][]>(
    () =>
      displayRowNumbers.map((rowNumber) => {
        const sourceRow = sourceRowMap.get(rowNumber);
        return displayColumns.map((colNumber) => {
          const key = `${rowNumber}:${colNumber}`;
          const mergeCell = mergeCellMap.get(key);
          const baseValue = sourceRow?.base[colNumber - 1] ?? mergeCell?.baseValue ?? null;
          const oursValue = sourceRow?.ours[colNumber - 1] ?? mergeCell?.oursValue ?? null;
          const theirsValue = sourceRow?.theirs[colNumber - 1] ?? mergeCell?.theirsValue ?? null;
          return {
            key,
            address:
              mergeCell?.address ??
              makeCellAddress(
                colNumber,
                sourceRow?.oursRowNumber ?? sourceRow?.baseRowNumber ?? sourceRow?.theirsRowNumber ?? rowNumber,
              ),
            rowNumber,
            displayRowNumber: rowNumber,
            colNumber,
            baseValue,
            oursValue,
            theirsValue,
            mergedValue: getDefaultMergedValue(mergeCell, sourceRow, colNumber, mergeCell ? resolvedKeySet.has(key) : true),
            status: mergeCell?.status ?? 'unchanged',
            resolved: mergeCell ? resolvedKeySet.has(key) : true,
            isDiffCell: !!mergeCell && mergeCell.status !== 'unchanged',
            isContextCell: !!mergeCell && mergeCell.status === 'unchanged',
            formulaControlled: mergeCell?.formulaControlled === true,
            sharedControlled: mergeCell?.sharedControlled === true,
            sharedControlMasterSheetName: mergeCell?.sharedControlMasterSheetName ?? null,
            sharedControlIsMaster: mergeCell?.sharedControlIsMaster === true,
          };
        });
      }),
    [displayColumns, displayRowNumbers, mergeCellMap, resolvedKeySet, sourceRowMap],
  );

  const attentionCells = useMemo(
    () =>
      conflictCells
        .filter((cell) => isAttentionBucketCell(cell, resolvedKeySet, showTheirsChangedReviewFallback))
        .sort((a, b) => a.row - b.row || a.col - b.col),
    [conflictCells, resolvedKeySet, showTheirsChangedReviewFallback],
  );
  const totalConflictCount = conflictCells.length;
  const unresolvedConflictCount = remainingCount;
  const attentionKeySet = useMemo(
    () => new Set(attentionCells.map((cell) => `${cell.row}:${cell.col}`)),
    [attentionCells],
  );
  const attentionListTitle = '冲突列表';
  const attentionListHint = '可滚动，点击跳转';
  const navButtonLabelPrefix = '冲突';
  const diffSummary = useMemo(() => {
    let totalDiff = 0;
    let oursChanged = 0;
    let theirsChanged = 0;
    let bothChangedSame = 0;
    cells.forEach((cell) => {
      if (cell.status === 'unchanged') return;
      totalDiff += 1;
      if (cell.status === 'ours-changed') {
        oursChanged += 1;
        return;
      }
      if (cell.status === 'theirs-changed') {
        theirsChanged += 1;
        return;
      }
      if (cell.status === 'both-changed-same') {
        bothChangedSame += 1;
        return;
      }
    });

    return {
      totalDiff,
      oursChanged,
      theirsChanged,
      bothChangedSame,
      totalConflict: totalConflictCount,
      unresolvedConflict: unresolvedConflictCount,
      autoMerged: totalDiff - totalConflictCount,
    };
  }, [cells, totalConflictCount, unresolvedConflictCount]);
  const conflictCountByRow = useMemo(() => {
    const map = new Map<number, number>();
    attentionCells.forEach((cell) => {
      map.set(cell.row, (map.get(cell.row) ?? 0) + 1);
    });
    return map;
  }, [attentionCells]);
  const usePhysicalGrid =
    Array.isArray(fullBaseRows) &&
    Array.isArray(fullOursRows) &&
    Array.isArray(fullTheirsRows) &&
    Array.isArray(mergedPreviewRows) &&
    Array.isArray(mergedPreviewRowVisuals) &&
    mergedPreviewRows.length > 0 &&
    mergedPreviewRowVisuals.length === mergedPreviewRows.length;
  const physicalRowMapBySide = useMemo(() => {
    const base = new Map<number, number>();
    const ours = new Map<number, number>();
    const theirs = new Map<number, number>();
    rowsMeta.forEach((meta) => {
      if (meta.baseRowNumber) base.set(meta.baseRowNumber, meta.visualRowNumber);
      if (meta.oursRowNumber) ours.set(meta.oursRowNumber, meta.visualRowNumber);
      if (meta.theirsRowNumber) theirs.set(meta.theirsRowNumber, meta.visualRowNumber);
    });
    return { base, ours, theirs };
  }, [rowsMeta]);
  const mergedPreviewColIndexMap = useMemo(() => {
    const map = new Map<number, number>();
    (mergedPreviewAlignedCols ?? []).forEach((alignedCol, idx) => {
      if (typeof alignedCol === 'number' && alignedCol >= 1) {
        map.set(alignedCol, idx);
      }
    });
    return map;
  }, [mergedPreviewAlignedCols]);
  const buildPhysicalSourceGridRows = useMemo(() => {
    if (!usePhysicalGrid) {
      return {
        base: [] as MergeWorkbenchCell[][],
        ours: [] as MergeWorkbenchCell[][],
        theirs: [] as MergeWorkbenchCell[][],
      };
    }
    const buildRows = (
      side: 'base' | 'ours' | 'theirs',
      rawRows: (string | number | null)[][],
      visualRowByPhysical: Map<number, number>,
    ) => {
      return rawRows.reduce<MergeWorkbenchCell[][]>((acc, rawRow, rowIndex) => {
        const physicalRowNumber = rowIndex + 1;
        const visualRowNumber = visualRowByPhysical.get(physicalRowNumber) ?? 0;
        if ((effectiveShowDiffRowsOnly || effectiveShowConflictRowsOnly) && !displayRowNumberSet.has(visualRowNumber)) {
          return acc;
        }
        const rowCells = displayColumns.map((alignedCol) => {
          const meta = columnsMetaMap.get(alignedCol);
          const physicalColNumber =
            side === 'base'
              ? meta?.baseCol ?? alignedCol
              : side === 'ours'
                ? meta?.oursCol ?? alignedCol
                : meta?.theirsCol ?? alignedCol;
          const rawValue =
            physicalColNumber && physicalColNumber >= 1 ? rawRow?.[physicalColNumber - 1] ?? null : null;
          const mergeCell = visualRowNumber > 0 ? mergeCellMap.get(`${visualRowNumber}:${alignedCol}`) : undefined;
          const sourceRow = visualRowNumber > 0 ? sourceRowMap.get(visualRowNumber) : undefined;
          const key =
            visualRowNumber > 0
              ? `${visualRowNumber}:${alignedCol}`
              : `${side}:${physicalRowNumber}:${alignedCol}`;
          return {
            key,
            address: makeCellAddress(physicalColNumber ?? alignedCol, physicalRowNumber),
            rowNumber: visualRowNumber,
            displayRowNumber: physicalRowNumber,
            colNumber: alignedCol,
            baseValue: side === 'base' ? rawValue : sourceRow?.base[alignedCol - 1] ?? mergeCell?.baseValue ?? null,
            oursValue: side === 'ours' ? rawValue : sourceRow?.ours[alignedCol - 1] ?? mergeCell?.oursValue ?? null,
            theirsValue:
              side === 'theirs' ? rawValue : sourceRow?.theirs[alignedCol - 1] ?? mergeCell?.theirsValue ?? null,
            mergedValue: getDefaultMergedValue(
              mergeCell,
              sourceRow,
              alignedCol,
              mergeCell ? resolvedKeySet.has(key) : true,
            ),
            status: mergeCell?.status ?? 'unchanged',
            resolved: mergeCell ? resolvedKeySet.has(key) : true,
            isDiffCell: !!mergeCell && mergeCell.status !== 'unchanged',
            isContextCell: !!mergeCell && mergeCell.status === 'unchanged',
            formulaControlled: mergeCell?.formulaControlled === true,
            sharedControlled: mergeCell?.sharedControlled === true,
            sharedControlMasterSheetName: mergeCell?.sharedControlMasterSheetName ?? null,
            sharedControlIsMaster: mergeCell?.sharedControlIsMaster === true,
          };
        });
        acc.push(rowCells);
        return acc;
      }, []);
    };
    return {
      base: buildRows('base', fullBaseRows ?? [], physicalRowMapBySide.base),
      ours: buildRows('ours', fullOursRows ?? [], physicalRowMapBySide.ours),
      theirs: buildRows('theirs', fullTheirsRows ?? [], physicalRowMapBySide.theirs),
    };
  }, [
    columnsMetaMap,
    displayColumns,
    fullBaseRows,
    fullOursRows,
    fullTheirsRows,
    mergeCellMap,
    physicalRowMapBySide,
    displayRowNumberSet,
    resolvedKeySet,
    sourceRowMap,
    effectiveShowConflictRowsOnly,
    effectiveShowDiffRowsOnly,
    usePhysicalGrid,
  ]);
  const mergedGridRows = useMemo<MergeWorkbenchCell[][]>(() => {
    if (!usePhysicalGrid) return [];
    return (mergedPreviewRows ?? []).reduce<MergeWorkbenchCell[][]>((acc, previewRow, rowIndex) => {
      const visualRowNumber = mergedPreviewRowVisuals?.[rowIndex] ?? 0;
      if ((effectiveShowDiffRowsOnly || effectiveShowConflictRowsOnly) && !displayRowNumberSet.has(visualRowNumber)) {
        return acc;
      }
      const physicalRowNumber = rowIndex + 1;
      const rowCells = displayColumns.map((alignedCol) => {
        const colIndex = mergedPreviewColIndexMap.get(alignedCol) ?? -1;
        const mergeCell = visualRowNumber > 0 ? mergeCellMap.get(`${visualRowNumber}:${alignedCol}`) : undefined;
        const sourceRow = visualRowNumber > 0 ? sourceRowMap.get(visualRowNumber) : undefined;
        const key =
          visualRowNumber > 0
            ? `${visualRowNumber}:${alignedCol}`
            : `merged:${physicalRowNumber}:${alignedCol}`;
        return {
          key,
          address: mergeCell?.address ?? makeCellAddress(alignedCol, physicalRowNumber),
          rowNumber: visualRowNumber,
          displayRowNumber: physicalRowNumber,
          colNumber: alignedCol,
          baseValue: sourceRow?.base[alignedCol - 1] ?? mergeCell?.baseValue ?? null,
          oursValue: sourceRow?.ours[alignedCol - 1] ?? mergeCell?.oursValue ?? null,
          theirsValue: sourceRow?.theirs[alignedCol - 1] ?? mergeCell?.theirsValue ?? null,
          mergedValue: colIndex >= 0 ? previewRow?.[colIndex] ?? null : null,
          status: mergeCell?.status ?? 'unchanged',
          resolved: mergeCell ? resolvedKeySet.has(key) : true,
          isDiffCell: !!mergeCell && mergeCell.status !== 'unchanged',
          isContextCell: !!mergeCell && mergeCell.status === 'unchanged',
          formulaControlled: mergeCell?.formulaControlled === true,
          sharedControlled: mergeCell?.sharedControlled === true,
          sharedControlMasterSheetName: mergeCell?.sharedControlMasterSheetName ?? null,
          sharedControlIsMaster: mergeCell?.sharedControlIsMaster === true,
        };
      });
      acc.push(rowCells);
      return acc;
    }, []);
  }, [
    displayColumns,
    mergeCellMap,
    mergedPreviewColIndexMap,
    mergedPreviewRowVisuals,
    mergedPreviewRows,
    displayRowNumberSet,
    resolvedKeySet,
    sourceRowMap,
    effectiveShowConflictRowsOnly,
    effectiveShowDiffRowsOnly,
    usePhysicalGrid,
  ]);
  const panelRowsBySide = useMemo(
    () => ({
      base: usePhysicalGrid ? buildPhysicalSourceGridRows.base : gridRows,
      ours: usePhysicalGrid ? buildPhysicalSourceGridRows.ours : gridRows,
      theirs: usePhysicalGrid ? buildPhysicalSourceGridRows.theirs : gridRows,
      merged: usePhysicalGrid ? mergedGridRows : gridRows,
    }),
    [buildPhysicalSourceGridRows, gridRows, mergedGridRows, usePhysicalGrid],
  );
  const displayRowIndexBySide = useMemo(() => {
    const buildIndexMap = (rows: MergeWorkbenchCell[][]) => {
      const map = new Map<number, number>();
      rows.forEach((row, rowIndex) => {
        const displayRowNumber = row[0]?.displayRowNumber;
        if (typeof displayRowNumber !== 'number' || map.has(displayRowNumber)) return;
        map.set(displayRowNumber, rowIndex);
      });
      return map;
    };
    return {
      base: buildIndexMap(panelRowsBySide.base),
      ours: buildIndexMap(panelRowsBySide.ours),
      theirs: buildIndexMap(panelRowsBySide.theirs),
      merged: buildIndexMap(panelRowsBySide.merged),
    };
  }, [panelRowsBySide]);

  const selectedGridRowIndex = selected ? displayRowNumbers.indexOf(selected.rowIndex + 1) : -1;
  const selectedGridColIndex = selected ? displayColumns.indexOf(selected.colIndex + 1) : -1;
  const selectedCell =
    selectedGridRowIndex >= 0 && selectedGridColIndex >= 0
      ? gridRows[selectedGridRowIndex]?.[selectedGridColIndex] ?? null
      : null;
  const selectedSourceRowIndexBySide = useMemo(() => {
    const alignedRowNumber = selected ? selected.rowIndex + 1 : 0;
    const findFilteredRowIndex = (rows: MergeWorkbenchCell[][]): number =>
      rows.findIndex((row) => (row[0]?.rowNumber ?? 0) === alignedRowNumber);
    return {
      base: alignedRowNumber > 0 ? findFilteredRowIndex(buildPhysicalSourceGridRows.base) : -1,
      ours: alignedRowNumber > 0 ? findFilteredRowIndex(buildPhysicalSourceGridRows.ours) : -1,
      theirs: alignedRowNumber > 0 ? findFilteredRowIndex(buildPhysicalSourceGridRows.theirs) : -1,
      merged: alignedRowNumber > 0 ? findFilteredRowIndex(mergedGridRows) : -1,
    };
  }, [buildPhysicalSourceGridRows, mergedGridRows, selected]);
  const selectedRowMeta = selectedCell ? rowsMetaMap.get(selectedCell.rowNumber) : undefined;
  const selectedIsDiffCell = selectedCell?.isDiffCell === true;
  const selectedCanApplyCellChoice =
    !!selectedCell && selectedIsDiffCell && !isProtectedCell(selectedCell);
  const selectedRowHasDiff = selectedCell
    ? cells.some((cell) => cell.row === selectedCell.rowNumber && cell.status !== 'unchanged')
    : false;
  const selectedColumnHasDiff = selectedCell
    ? cells.some((cell) => cell.col === selectedCell.colNumber && cell.status !== 'unchanged')
    : false;
  const selectedRowAttentionCount = selectedCell ? conflictCountByRow.get(selectedCell.rowNumber) ?? 0 : 0;

  const jumpToNextConflict = () => {
    if (!onSelectCell || attentionCells.length === 0) return;
    const currentIndex = selectedCell
      ? attentionCells.findIndex(
          (cell) => cell.row === selectedCell.rowNumber && cell.col === selectedCell.colNumber,
        )
      : -1;
    const nextCell = attentionCells[currentIndex >= 0 ? (currentIndex + 1) % attentionCells.length : 0];
    if (!nextCell) return;
    onSelectCell(nextCell.row - 1, nextCell.col - 1);
  };
  const jumpToPreviousConflict = () => {
    if (!onSelectCell || attentionCells.length === 0) return;
    const currentIndex = selectedCell
      ? attentionCells.findIndex(
          (cell) => cell.row === selectedCell.rowNumber && cell.col === selectedCell.colNumber,
        )
      : -1;
    const previousCell =
      attentionCells[
        currentIndex >= 0
          ? (currentIndex - 1 + attentionCells.length) % attentionCells.length
          : attentionCells.length - 1
      ];
    if (!previousCell) return;
    onSelectCell(previousCell.row - 1, previousCell.col - 1);
  };
  const effectiveCanJumpToPreviousConflict =
    canJumpToPreviousConflict ?? (attentionCells.length > 0);
  const effectiveJumpToPreviousConflict =
    onJumpToPreviousConflict ?? jumpToPreviousConflict;
  const effectiveCanJumpToNextConflict = canJumpToNextConflict ?? (attentionCells.length > 0);
  const effectiveJumpToNextConflict = onJumpToNextConflict ?? jumpToNextConflict;

  const applyAllConflictsChoice = (source: 'base' | 'ours' | 'theirs') => {
    if (!onApplyCellsChoice || attentionCells.length === 0) return;
    onApplyCellsChoice(
      attentionCells.map((cell) => ({ rowNumber: cell.row, colNumber: cell.col })),
      source,
    );
  };

  const makeAlignedRowHeader = (rowIndex: number) => {
    const rowNumber = displayRowNumbers[rowIndex];
    return (
      <div title={`row ${rowNumber}`} style={{ display: 'flex', justifyContent: 'flex-end', fontWeight: 700 }}>
        {rowNumber}
      </div>
    );
  };
  const makePhysicalRowHeader = (rows: MergeWorkbenchCell[][], rowIndex: number) => {
    const rowNumber = rows[rowIndex]?.[0]?.displayRowNumber ?? rowIndex + 1;
    return (
      <div title={`row ${rowNumber}`} style={{ display: 'flex', justifyContent: 'flex-end', fontWeight: 700 }}>
        {rowNumber}
      </div>
    );
  };

  const renderHeaderCell = (gridColIndex: number) => {
    const alignedCol = displayColumns[gridColIndex];
    const meta = columnsMeta?.find((column) => column.col === alignedCol);
    const title = meta
      ? isSimpleMergeMode
        ? `aligned ${colNumberToLabel(alignedCol)} | ours=${meta.oursCol ? colNumberToLabel(meta.oursCol) : '-'} | theirs=${meta.theirsCol ? colNumberToLabel(meta.theirsCol) : '-'}`
        : `aligned ${colNumberToLabel(alignedCol)} | base=${meta.baseCol ? colNumberToLabel(meta.baseCol) : '-'} | ours=${meta.oursCol ? colNumberToLabel(meta.oursCol) : '-'} | theirs=${meta.theirsCol ? colNumberToLabel(meta.theirsCol) : '-'}`
      : colNumberToLabel(alignedCol);
    return (
      <span title={title} style={{ whiteSpace: 'nowrap' }}>
        {colNumberToLabel(alignedCol)}
      </span>
    );
  };

  const makeRenderCell =
    (side: SourceSide) =>
    (cell: MergeWorkbenchCell | null, _ctx: VirtualGridRenderCtx) => {
      if (!cell) return null;
      const value =
        side === 'base'
          ? cell.baseValue
          : side === 'ours'
            ? cell.oursValue
            : side === 'theirs'
              ? cell.theirsValue
              : cell.mergedValue;
      return (
        <div
          onMouseDown={() => {
            if (cell.rowNumber <= 0) return;
            onSelectCell?.(cell.rowNumber - 1, cell.colNumber - 1);
          }}
          onClick={() => {
            if (cell.rowNumber <= 0) return;
            onSelectCell?.(cell.rowNumber - 1, cell.colNumber - 1);
          }}
          title={[
            `地址: ${cell.address}`,
            `${side}: ${makeValueText(value)}`,
            `状态: ${getDisplayedCellStatusLabel(cell)}`,
            getProtectedCellHint(cell),
          ]
            .filter(Boolean)
            .join('\n')}
          style={{
            width: '100%',
            height: '100%',
            boxSizing: 'border-box',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            cursor: 'pointer',
            userSelect: 'none',
            fontWeight: side === 'merged' && cell.isDiffCell ? 600 : 400,
            color: isProtectedCell(cell) ? FORMULA_TEXT : '#111827',
          }}
        >
          {value == null ? '' : String(value)}
        </div>
      );
    };

  const makeGetCellStyle =
    (side: SourceSide) =>
    (cell: MergeWorkbenchCell | null, ctx: VirtualGridRenderCtx): React.CSSProperties => {
      const style: React.CSSProperties = {};
      if (ctx.isFrozenRow) {
        style.backgroundColor = FROZEN_BG;
      }
      if (cell) {
        const isPendingReview = attentionKeySet.has(cell.key);
        const isAutoMerged = cell.isDiffCell && cell.status !== 'unchanged' && !isPendingReview;
        if (isProtectedCell(cell)) {
          style.backgroundColor = FORMULA_BG;
          style.borderLeft = `3px solid ${FORMULA_BORDER}`;
          style.boxShadow = `inset 0 0 0 1px ${FORMULA_BORDER}`;
        } else {
          style.backgroundColor = getPanelBackground(
            cell.status,
            side,
            cell.resolved,
            isPendingReview,
            isAutoMerged,
          );
          if (cell.status === 'conflict' && !cell.resolved) {
            style.boxShadow = `inset 0 0 0 2px ${CONFLICT_OUTLINE}`;
          } else if (side === 'merged' && isPendingReview) {
            style.borderLeft = `3px solid ${PENDING_REVIEW_OUTLINE}`;
            style.boxShadow = `inset 0 0 0 1px ${PENDING_REVIEW_OUTLINE}`;
          } else if (cell.status !== 'unchanged' && !cell.resolved) {
            style.borderLeft = `3px solid ${getPanelAccent(side)}`;
          } else if (side === 'merged' && isAutoMerged) {
            style.borderLeft = `3px solid ${getPanelAccent(side)}`;
            style.boxShadow = `inset 0 0 0 2px ${getPanelAccent(side)}`;
          }
        }
        if (selected && selected.rowIndex === cell.rowNumber - 1 && selected.colIndex === cell.colNumber - 1) {
          style.outline = '2px solid #2563eb';
          style.outlineOffset = '-2px';
          style.position = 'relative';
          style.zIndex = 6;
        }
      }
      return style;
    };

  const panelSpecs: Array<{
    side: 'base' | 'ours' | 'theirs';
    title: string;
    path: string | null | undefined;
    ref: React.RefObject<HTMLDivElement>;
    showRowHeader: boolean;
  }> = isSimpleMergeMode
    ? [
        { side: 'ours', title: 'ours（只读）', path: oursPath, ref: oursScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: true },
        { side: 'theirs', title: 'theirs（只读）', path: theirsPath, ref: theirsScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: true },
      ]
    : [
        { side: 'base', title: 'base（只读）', path: basePath, ref: baseScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: true },
        { side: 'ours', title: 'ours（只读）', path: oursPath, ref: oursScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: false },
        { side: 'theirs', title: 'theirs（只读）', path: theirsPath, ref: theirsScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: false },
      ];

  const statusTone =
    isProtectedCell(selectedCell)
      ? FORMULA_TEXT
      : selectedCell?.status === 'conflict'
        ? '#c2410c'
        : selectedCell?.status === 'ours-changed' ||
            selectedCell?.status === 'theirs-changed' ||
            selectedCell?.status === 'both-changed-same'
          ? '#1d4ed8'
          : '#475569';
  const showGridPane = layoutMode !== 'panel-only';
  const showDecisionPane = layoutMode !== 'grids-only';
  const showToolbar = layoutMode !== 'panel-only';
  const hasRenderedRows = usePhysicalGrid
    ? buildPhysicalSourceGridRows.base.length > 0 ||
      buildPhysicalSourceGridRows.ours.length > 0 ||
      buildPhysicalSourceGridRows.theirs.length > 0 ||
      mergedGridRows.length > 0
    : displayRowNumbers.length > 0;

  return displayColumns.length === 0 || !hasRenderedRows ? (
    <div style={{ border: '1px solid #d0d7de', borderRadius: 12, padding: 16 }}>没有可展示的 merge 差异。</div>
  ) : (
    <div
      style={{
        border: '1px solid #d0d7de',
        borderRadius: 12,
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        height: '100%',
        minHeight: 0,
        backgroundColor: '#fff',
      }}
    >
      {showToolbar && (
        <div
          style={{
            padding: '10px 12px',
            borderBottom: '1px solid #e5e7eb',
            display: 'flex',
            alignItems: 'center',
            gap: 10,
            flexWrap: 'wrap',
            backgroundColor: '#fafbfc',
          }}
        >
          <span style={{ fontWeight: 700, color: '#111827' }}>Merge 工作台{sheetName ? ` · ${sheetName}` : ''}</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>
            {isSimpleMergeMode ? '两侧源数据只读；所有选择都通过右侧决策栏完成。' : '三侧源数据只读；所有选择都通过右侧决策栏完成。'}
          </span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>差异单元格: {diffSummary.totalDiff}</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>自动并入: {diffSummary.autoMerged}</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>剩余冲突: {unresolvedConflictCount}</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>显示行: {usePhysicalGrid ? mergedGridRows.length : displayRowNumbers.length}</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>显示列: {displayColumns.length}</span>
          {!isSimpleMergeMode && (
            <label style={{ fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 4, color: '#334155' }}>
              <input
                type="checkbox"
                checked={showDiffRowsOnly}
                onChange={(event) => setShowDiffRowsOnly(event.target.checked)}
              />
              仅差异行
            </label>
          )}
          <label style={{ fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 4, color: '#334155' }}>
            <input
              type="checkbox"
              checked={showConflictRowsOnly}
              onChange={(event) => setShowConflictRowsOnly(event.target.checked)}
            />
            仅冲突行
          </label>
          {diffRowFilterFallback && (
            <span style={{ fontSize: 12, color: '#92400e' }}>当前工作表没有差异行，已保留全部行显示。</span>
          )}
          {conflictRowFilterFallback && (
            <span style={{ fontSize: 12, color: '#92400e' }}>当前工作表没有冲突行，已保留全部行显示。</span>
          )}
          <button type="button" onClick={effectiveJumpToPreviousConflict} disabled={!effectiveCanJumpToPreviousConflict}>
            上一个{navButtonLabelPrefix}
          </button>
          <button type="button" onClick={effectiveJumpToNextConflict} disabled={!effectiveCanJumpToNextConflict}>
            下一个{navButtonLabelPrefix}
          </button>
          <button type="button" onClick={onUndo} disabled={!canUndo}>
            撤销上一步
          </button>
          {onSaveMergeResult && (
            <button type="button" onClick={onSaveMergeResult}>
              {saveMergeResultLabel ?? '保存合并结果'}
            </button>
          )}
          {mergedPath && (
            <span
              title={mergedPath}
              style={{
                marginLeft: 'auto',
                maxWidth: 320,
                fontSize: 11,
                color: '#6b7280',
                whiteSpace: 'nowrap',
                overflow: 'hidden',
                textOverflow: 'ellipsis',
              }}
            >
              merged: {mergedPath}
            </span>
          )}
        </div>
      )}
      <div style={{ display: 'flex', flex: 1, minHeight: 0 }}>
        {showGridPane && (
          <div
            ref={leftPaneRef}
            style={{
              flex: 1,
              minWidth: 0,
              minHeight: 0,
              display: 'flex',
              flexDirection: 'column',
              gap: 0,
              padding: 8,
              overflow: 'hidden',
              position: 'relative',
              zIndex: 1,
            }}
          >
          <div
            style={{
              flex: `0 0 ${topPanelRatio}%`,
              minHeight: 200,
              display: 'grid',
              gridTemplateColumns: `repeat(${panelSpecs.length}, minmax(0, 1fr))`,
              gap: 8,
              minWidth: 0,
              overflow: 'hidden',
            }}
          >
            {panelSpecs.map((panel) => {
              const panelRows = panelRowsBySide[panel.side];
              const scrollRowIndex = usePhysicalGrid
                ? selectedSourceRowIndexBySide[panel.side]
                : selectedGridRowIndex;
              return (
                <div
                  key={panel.side}
                  style={{
                    border: '1px solid #e2e8f0',
                    borderRadius: 10,
                    display: 'flex',
                    flexDirection: 'column',
                    minWidth: 0,
                    minHeight: 0,
                    backgroundColor: '#fff',
                    overflow: 'hidden',
                  }}
                >
                  <div
                    style={{
                      padding: '8px 10px',
                      borderBottom: '1px solid #eef1f4',
                      backgroundColor: '#fcfcfd',
                      display: 'grid',
                      gap: 2,
                    }}
                  >
                    <span style={{ fontSize: 12, fontWeight: 700, color: getPanelAccent(panel.side) }}>{panel.title}</span>
                    <span
                      title={panel.path ?? ''}
                      style={{ fontSize: 11, color: '#6b7280', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}
                    >
                      {truncatePath(panel.path)}
                    </span>
                  </div>
                  <div style={{ flex: 1, minHeight: 0 }}>
                    <VirtualGrid<MergeWorkbenchCell>
                      rows={panelRows}
                      frozenRowCount={Math.max(0, Math.floor(frozenRowCount))}
                      frozenColCount={0}
                      rowHeaderWidth={panel.showRowHeader ? ROW_HEADER_WIDTH : 0}
                      showRowHeader={panel.showRowHeader}
                      renderRowHeader={(rowIndex) =>
                        usePhysicalGrid ? makePhysicalRowHeader(panelRows, rowIndex) : makeAlignedRowHeader(rowIndex)
                      }
                      renderCell={makeRenderCell(panel.side)}
                      getCellStyle={makeGetCellStyle(panel.side)}
                      renderHeaderCell={renderHeaderCell}
                      defaultColWidth={DEFAULT_COL_WIDTH}
                      columnWidths={columnWidths}
                      onColumnWidthsChange={setColumnWidths}
                      forceScrollbars
                      containerRef={panel.ref}
                      onScrollXChange={(left) => syncScrollX(panel.side, left)}
                      onScrollYChange={(top) => syncScrollY(panel.side, top)}
                      scrollToCell={
                        scrollRowIndex >= 0 && selectedGridColIndex >= 0
                          ? { rowIndex: scrollRowIndex, colIndex: selectedGridColIndex }
                          : null
                      }
                      scrollToCellAlign="nearest"
                    />
                  </div>
                </div>
              );
            })}
          </div>
          <div
            onMouseDown={startResizePanels}
            style={{
              height: 10,
              margin: '4px 0',
              cursor: 'row-resize',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              userSelect: 'none',
              flexShrink: 0,
            }}
            title={isSimpleMergeMode ? '拖动调整上方两栏和下方 merged 面板高度' : '拖动调整上方三栏和下方 merged 面板高度'}
          >
            <div
              style={{
                width: 120,
                height: 3,
                borderRadius: 999,
                backgroundColor: '#cbd5e1',
              }}
            />
          </div>
          <div
            style={{
              flex: 1,
              minHeight: 150,
              border: '1px solid #e2e8f0',
              borderRadius: 10,
              display: 'flex',
              flexDirection: 'column',
              backgroundColor: '#fff',
              overflow: 'hidden',
            }}
          >
            <div
              style={{
                padding: '8px 10px',
                borderBottom: '1px solid #eef1f4',
                backgroundColor: '#fcfcfd',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: 8,
              }}
            >
              <span style={{ fontSize: 12, fontWeight: 700, color: getPanelAccent('merged') }}>merged（结果）</span>
              <span style={{ fontSize: 11, color: '#6b7280' }}>
                {isSimpleMergeMode ? '橙色=冲突 · merged 默认保留 ours' : '橙色=冲突 · 蓝色=自动并入'}
              </span>
            </div>
            <div style={{ flex: 1, minHeight: 0 }}>
              <VirtualGrid<MergeWorkbenchCell>
                rows={usePhysicalGrid ? mergedGridRows : gridRows}
                frozenRowCount={Math.max(0, Math.floor(frozenRowCount))}
                frozenColCount={0}
                rowHeaderWidth={ROW_HEADER_WIDTH}
                showRowHeader
                renderRowHeader={(rowIndex) =>
                  usePhysicalGrid ? makePhysicalRowHeader(mergedGridRows, rowIndex) : makeAlignedRowHeader(rowIndex)
                }
                renderCell={makeRenderCell('merged')}
                getCellStyle={makeGetCellStyle('merged')}
                renderHeaderCell={renderHeaderCell}
                defaultColWidth={DEFAULT_COL_WIDTH}
                columnWidths={columnWidths}
                onColumnWidthsChange={setColumnWidths}
                forceScrollbars
                containerRef={mergedScrollRef as React.RefObject<HTMLDivElement>}
                onScrollXChange={(left) => syncScrollX('merged', left)}
                onScrollYChange={(top) => syncScrollY('merged', top)}
                scrollToCell={
                  !isSimpleMergeMode &&
                  (usePhysicalGrid ? selectedSourceRowIndexBySide.merged : selectedGridRowIndex) >= 0 &&
                  selectedGridColIndex >= 0
                    ? {
                        rowIndex: usePhysicalGrid ? selectedSourceRowIndexBySide.merged : selectedGridRowIndex,
                        colIndex: selectedGridColIndex,
                      }
                    : null
                }
                scrollToCellAlign="nearest"
              />
            </div>
          </div>
          </div>
        )}
        {showDecisionPane && (
          <div
            style={{
              width: layoutMode === 'panel-only' ? '100%' : 460,
              minWidth: layoutMode === 'panel-only' ? 0 : 420,
              flex: layoutMode === 'panel-only' ? 1 : undefined,
              borderLeft: layoutMode === 'full' ? '1px solid #e5e7eb' : 'none',
              backgroundColor: '#fbfcfe',
              display: 'flex',
              flexDirection: 'column',
              minHeight: 0,
              overflow: 'hidden',
              position: 'relative',
              zIndex: 40,
              isolation: 'isolate',
            }}
          >
          <div style={{ padding: 12, borderBottom: '1px solid #e5e7eb', display: 'grid', gap: 8, backgroundColor: '#fbfcfe', flexShrink: 0 }}>
            <div style={{ fontWeight: 700, color: '#111827' }}>决策与解释</div>
            <div style={{ fontSize: 12, color: '#4b5563', lineHeight: 1.5, wordBreak: 'break-word' }}>
              左侧列表可滚动定位，右侧固定展示当前选中项解释和快速处理。
            </div>
          </div>
          <div style={{ padding: 12, display: 'grid', gap: 12, borderBottom: '1px solid #e5e7eb', backgroundColor: '#fbfcfe', flexShrink: 0 }}>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: 8 }}>
              <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, padding: 10, backgroundColor: '#fff' }}>
                <div style={{ fontSize: 11, color: '#6b7280' }}>差异单元格</div>
                <div style={{ marginTop: 4, fontSize: 20, fontWeight: 700 }}>{diffSummary.totalDiff}</div>
              </div>
              <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, padding: 10, backgroundColor: '#fff' }}>
                <div style={{ fontSize: 11, color: '#6b7280' }}>自动并入</div>
                <div style={{ marginTop: 4, fontSize: 20, fontWeight: 700, color: '#1d4ed8' }}>{diffSummary.autoMerged}</div>
              </div>
              <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, padding: 10, backgroundColor: '#fff' }}>
                <div style={{ fontSize: 11, color: '#6b7280' }}>冲突</div>
                <div
                  style={{
                    marginTop: 4,
                    fontSize: 20,
                    fontWeight: 700,
                    color: '#c2410c',
                  }}
                >
                  {unresolvedConflictCount}
                </div>
              </div>
            </div>
          </div>
          <div
            style={{
              padding: 12,
              display: 'grid',
              gap: 12,
              gridTemplateColumns: '170px minmax(0, 1fr)',
              minHeight: 0,
              overflow: 'hidden',
              position: 'relative',
              zIndex: 1,
              backgroundColor: '#fbfcfe',
            }}
          >
            <div style={{ border: '1px solid #dbe3ef', borderRadius: 12, backgroundColor: '#fff', display: 'flex', flexDirection: 'column', minHeight: 0, overflow: 'hidden' }}>
              <div style={{ padding: '10px 10px 8px', borderBottom: '1px solid #edf0f5', display: 'grid', gap: 2 }}>
                <div style={{ fontWeight: 700, color: '#1f2937' }}>{attentionListTitle}</div>
                <div style={{ fontSize: 11, color: '#6b7280' }}>{attentionListHint}</div>
              </div>
              <div style={{ flex: 1, minHeight: 0, overflowY: 'scroll', padding: 8, display: 'grid', gap: 6 }}>
                {attentionCells.length === 0 ? (
                  <div style={{ fontSize: 12, color: '#6b7280', padding: 6 }}>
                    {diffSummary.totalDiff > 0
                      ? isSimpleMergeMode
                        ? unresolvedConflictCount > 0
                          ? `当前工作表共有 ${diffSummary.totalDiff} 个差异单元格；剩余 ${unresolvedConflictCount} 个冲突等待你确认。`
                          : `当前工作表共有 ${diffSummary.totalDiff} 个差异单元格；冲突已全部处理。`
                        : `当前工作表没有冲突；共有 ${diffSummary.totalDiff} 个差异单元格，其中 ${diffSummary.autoMerged} 个已自动并入 merged。`
                      : '当前工作表没有差异，也没有冲突。'}
                  </div>
                ) : (
                  attentionCells.map((cell) => {
                    const isActive = !!selectedCell && selectedCell.rowNumber === cell.row && selectedCell.colNumber === cell.col;
                    const rowMeta = rowsMetaMap.get(cell.row);
                    return (
                      <button
                        key={`${cell.row}:${cell.col}`}
                        type="button"
                        onClick={() => onSelectCell?.(cell.row - 1, cell.col - 1)}
                        style={{
                          textAlign: 'left',
                          border: isActive ? '1px solid #2563eb' : '1px solid #e5e7eb',
                          borderRadius: 8,
                          padding: '7px 8px',
                          backgroundColor: isActive ? '#eff6ff' : '#fff7ed',
                          color: '#111827',
                          cursor: 'pointer',
                        }}
                      >
                        <div style={{ display: 'grid', gap: 2 }}>
                          <span style={{ fontWeight: 600 }}>{cell.address}</span>
                          <span style={{ fontSize: 11, color: '#6b7280' }}>
                            视觉行 {cell.row}
                            {rowMeta
                              ? isSimpleMergeMode
                                ? ` · ours ${rowMeta.oursRowNumber ?? '-'} / theirs ${rowMeta.theirsRowNumber ?? '-'}`
                                : ` · base ${rowMeta.baseRowNumber ?? '-'} / ours ${rowMeta.oursRowNumber ?? '-'} / theirs ${rowMeta.theirsRowNumber ?? '-'}`
                              : ''}
                          </span>
                        </div>
                      </button>
                    );
                  })
                )}
              </div>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10, minWidth: 0, minHeight: 0 }}>
              <div
                style={{
                  border: '1px solid #f0d4b5',
                  borderRadius: 12,
                  padding: 10,
                  backgroundColor: '#fffbf5',
                  display: 'grid',
                  gap: 8,
                  flexShrink: 0,
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 12, color: '#9a3412' }}>快速处理</div>
                <div style={{ display: 'grid', gap: 6 }}>
                  <div style={{ fontSize: 11, color: '#7c2d12' }}>全部冲突</div>
                  <div style={{ display: 'grid', gridTemplateColumns: isSimpleMergeMode ? '1fr 1fr' : '1fr 1fr 1fr', gap: 6 }}>
                    {!isSimpleMergeMode && (
                      <button
                        type="button"
                        disabled={attentionCells.length === 0}
                        onClick={() => applyAllConflictsChoice('base')}
                        style={getQuickActionButtonStyle('base', attentionCells.length === 0)}
                      >
                        全部用 base
                      </button>
                    )}
                    <button
                      type="button"
                      disabled={attentionCells.length === 0}
                      onClick={() => applyAllConflictsChoice('ours')}
                      style={getQuickActionButtonStyle('ours', attentionCells.length === 0)}
                    >
                      全部用 ours
                    </button>
                    <button
                      type="button"
                      disabled={attentionCells.length === 0}
                      onClick={() => applyAllConflictsChoice('theirs')}
                      style={getQuickActionButtonStyle('theirs', attentionCells.length === 0)}
                    >
                      全部用 theirs
                    </button>
                  </div>
                </div>
                <div style={{ display: 'grid', gap: 6 }}>
                  <div style={{ fontSize: 11, color: '#7c2d12' }}>当前单元格</div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                    <button
                      type="button"
                      disabled={!selectedCanApplyCellChoice}
                      onClick={() => {
                        if (!selectedCell) return;
                        onResolveCell?.(selectedCell.rowNumber, selectedCell.colNumber);
                      }}
                      style={getQuickActionButtonStyle('merged', !selectedCanApplyCellChoice)}
                    >
                      接受当前结果
                    </button>
                    {!isSimpleMergeMode && (
                      <button
                        type="button"
                        disabled={!selectedCanApplyCellChoice}
                        onClick={() => onApplySelectedCellChoice?.('base')}
                        style={getQuickActionButtonStyle('base', !selectedCanApplyCellChoice)}
                      >
                        用 base
                      </button>
                    )}
                    <button
                      type="button"
                      disabled={!selectedCanApplyCellChoice}
                      onClick={() => onApplySelectedCellChoice?.('ours')}
                      style={getQuickActionButtonStyle('ours', !selectedCanApplyCellChoice)}
                    >
                      用 ours
                    </button>
                    <button
                      type="button"
                      disabled={!selectedCanApplyCellChoice}
                      onClick={() => onApplySelectedCellChoice?.('theirs')}
                      style={getQuickActionButtonStyle('theirs', !selectedCanApplyCellChoice)}
                    >
                      用 theirs
                    </button>
                  </div>
                </div>
                <div style={{ display: 'grid', gap: 6 }}>
                  <div style={{ fontSize: 11, color: '#7c2d12' }}>当前整列</div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedColumnHasDiff}
                      onClick={() => {
                        if (!selectedCell) return;
                        onApplyColumnChoice?.(selectedCell.colNumber, 'ours');
                      }}
                      style={getQuickActionButtonStyle('ours', !selectedCell || !selectedColumnHasDiff)}
                    >
                      采用 ours 整列
                    </button>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedColumnHasDiff}
                      onClick={() => {
                        if (!selectedCell) return;
                        onApplyColumnChoice?.(selectedCell.colNumber, 'theirs');
                      }}
                      style={getQuickActionButtonStyle('theirs', !selectedCell || !selectedColumnHasDiff)}
                    >
                      采用 theirs 整列
                    </button>
                  </div>
                </div>
                <div style={{ display: 'grid', gap: 6 }}>
                  <div style={{ fontSize: 11, color: '#7c2d12' }}>当前整行</div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedRowMeta || !selectedRowHasDiff}
                      onClick={() => {
                        if (!selectedCell) return;
                        onApplyRowChoice?.(selectedCell.rowNumber, 'ours');
                      }}
                      style={getQuickActionButtonStyle('ours', !selectedCell || !selectedRowMeta || !selectedRowHasDiff)}
                    >
                      采用 ours 整行
                    </button>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedRowMeta || !selectedRowHasDiff}
                      onClick={() => {
                        if (!selectedCell) return;
                        onApplyRowChoice?.(selectedCell.rowNumber, 'theirs');
                      }}
                      style={getQuickActionButtonStyle('theirs', !selectedCell || !selectedRowMeta || !selectedRowHasDiff)}
                    >
                      采用 theirs 整行
                    </button>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedRowMeta || !selectedRowHasDiff || !selectedRowMeta.oursRowNumber}
                      onClick={() => {
                        if (!selectedCell) return;
                        onDeleteRow?.(selectedCell.rowNumber);
                      }}
                      style={{
                        ...getQuickActionButtonStyle(
                          'danger',
                          !selectedCell || !selectedRowMeta || !selectedRowHasDiff || !selectedRowMeta.oursRowNumber,
                        ),
                        gridColumn: '1 / -1',
                      }}
                    >
                      删除结果中的这行
                    </button>
                  </div>
                </div>
              </div>
              <div style={{ flex: 1, minHeight: 0, overflowY: 'auto', display: 'grid', gap: 10 }}>
                {selectedCell ? (
                  <>
                    <div style={{ border: '1px solid #dbe3ef', borderRadius: 12, padding: 10, backgroundColor: '#fff', display: 'grid', gap: 8 }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                        <div>
                          <div style={{ fontSize: 12, color: '#6b7280' }}>当前单元格</div>
                          <div style={{ fontWeight: 700 }}>
                            {selectedCell.address}
                          </div>
                        </div>
                        <span
                          style={{
                            fontSize: 12,
                            color: statusTone,
                            border: `1px solid ${statusTone}`,
                            borderRadius: 999,
                            padding: '2px 8px',
                            backgroundColor: '#fff',
                            whiteSpace: 'nowrap',
                          }}
                        >
                          {getDisplayedCellStatusLabel(selectedCell)}
                        </span>
                      </div>
                      <div style={{ fontSize: 12, color: '#475569', lineHeight: 1.55 }}>
                        {describeCellDecision(selectedCell, selectedRowMeta, compareMode)}
                      </div>
                      {isProtectedCell(selectedCell) && (
                        <div
                          style={{
                            fontSize: 12,
                            color: '#92400e',
                            backgroundColor: '#fffbeb',
                            border: '1px solid #fed7aa',
                            borderRadius: 8,
                            padding: '6px 8px',
                          }}
                        >
                          {getProtectedCellMode(selectedCell) === 'formula'
                            ? '公式控制位：结果会保留模板公式，不会把当前显示值写回成文本。'
                            : getSharedControlMasterSheetName(selectedCell)
                              ? `共享控制位：这个位置跟随 ${getSharedControlMasterSheetName(selectedCell)} sheet 的主位同步，不支持单独写回。`
                              : '共享控制位：这个位置会跟随共享组结果一起变化，不支持单独写回。'}
                        </div>
                      )}
                      <div style={{ fontSize: 12, color: '#6b7280' }}>
                        视觉行：{selectedCell.rowNumber}
                      </div>
                      <div style={{ fontSize: 12, color: '#6b7280' }}>
                        {isSimpleMergeMode
                          ? `行映射：ours=${selectedRowMeta?.oursRowNumber ?? '-'} / theirs=${selectedRowMeta?.theirsRowNumber ?? '-'}`
                          : `行映射：base=${selectedRowMeta?.baseRowNumber ?? '-'} / ours=${selectedRowMeta?.oursRowNumber ?? '-'} / theirs=${selectedRowMeta?.theirsRowNumber ?? '-'}`}
                      </div>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                        {([
                          ...(isSimpleMergeMode ? [] : ([['base', selectedCell.baseValue]] as const)),
                          ['ours', selectedCell.oursValue],
                          ['theirs', selectedCell.theirsValue],
                          ['merged', selectedCell.mergedValue],
                        ] as const).map(([label, value]) => {
                          const palette = getCurrentValueCardPalette(label, selectedCell);
                          return (
                            <div
                              key={label}
                              style={{
                                border: `1px solid ${palette.borderColor}`,
                                borderRadius: 10,
                                padding: 6,
                                backgroundColor: palette.backgroundColor,
                              }}
                            >
                              <div style={{ fontSize: 11, color: palette.color, fontWeight: 700 }}>{label}</div>
                              <div
                                style={{
                                  marginTop: 3,
                                  fontSize: 12,
                                  color: palette.valueColor,
                                  wordBreak: 'break-word',
                                }}
                              >
                                {makeValueText(value)}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    <div style={{ border: '1px solid #dbe3ef', borderRadius: 12, padding: 10, backgroundColor: '#fff', display: 'grid', gap: 8 }}>
                      <div style={{ fontWeight: 700, color: '#1f2937' }}>行解释</div>
                      <div style={{ fontSize: 12, color: '#475569', lineHeight: 1.55 }}>{describeRowDecision(selectedRowMeta, compareMode)}</div>
                      <div style={{ fontSize: 12, color: '#6b7280' }}>
                        本行剩余冲突：{selectedRowAttentionCount}
                      </div>
                    </div>
                  </>
                ) : (
                  <div style={{ border: '1px dashed #cbd5e1', borderRadius: 12, padding: 16, backgroundColor: '#fff', fontSize: 12, color: '#64748b' }}>
                    从左侧{attentionListTitle}点一个单元格，下面会展示解释信息。
                  </div>
                )}
              </div>
            </div>
          </div>
          </div>
        )}
      </div>
    </div>
  );
};

export const MergeWorkbench = React.memo(MergeWorkbenchComponent);
