import React, { useEffect, useMemo, useRef, useState } from 'react';
import type { MergeCell, MergeColumnMeta, MergeRowMeta, RowStatus, ThreeWayRowResult } from '../main/preload';
import { VirtualGrid, VirtualGridRenderCtx } from './VirtualGrid';

type SourceSide = 'base' | 'ours' | 'theirs' | 'merged';

type MergeWorkbenchCell = {
  key: string;
  rowNumber: number;
  colNumber: number;
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  mergedValue: string | number | null;
  status: MergeCell['status'] | 'unchanged';
  resolved: boolean;
  isDiffCell: boolean;
};

export interface MergeWorkbenchProps {
  cells: MergeCell[];
  rowsMeta: MergeRowMeta[];
  columnsMeta?: MergeColumnMeta[];
  sourceRows: ThreeWayRowResult[];
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
  onDeleteRow?: (rowNumber: number) => void;
  resolvedCellKeys?: Set<string>;
  frozenRowCount?: number;
  primaryKeyCol?: number;
  sheetName?: string;
  basePath?: string | null;
  oursPath?: string | null;
  theirsPath?: string | null;
  mergedPath?: string | null;
  remainingCount: number;
  canUndo?: boolean;
  onUndo?: () => void;
}

const DEFAULT_FROZEN_HEADER_ROWS = 3;
const DEFAULT_COL_WIDTH = 108;
const ROW_HEADER_WIDTH = 62;
const FROZEN_BG = '#f2f4f7';
const BASE_BG = '#fff9e8';
const OURS_BG = '#ecf8ec';
const THEIRS_BG = '#fff0f0';
const RESOLVED_BG = '#f5f5f5';
const CONFLICT_OUTLINE = '#ff8a00';

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
      return '仅 ours 改动';
    case 'theirs-changed':
      return '仅 theirs 改动';
    case 'both-changed-same':
      return '双方同改同值';
    case 'conflict':
      return '冲突';
    case 'unchanged':
    default:
      return '无差异';
  }
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

const getPanelBackground = (status: MergeWorkbenchCell['status'], side: SourceSide, resolved: boolean) => {
  if (side === 'merged') {
    if (status === 'conflict' && !resolved) return '#fff1e6';
    if (status !== 'unchanged' && !resolved) return '#e8efff';
    if (status !== 'unchanged') return '#f4f8ff';
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

const getDefaultMergedValue = (cell: MergeCell | undefined, row: ThreeWayRowResult | undefined, colNumber: number) => {
  if (cell) return cell.mergedValue ?? null;
  if (!row) return null;
  const oursValue = row.ours[colNumber - 1] ?? null;
  const theirsValue = row.theirs[colNumber - 1] ?? null;
  const baseValue = row.base[colNumber - 1] ?? null;
  if (oursValue !== null && oursValue !== undefined) return oursValue;
  if (theirsValue !== null && theirsValue !== undefined) return theirsValue;
  return baseValue;
};

const describeCellDecision = (cell: MergeWorkbenchCell, rowMeta?: MergeRowMeta) => {
  const prefix =
    rowMeta && (rowMeta.oursStatus === 'ambiguous' || rowMeta.theirsStatus === 'ambiguous')
      ? '该行对齐存在歧义，先看清三侧原始值再决定。'
      : '';
  switch (cell.status) {
    case 'ours-changed':
      return `${prefix}ours 相对 base 发生变化，theirs 保持与 base 一致；如果你认可当前分支改动，直接采用当前 merged 结果即可。`.trim();
    case 'theirs-changed':
      return `${prefix}theirs 相对 base 发生变化，ours 保持与 base 一致；如果要把对方改动并入结果，优先采用 theirs。`.trim();
    case 'both-changed-same':
      return `${prefix}ours 和 theirs 都相对 base 改了，但结果相同；系统已经自动给出同一 merged 值，你只需要确认。`.trim();
    case 'conflict':
      return `${prefix}ours 和 theirs 都相对 base 改了，而且结果不同；这是人工决策点，需要你在 base / ours / theirs 之间做选择。`.trim();
    case 'unchanged':
    default:
      return rowMeta && (rowMeta.oursStatus === 'ambiguous' || rowMeta.theirsStatus === 'ambiguous')
        ? '当前单元格本身没有值冲突，但所在行的对齐并不稳定。'
        : '当前单元格在三侧没有形成需要人工处理的差异。';
  }
};

const describeRowDecision = (rowMeta?: MergeRowMeta) => {
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
  return '这行已经完成三方对齐；一般只需要对其中的冲突单元格做选择。';
};

const makeValueText = (value: string | number | null) => (value == null ? '∅' : String(value));

const MergeWorkbenchComponent: React.FC<MergeWorkbenchProps> = ({
  cells,
  rowsMeta,
  columnsMeta,
  sourceRows,
  layoutMode = 'full',
  selected,
  onSelectCell,
  onApplySelectedCellChoice,
  onApplyCellsChoice,
  onResolveCell,
  onApplyRowChoice,
  onDeleteRow,
  resolvedCellKeys,
  frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS,
  primaryKeyCol,
  sheetName,
  basePath,
  oursPath,
  theirsPath,
  mergedPath,
  remainingCount,
  canUndo = false,
  onUndo,
}) => {
  const baseScrollRef = useRef<HTMLDivElement | null>(null);
  const oursScrollRef = useRef<HTMLDivElement | null>(null);
  const theirsScrollRef = useRef<HTMLDivElement | null>(null);
  const mergedScrollRef = useRef<HTMLDivElement | null>(null);
  const leftPaneRef = useRef<HTMLDivElement | null>(null);
  const isSyncingXRef = useRef(false);
  const isSyncingYRef = useRef(false);
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [showDiffColumnsOnly, setShowDiffColumnsOnly] = useState(false);
  const [showUnresolvedRowsOnly, setShowUnresolvedRowsOnly] = useState(false);
  const [topPanelRatio, setTopPanelRatio] = useState(62);

  const syncScrollX = (from: SourceSide, scrollLeft: number) => {
    const targets = [
      { side: 'base' as const, ref: baseScrollRef },
      { side: 'ours' as const, ref: oursScrollRef },
      { side: 'theirs' as const, ref: theirsScrollRef },
      { side: 'merged' as const, ref: mergedScrollRef },
    ];
    if (isSyncingXRef.current) return;
    isSyncingXRef.current = true;
    targets.forEach((target) => {
      if (target.side === from || !target.ref.current) return;
      target.ref.current.scrollLeft = scrollLeft;
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
    if (isSyncingYRef.current) return;
    isSyncingYRef.current = true;
    targets.forEach((target) => {
      if (target.side === from || !target.ref.current) return;
      target.ref.current.scrollTop = scrollTop;
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

  const resolvedKeySet = resolvedCellKeys ?? new Set<string>();
  const normalizedPrimaryKeyCol =
    typeof primaryKeyCol === 'number' && primaryKeyCol >= 1 ? Math.floor(primaryKeyCol) : null;

  const diffColumns = useMemo(() => {
    const diffCols = new Set<number>();
    cells.forEach((cell) => diffCols.add(cell.col));
    if (normalizedPrimaryKeyCol) diffCols.add(normalizedPrimaryKeyCol);
    return Array.from(diffCols).sort((a, b) => a - b);
  }, [cells, normalizedPrimaryKeyCol]);

  const allColumns = useMemo(() => {
    if (columnsMeta && columnsMeta.length > 0) {
      return Array.from(new Set(columnsMeta.map((meta) => meta.col))).sort((a, b) => a - b);
    }
    const colCountFromSource = sourceRows[0]?.colCount ?? 0;
    if (colCountFromSource > 0) {
      return Array.from({ length: colCountFromSource }, (_, idx) => idx + 1);
    }
    return diffColumns;
  }, [columnsMeta, diffColumns, sourceRows]);

  const displayColumns = useMemo(() => {
    const cols = showDiffColumnsOnly ? diffColumns : allColumns;
    if (!normalizedPrimaryKeyCol || cols.includes(normalizedPrimaryKeyCol)) return cols;
    return [normalizedPrimaryKeyCol, ...cols].sort((a, b) => a - b);
  }, [allColumns, diffColumns, normalizedPrimaryKeyCol, showDiffColumnsOnly]);

  const allRowNumbers = useMemo(() => {
    if (rowsMeta.length > 0) {
      return rowsMeta.map((row) => row.visualRowNumber);
    }
    const rowNumbers = new Set<number>();
    cells.forEach((cell) => rowNumbers.add(cell.row));
    return Array.from(rowNumbers).sort((a, b) => a - b);
  }, [cells, rowsMeta]);

  const displayRowNumbers = useMemo(() => {
    if (!showUnresolvedRowsOnly) return allRowNumbers;
    const unresolvedRows = new Set<number>();
    cells.forEach((cell) => {
      if (cell.status === 'conflict' && !resolvedKeySet.has(`${cell.row}:${cell.col}`)) {
        unresolvedRows.add(cell.row);
      }
    });
    return allRowNumbers.filter((rowNumber) => unresolvedRows.has(rowNumber));
  }, [allRowNumbers, cells, resolvedKeySet, showUnresolvedRowsOnly]);

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
            rowNumber,
            colNumber,
            baseValue,
            oursValue,
            theirsValue,
            mergedValue: getDefaultMergedValue(mergeCell, sourceRow, colNumber),
            status: mergeCell?.status ?? 'unchanged',
            resolved: mergeCell ? resolvedKeySet.has(key) : true,
            isDiffCell: !!mergeCell,
          };
        });
      }),
    [displayColumns, displayRowNumbers, mergeCellMap, resolvedKeySet, sourceRowMap],
  );

  const unresolvedConflicts = useMemo(
    () =>
      cells
        .filter((cell) => cell.status === 'conflict' && !resolvedKeySet.has(`${cell.row}:${cell.col}`))
        .sort((a, b) => a.row - b.row || a.col - b.col),
    [cells, resolvedKeySet],
  );

  const unresolvedConflictCountByRow = useMemo(() => {
    const map = new Map<number, number>();
    unresolvedConflicts.forEach((cell) => {
      map.set(cell.row, (map.get(cell.row) ?? 0) + 1);
    });
    return map;
  }, [unresolvedConflicts]);

  const selectedGridRowIndex = selected ? displayRowNumbers.indexOf(selected.rowIndex + 1) : -1;
  const selectedGridColIndex = selected ? displayColumns.indexOf(selected.colIndex + 1) : -1;
  const selectedCell =
    selectedGridRowIndex >= 0 && selectedGridColIndex >= 0
      ? gridRows[selectedGridRowIndex]?.[selectedGridColIndex] ?? null
      : null;
  const selectedRowMeta = selectedCell ? rowsMetaMap.get(selectedCell.rowNumber) : undefined;
  const selectedIsDiffCell = selectedCell ? mergeCellMap.has(selectedCell.key) : false;
  const selectedRowConflictCount = selectedCell ? unresolvedConflictCountByRow.get(selectedCell.rowNumber) ?? 0 : 0;

  const jumpToNextConflict = () => {
    if (!onSelectCell || unresolvedConflicts.length === 0) return;
    const currentIndex = selectedCell
      ? unresolvedConflicts.findIndex(
          (cell) => cell.row === selectedCell.rowNumber && cell.col === selectedCell.colNumber,
        )
      : -1;
    const nextCell =
      unresolvedConflicts[currentIndex >= 0 ? (currentIndex + 1) % unresolvedConflicts.length : 0];
    if (!nextCell) return;
    onSelectCell(nextCell.row - 1, nextCell.col - 1);
  };

  const applyAllConflictsChoice = (source: 'base' | 'ours' | 'theirs') => {
    if (!onApplyCellsChoice || unresolvedConflicts.length === 0) return;
    onApplyCellsChoice(
      unresolvedConflicts.map((cell) => ({ rowNumber: cell.row, colNumber: cell.col })),
      source,
    );
  };

  const makeRowHeader = (rowIndex: number) => {
    const rowNumber = displayRowNumbers[rowIndex];
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
      ? `aligned ${colNumberToLabel(alignedCol)} | base=${meta.baseCol ? colNumberToLabel(meta.baseCol) : '-'} | ours=${meta.oursCol ? colNumberToLabel(meta.oursCol) : '-'} | theirs=${meta.theirsCol ? colNumberToLabel(meta.theirsCol) : '-'}`
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
          onMouseDown={() => onSelectCell?.(cell.rowNumber - 1, cell.colNumber - 1)}
          onClick={() => onSelectCell?.(cell.rowNumber - 1, cell.colNumber - 1)}
          title={`${side}: ${makeValueText(value)}\n状态: ${getCellStatusLabel(cell.status)}`}
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
        style.backgroundColor = getPanelBackground(cell.status, side, cell.resolved);
        if (cell.status === 'conflict' && !cell.resolved) {
          style.boxShadow = `inset 0 0 0 2px ${CONFLICT_OUTLINE}`;
        } else if (cell.status !== 'unchanged' && !cell.resolved) {
          style.borderLeft = `3px solid ${getPanelAccent(side)}`;
        } else if (side === 'merged' && cell.isDiffCell) {
          style.borderLeft = `3px solid ${getPanelAccent(side)}`;
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
  }> = [
    { side: 'base', title: 'base（只读）', path: basePath, ref: baseScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: true },
    { side: 'ours', title: 'ours（只读）', path: oursPath, ref: oursScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: false },
    { side: 'theirs', title: 'theirs（只读）', path: theirsPath, ref: theirsScrollRef as React.RefObject<HTMLDivElement>, showRowHeader: false },
  ];

  const statusTone =
    selectedCell?.status === 'conflict'
      ? '#b45309'
      : selectedCell?.status === 'both-changed-same'
        ? '#334155'
        : selectedCell?.status === 'ours-changed'
          ? '#166534'
          : selectedCell?.status === 'theirs-changed'
            ? '#b91c1c'
            : '#475569';
  const showGridPane = layoutMode !== 'panel-only';
  const showDecisionPane = layoutMode !== 'grids-only';
  const showToolbar = layoutMode !== 'panel-only';

  return displayColumns.length === 0 || displayRowNumbers.length === 0 ? (
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
          <span style={{ fontSize: 12, color: '#4b5563' }}>三侧源数据只读；所有选择都通过右侧决策栏完成。</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>未解决冲突: {remainingCount}</span>
          <span style={{ fontSize: 12, color: '#4b5563' }}>显示列: {displayColumns.length}</span>
          <label style={{ fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 4, color: '#334155' }}>
            <input
              type="checkbox"
              checked={showDiffColumnsOnly}
              onChange={(event) => setShowDiffColumnsOnly(event.target.checked)}
            />
            仅差异列
          </label>
          <label style={{ fontSize: 12, display: 'inline-flex', alignItems: 'center', gap: 4, color: '#334155' }}>
            <input
              type="checkbox"
              checked={showUnresolvedRowsOnly}
              onChange={(event) => setShowUnresolvedRowsOnly(event.target.checked)}
            />
            仅未解决行
          </label>
          <button type="button" onClick={jumpToNextConflict} disabled={unresolvedConflicts.length === 0}>
            下一个冲突
          </button>
          <button type="button" onClick={onUndo} disabled={!canUndo}>
            撤销上一步
          </button>
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
              gridTemplateColumns: 'repeat(3, minmax(0, 1fr))',
              gap: 8,
              minWidth: 0,
              overflow: 'hidden',
            }}
          >
            {panelSpecs.map((panel) => (
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
                    rows={gridRows}
                    frozenRowCount={Math.max(0, Math.floor(frozenRowCount))}
                    frozenColCount={0}
                    rowHeaderWidth={panel.showRowHeader ? ROW_HEADER_WIDTH : 0}
                    showRowHeader={panel.showRowHeader}
                    renderRowHeader={(rowIndex) => makeRowHeader(rowIndex)}
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
                      selectedGridRowIndex >= 0 && selectedGridColIndex >= 0
                        ? { rowIndex: selectedGridRowIndex, colIndex: selectedGridColIndex }
                        : null
                    }
                    scrollToCellAlign="center"
                  />
                </div>
              </div>
            ))}
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
            title="拖动调整上方三栏和下方 merged 面板高度"
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
              <span style={{ fontSize: 11, color: '#6b7280' }}>橙色=冲突未解决 · 蓝色=差异已入结果</span>
            </div>
            <div style={{ flex: 1, minHeight: 0 }}>
              <VirtualGrid<MergeWorkbenchCell>
                rows={gridRows}
                frozenRowCount={Math.max(0, Math.floor(frozenRowCount))}
                frozenColCount={0}
                rowHeaderWidth={ROW_HEADER_WIDTH}
                showRowHeader
                renderRowHeader={(rowIndex) => makeRowHeader(rowIndex)}
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
                  selectedGridRowIndex >= 0 && selectedGridColIndex >= 0
                    ? { rowIndex: selectedGridRowIndex, colIndex: selectedGridColIndex }
                    : null
                }
                scrollToCellAlign="center"
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
              左侧冲突列表可滚动定位，右侧固定展示当前选中项解释和快速处理。
            </div>
          </div>
          <div style={{ padding: 12, display: 'grid', gap: 12, borderBottom: '1px solid #e5e7eb', backgroundColor: '#fbfcfe', flexShrink: 0 }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: 8 }}>
              <div style={{ border: '1px solid #e5e7eb', borderRadius: 10, padding: 10, backgroundColor: '#fff' }}>
                <div style={{ fontSize: 11, color: '#6b7280' }}>未解决冲突</div>
                <div style={{ marginTop: 4, fontSize: 20, fontWeight: 700 }}>{remainingCount}</div>
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
                <div style={{ fontWeight: 700, color: '#1f2937' }}>冲突列表</div>
                <div style={{ fontSize: 11, color: '#6b7280' }}>可滚动，点击跳转</div>
              </div>
              <div style={{ flex: 1, minHeight: 0, overflowY: 'scroll', padding: 8, display: 'grid', gap: 6 }}>
                {unresolvedConflicts.length === 0 ? (
                  <div style={{ fontSize: 12, color: '#6b7280', padding: 6 }}>当前工作表没有未解决冲突。</div>
                ) : (
                  unresolvedConflicts.map((cell) => {
                    const isActive = !!selectedCell && selectedCell.rowNumber === cell.row && selectedCell.colNumber === cell.col;
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
                        {colNumberToLabel(cell.col)}
                        {cell.row}
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
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 6 }}>
                    <button type="button" disabled={unresolvedConflicts.length === 0} onClick={() => applyAllConflictsChoice('base')}>
                      全部用 base
                    </button>
                    <button type="button" disabled={unresolvedConflicts.length === 0} onClick={() => applyAllConflictsChoice('ours')}>
                      全部用 ours
                    </button>
                    <button type="button" disabled={unresolvedConflicts.length === 0} onClick={() => applyAllConflictsChoice('theirs')}>
                      全部用 theirs
                    </button>
                  </div>
                </div>
                <div style={{ display: 'grid', gap: 6 }}>
                  <div style={{ fontSize: 11, color: '#7c2d12' }}>当前单元格</div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedIsDiffCell}
                      onClick={() => {
                        if (!selectedCell) return;
                        onResolveCell?.(selectedCell.rowNumber, selectedCell.colNumber);
                      }}
                    >
                      接受当前结果
                    </button>
                    <button type="button" disabled={!selectedCell || !selectedIsDiffCell} onClick={() => onApplySelectedCellChoice?.('base')}>
                      用 base
                    </button>
                    <button type="button" disabled={!selectedCell || !selectedIsDiffCell} onClick={() => onApplySelectedCellChoice?.('ours')}>
                      用 ours
                    </button>
                    <button type="button" disabled={!selectedCell || !selectedIsDiffCell} onClick={() => onApplySelectedCellChoice?.('theirs')}>
                      用 theirs
                    </button>
                  </div>
                </div>
                <div style={{ display: 'grid', gap: 6 }}>
                  <div style={{ fontSize: 11, color: '#7c2d12' }}>当前整行</div>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedRowMeta}
                      onClick={() => {
                        if (!selectedCell) return;
                        onApplyRowChoice?.(selectedCell.rowNumber, 'ours');
                      }}
                    >
                      采用 ours 整行
                    </button>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedRowMeta}
                      onClick={() => {
                        if (!selectedCell) return;
                        onApplyRowChoice?.(selectedCell.rowNumber, 'theirs');
                      }}
                    >
                      采用 theirs 整行
                    </button>
                    <button
                      type="button"
                      disabled={!selectedCell || !selectedRowMeta || !selectedRowMeta.oursRowNumber}
                      onClick={() => {
                        if (!selectedCell) return;
                        onDeleteRow?.(selectedCell.rowNumber);
                      }}
                      style={{ gridColumn: '1 / -1', color: '#b42318' }}
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
                            {colNumberToLabel(selectedCell.colNumber)}
                            {selectedCell.rowNumber}
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
                          {getCellStatusLabel(selectedCell.status)}
                        </span>
                      </div>
                      <div style={{ fontSize: 12, color: '#475569', lineHeight: 1.55 }}>
                        {describeCellDecision(selectedCell, selectedRowMeta)}
                      </div>
                      <div style={{ fontSize: 12, color: '#6b7280' }}>
                        行映射：base={selectedRowMeta?.baseRowNumber ?? '-'} / ours={selectedRowMeta?.oursRowNumber ?? '-'} / theirs={selectedRowMeta?.theirsRowNumber ?? '-'}
                      </div>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 6 }}>
                        {([
                          ['base', selectedCell.baseValue],
                          ['ours', selectedCell.oursValue],
                          ['theirs', selectedCell.theirsValue],
                          ['merged', selectedCell.mergedValue],
                        ] as const).map(([label, value]) => (
                          <div key={label} style={{ border: '1px solid #edf0f5', borderRadius: 10, padding: 6 }}>
                            <div style={{ fontSize: 11, color: '#6b7280' }}>{label}</div>
                            <div style={{ marginTop: 3, fontSize: 12, color: '#111827', wordBreak: 'break-word' }}>{makeValueText(value)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                    <div style={{ border: '1px solid #dbe3ef', borderRadius: 12, padding: 10, backgroundColor: '#fff', display: 'grid', gap: 8 }}>
                      <div style={{ fontWeight: 700, color: '#1f2937' }}>行解释</div>
                      <div style={{ fontSize: 12, color: '#475569', lineHeight: 1.55 }}>{describeRowDecision(selectedRowMeta)}</div>
                      <div style={{ fontSize: 12, color: '#6b7280' }}>本行未解决冲突：{selectedRowConflictCount}</div>
                    </div>
                  </>
                ) : (
                  <div style={{ border: '1px dashed #cbd5e1', borderRadius: 12, padding: 16, backgroundColor: '#fff', fontSize: 12, color: '#64748b' }}>
                    从左侧冲突列表点一个单元格，下面会展示解释信息。
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
