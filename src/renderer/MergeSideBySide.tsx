import React, { useEffect, useMemo, useRef, useState } from 'react';
import type { MergeCell, MergeRowMeta, RowStatus } from '../main/preload';
import { VirtualGrid, VirtualGridRenderCtx } from './VirtualGrid';

const ROW_HEIGHT = 24; // px, approximate row height for virtualization
const OVERSCAN_ROWS = 8; // render a few extra rows above/below viewport for smooth scroll
const DEFAULT_FROZEN_HEADER_ROWS = 3; // merge/diff 视图中固定展示的前几行默认值

/**
 * 三方对比视图中使用的 side-by-side 表格组件。
 *
 * 左右两侧分别显示 ours / theirs，行集只包含有差异的单元格所在行，
 * 并通过虚拟滚动与水平滚动联动减少 DOM 数量、提升性能。
 */
export interface MergeSideBySideProps {
  // 性能优化：只传差异单元格列表（稀疏结构）
  cells: MergeCell[];
  rowsMeta?: MergeRowMeta[];
  selected?: { rowIndex: number; colIndex: number } | null;
  onSelectCell?: (rowIndex: number, colIndex: number) => void;
  onApplyRowChoice?: (rowNumber: number, source: 'ours' | 'theirs') => void;
  onApplyCellChoice?: (rowNumber: number, colNumber: number, source: 'ours' | 'theirs') => void;
  onApplyCellsChoice?: (keys: { rowNumber: number; colNumber: number }[], source: 'ours' | 'theirs') => void;
  /** 已确认合并（resolved）的单元格 key 集合，key 格式为 "row:col"（1-based） */
  resolvedCellKeys?: Set<string>;
  /** 冻结在顶部展示的行数，可配置 */
  frozenRowCount?: number;
  /** 主键列（1-based）；传入时会在 diff 视图中额外展示该列 */
  primaryKeyCol?: number;
  /** ours/theirs 文件路径，用于顶部标识 */
  oursPath?: string | null;
  theirsPath?: string | null;
  /** base 文件路径 */
  basePath?: string | null;
  /** 是否显示全表 */
  showFullTables?: boolean;
  /** ours 全表数据（二维数组，仅值） */
  fullOursRows?: (string | number | null)[][];
  /** theirs 全表数据（二维数组，仅值） */
  fullTheirsRows?: (string | number | null)[][];
}
const MERGED_COLOR = '#fafafa';
const FROZEN_COLOR = '#e6e6e6';

const getBackgroundColor = (status: MergeCell['status'], side: 'ours' | 'theirs'): string => {
  switch (status) {
    case 'unchanged':
      return 'white';
    case 'ours-changed':
      return side === 'ours' ? '#d4f8d4' : 'white';
    case 'theirs-changed':
      // 需求：theirs 侧用 red
      return side === 'theirs' ? '#ffc8c8' : 'white';
    case 'both-changed-same':
      return MERGED_COLOR; // both sides same -> merge color
    case 'conflict':
      // 需求：当 ours/theirs 不同（冲突）时：ours=green, theirs=red
      return side === 'ours' ? '#d4f8d4' : '#ffc8c8';
    default:
      return 'white';
  }
};

const colNumberToLabel = (colNumber: number): string => {
  // Excel 列号是 1-based：1 -> A, 26 -> Z, 27 -> AA ...
  let n = Math.max(1, Math.floor(colNumber));
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode('A'.charCodeAt(0) + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
};

const getColumnLabels = (cols: number[]): string[] => {
  return cols.map((c) => colNumberToLabel(c));
};

const DATA_COL_WIDTH = 160; // px, keep ours/theirs columns visually aligned

const MergeSideBySideComponent: React.FC<MergeSideBySideProps> = ({
  cells,
  rowsMeta,
  selected,
  onSelectCell,
  onApplyRowChoice,
  onApplyCellChoice,
  onApplyCellsChoice,
  resolvedCellKeys,
  frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS,
  primaryKeyCol,
  oursPath,
  theirsPath,
  basePath,
  showFullTables,
  fullOursRows,
  fullTheirsRows,
}) => {
  const useFullTables =
    !!showFullTables && Array.isArray(fullOursRows) && Array.isArray(fullTheirsRows);

  // 水平/竖直滚动同步：左右两边 VirtualGrid 的 scrollLeft/scrollTop 要联动
  const oursScrollRef = useRef<HTMLDivElement | null>(null);
  const theirsScrollRef = useRef<HTMLDivElement | null>(null);
  const isSyncingHorizontalRef = useRef(false);
  const isSyncingVerticalRef = useRef(false);

  const syncScrollX = (from: 'ours' | 'theirs', scrollLeft: number) => {
    const otherRef = from === 'ours' ? theirsScrollRef : oursScrollRef;
    if (!otherRef.current) return;
    if (isSyncingHorizontalRef.current) return;
    isSyncingHorizontalRef.current = true;
    otherRef.current.scrollLeft = scrollLeft;
    requestAnimationFrame(() => {
      isSyncingHorizontalRef.current = false;
    });
  };

  const syncScrollY = (from: 'ours' | 'theirs', scrollTop: number) => {
    const otherRef = from === 'ours' ? theirsScrollRef : oursScrollRef;
    if (!otherRef.current) return;
    if (isSyncingVerticalRef.current) return;
    isSyncingVerticalRef.current = true;
    otherRef.current.scrollTop = scrollTop;
    requestAnimationFrame(() => {
      isSyncingVerticalRef.current = false;
    });
  };

  const fullColCount = useMemo(() => {
    if (!useFullTables) return 0;
    const oursMax = (fullOursRows ?? []).reduce((m, r) => Math.max(m, r?.length ?? 0), 0);
    const theirsMax = (fullTheirsRows ?? []).reduce((m, r) => Math.max(m, r?.length ?? 0), 0);
    return Math.max(oursMax, theirsMax, 1);
  }, [useFullTables, fullOursRows, fullTheirsRows]);
  const padRow = (row: (string | number | null)[], count: number) => {
    if (row.length >= count) return row.slice(0, count);
    return [...row, ...Array(count - row.length).fill(null)];
  };
  const fullOursGrid: (string | number | null)[][] = useMemo(() => {
    if (!useFullTables) return [];
    return (fullOursRows ?? []).map((r) => padRow(r ?? [], fullColCount));
  }, [useFullTables, fullOursRows, fullColCount]);
  const fullTheirsGrid: (string | number | null)[][] = useMemo(() => {
    if (!useFullTables) return [];
    return (fullTheirsRows ?? []).map((r) => padRow(r ?? [], fullColCount));
  }, [useFullTables, fullTheirsRows, fullColCount]);

  const cellMap = useMemo(() => {
    const m = new Map<string, MergeCell>();
    cells.forEach((c) => {
      m.set(`${c.row}:${c.col}`, c);
    });
    return m;
  }, [cells]);

  // 只展示有差异的列（cells 本身已是 status !== 'unchanged'）
  const diffColumns = useMemo(() => {
    const cols = new Set<number>();
    cells.forEach((cell) => cols.add(cell.col));
    return Array.from(cols).sort((a, b) => a - b);
  }, [cells]);
  const rowsMetaMap = useMemo(() => {
    const m = new Map<number, MergeRowMeta>();
    (rowsMeta ?? []).forEach((r) => m.set(r.visualRowNumber, r));
    return m;
  }, [rowsMeta]);

  // 差异行号（对齐后的视觉行号）
  const diffRowNumbers = useMemo(() => {
    const rs = new Set<number>();
    cells.forEach((cell) => rs.add(cell.row));
    const headerCount = Math.max(0, Math.floor(frozenRowCount));
    for (let i = 1; i <= headerCount; i += 1) {
      rs.add(i);
    }
    return Array.from(rs).sort((a, b) => a - b);
  }, [cells, frozenRowCount]);
  const normalizedPrimaryKeyCol =
    typeof primaryKeyCol === 'number' && primaryKeyCol >= 1 ? Math.floor(primaryKeyCol) : null;

  const displayColumns = useMemo(() => {
    if (!normalizedPrimaryKeyCol) return diffColumns;
    const cols = new Set<number>(diffColumns);
    cols.add(normalizedPrimaryKeyCol);
    return Array.from(cols).sort((a, b) => a - b);
  }, [diffColumns, normalizedPrimaryKeyCol]);

  const displayCellMap = useMemo(() => {
    const m = new Map(cellMap);
    if (!normalizedPrimaryKeyCol) return m;
    for (const rowNumber of diffRowNumbers) {
      const key = `${rowNumber}:${normalizedPrimaryKeyCol}`;
      if (m.has(key)) continue;
      const meta = rowsMetaMap.get(rowNumber);
      const keyValue = meta?.key ?? null;
      const addressRow =
        meta?.oursRowNumber ?? meta?.baseRowNumber ?? meta?.theirsRowNumber ?? rowNumber;
      const address = `${colNumberToLabel(normalizedPrimaryKeyCol)}${addressRow}`;
      m.set(key, {
        address,
        row: rowNumber,
        col: normalizedPrimaryKeyCol,
        baseValue: keyValue,
        oursValue: keyValue,
        theirsValue: keyValue,
        status: 'unchanged',
        mergedValue: keyValue,
      });
    }
    return m;
  }, [cellMap, normalizedPrimaryKeyCol, diffRowNumbers, rowsMetaMap]);

  const hasDiffData = cells.length > 0 && displayColumns.length > 0 && diffRowNumbers.length > 0;

  // 将原始 rows + diffColumns/diffRowNumbers 转成 VirtualGrid 需要的矩阵
  const gridRowNumbers = useMemo(() => diffRowNumbers, [diffRowNumbers]);

  const gridRows: (MergeCell | null)[][] = useMemo(
    () =>
      gridRowNumbers.map((rowNumber) =>
        displayColumns.map((colNumber) => displayCellMap.get(`${rowNumber}:${colNumber}`) ?? null),
      ),
    [displayCellMap, gridRowNumbers, displayColumns],
  );

  // 两侧共享列宽，避免左右/表头/内容出现 1px 累积偏差或拖拽后不同步
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  useEffect(() => {
    const count = useFullTables ? fullColCount : displayColumns.length;
    setColumnWidths((prev) => {
      if (prev.length === count) return prev;
      return Array(count).fill(DATA_COL_WIDTH);
    });
  }, [displayColumns.length, useFullTables, fullColCount]);

  const getRowStatusIndicator = (status: RowStatus | undefined) => {
    switch (status) {
      case 'added':
        return { symbol: '+', color: '#2e7d32' };
      case 'deleted':
        return { symbol: '-', color: '#b00020' };
      case 'modified':
        return { symbol: '~', color: '#ef6c00' };
      case 'ambiguous':
        return { symbol: '?', color: '#6d6d6d' };
      case 'unchanged':
      default:
        return { symbol: '', color: '#666' };
    }
  };

  const makeRowHeaderRenderer =
    (side: 'ours' | 'theirs') =>
    (gridRowIndex: number, rowCells: (MergeCell | null)[]) => {
      const visualRowNumber = diffRowNumbers[gridRowIndex];
      const meta = visualRowNumber ? rowsMetaMap.get(visualRowNumber) : undefined;
      const status = side === 'ours' ? meta?.oursStatus : meta?.theirsStatus;
      const rowNumber =
        side === 'ours'
          ? meta?.oursRowNumber ?? meta?.baseRowNumber ?? visualRowNumber
          : meta?.theirsRowNumber ?? meta?.baseRowNumber ?? visualRowNumber;
      const indicator = getRowStatusIndicator(status);
      const display = rowNumber ?? '';
      const baseNum = meta?.baseRowNumber ?? '-';
      const oursNum = meta?.oursRowNumber ?? '-';
      const theirsNum = meta?.theirsRowNumber ?? '-';
      const originalLabel = `b${baseNum}/o${oursNum}/t${theirsNum}`;
      const sim = side === 'ours' ? meta?.oursSimilarity : meta?.theirsSimilarity;
      const simLabel = typeof sim === 'number' ? `s${sim.toFixed(2)}` : '';
      const title = `原始行号: base=${baseNum}, ours=${oursNum}, theirs=${theirsNum}`;
      return (
        <div
          title={title}
          style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: 4, overflow: 'hidden' }}
        >
          {indicator.symbol && (
            <span style={{ color: indicator.color, fontWeight: 700 }}>{indicator.symbol}</span>
          )}
          <span style={{ whiteSpace: 'nowrap' }}>{display}</span>
          <span style={{ fontSize: 10, color: '#666', whiteSpace: 'nowrap' }}>{originalLabel}</span>
          {simLabel && (
            <span style={{ fontSize: 10, color: '#888', whiteSpace: 'nowrap' }}>{simLabel}</span>
          )}
        </div>
      );
    };

  const renderFullRowHeader = (rowIndex: number) => rowIndex + 1;
  const makeFullRenderCell =
    (_side: 'ours' | 'theirs') =>
    (cell: string | number | null, ctx: VirtualGridRenderCtx) => {
      const value = cell == null ? '' : String(cell);
      const handleClick = () => {
        if (onSelectCell) onSelectCell(ctx.rowIndex, ctx.colIndex);
      };
      return (
        <div
          onMouseDown={handleClick}
          onClick={handleClick}
          title={value}
          style={{
            width: '100%',
            height: '100%',
            boxSizing: 'border-box',
            backgroundColor: 'transparent',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            cursor: 'pointer',
            userSelect: 'none',
          }}
        >
          {value}
        </div>
      );
    };
  const getFullCellStyle = (_cell: any, ctx: VirtualGridRenderCtx) => {
    const style: React.CSSProperties = {};
    if (ctx.isFrozenRow || ctx.isFrozenCol) {
      style.backgroundColor = FROZEN_COLOR;
    }
    if (selected) {
      if (selected.rowIndex === ctx.rowIndex && selected.colIndex === ctx.colIndex) {
        style.outline = '2px solid #ff8000';
        style.outlineOffset = '-2px';
        style.position = 'relative';
        style.zIndex = 6;
      }
    }
    return style;
  };

  const [contextMenu, setContextMenu] = useState<
    | {
        type: 'row';
        x: number;
        y: number;
        rowNumber: number;
        source: 'ours' | 'theirs';
      }
    | {
        type: 'cell';
        x: number;
        y: number;
        rowNumber: number;
        colNumber: number;
        source: 'ours' | 'theirs';
      }
    | {
        type: 'cells';
        x: number;
        y: number;
        rowNumber: number;
        colNumber: number;
        source: 'ours' | 'theirs';
      }
    | null
  >(null);

  useEffect(() => {
    if (!contextMenu) return;
    const close = () => setContextMenu(null);
    window.addEventListener('click', close);
    window.addEventListener('blur', close);
    return () => {
      window.removeEventListener('click', close);
      window.removeEventListener('blur', close);
    };
  }, [contextMenu]);

  const handleRowHeaderContextMenu = (
    source: 'ours' | 'theirs',
    gridRowIndex: number,
    e: React.MouseEvent<HTMLTableCellElement>,
  ) => {
    e.preventDefault();
    e.stopPropagation();
    const rowNumber = diffRowNumbers[gridRowIndex];
    if (!rowNumber) return;
    setContextMenu({ type: 'row', x: e.clientX, y: e.clientY, rowNumber, source });
  };

  // 框选多选：以 diffRows/diffColumns 的矩形范围来计算选中单元格 key（仅包含存在的 diff cell）
  const [selectedCellKeys, setSelectedCellKeys] = useState<Set<string>>(new Set());
  const dragStartRef = useRef<{ rowNumber: number; colNumber: number } | null>(null);
  const dragEndRef = useRef<{ rowNumber: number; colNumber: number } | null>(null);
  const isDraggingRef = useRef(false);
  const dragMovedRef = useRef(false);
  const selectionBounds = useMemo(() => {
    if (!selectedCellKeys || selectedCellKeys.size === 0) return null;
    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;
    selectedCellKeys.forEach((k) => {
      const [rStr, cStr] = k.split(':');
      const r = Number(rStr);
      const c = Number(cStr);
      if (Number.isNaN(r) || Number.isNaN(c)) return;
      if (r < minRow) minRow = r;
      if (r > maxRow) maxRow = r;
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    });
    if (!Number.isFinite(minRow) || !Number.isFinite(minCol)) return null;
    return { minRow, maxRow, minCol, maxCol };
  }, [selectedCellKeys]);

  useEffect(() => {
    const onUp = () => {
      if (!isDraggingRef.current) return;
      isDraggingRef.current = false;
      dragStartRef.current = null;
      dragEndRef.current = null;
      // 结束拖拽后允许 click 继续工作
      setTimeout(() => {
        dragMovedRef.current = false;
      }, 0);
    };
    window.addEventListener('mouseup', onUp);
    return () => window.removeEventListener('mouseup', onUp);
  }, []);

  const computeSelectionKeys = (a: { rowNumber: number; colNumber: number }, b: { rowNumber: number; colNumber: number }) => {
    const r1 = Math.min(a.rowNumber, b.rowNumber);
    const r2 = Math.max(a.rowNumber, b.rowNumber);
    const c1 = Math.min(a.colNumber, b.colNumber);
    const c2 = Math.max(a.colNumber, b.colNumber);
    const keys: string[] = [];
    for (let r = r1; r <= r2; r += 1) {
      for (let c = c1; c <= c2; c += 1) {
        const key = `${r}:${c}`;
        if (cellMap.has(key)) keys.push(key);
      }
    }
    return new Set(keys);
  };

  const beginDragSelect = (cell: MergeCell) => {
    isDraggingRef.current = true;
    dragMovedRef.current = false;
    const p = { rowNumber: cell.row, colNumber: cell.col };
    dragStartRef.current = p;
    dragEndRef.current = p;
    setSelectedCellKeys(new Set([`${cell.row}:${cell.col}`]));
  };

  const updateDragSelect = (cell: MergeCell) => {
    if (!isDraggingRef.current) return;
    const start = dragStartRef.current;
    if (!start) return;
    const end = { rowNumber: cell.row, colNumber: cell.col };
    dragEndRef.current = end;
    if (end.rowNumber !== start.rowNumber || end.colNumber !== start.colNumber) {
      dragMovedRef.current = true;
    }
    setSelectedCellKeys(computeSelectionKeys(start, end));
  };

  const handleCellContextMenu = (
    source: 'ours' | 'theirs',
    cell: MergeCell,
    e: React.MouseEvent<HTMLDivElement>,
  ) => {
    e.preventDefault();
    e.stopPropagation();

    const key = `${cell.row}:${cell.col}`;
    const isInMultiSelection = selectedCellKeys.size > 1 && selectedCellKeys.has(key);

    // 右键时也选中该单元格，便于顶部/底部信息同步
    if (onSelectCell) {
      onSelectCell(cell.row - 1, cell.col - 1);
    }

    setContextMenu({
      type: isInMultiSelection ? 'cells' : 'cell',
      x: e.clientX,
      y: e.clientY,
      rowNumber: cell.row,
      colNumber: cell.col,
      source,
    });
  };

  const makeRenderCell = (side: 'ours' | 'theirs') =>
    (cell: MergeCell | null, ctx: VirtualGridRenderCtx) => {
      if (!cell) return null;
      const value = side === 'ours' ? cell.oursValue : cell.theirsValue;
      const sourceRowIndex = cell.row - 1;
      const sourceColIndex = cell.col - 1;
      const handleClick = () => {
        // 拖拽框选过程中避免 click 覆盖选择
        if (dragMovedRef.current) return;
        if (onSelectCell) {
          onSelectCell(sourceRowIndex, sourceColIndex);
        }
        setSelectedCellKeys(new Set([`${cell.row}:${cell.col}`]));
      };

      return (
        <div
          onMouseDown={(e) => {
            if (e.button !== 0) return;
            // 非框选状态下：鼠标按下新单元格时立即清除之前的选择，并选中当前格
            if (onSelectCell) {
              onSelectCell(cell.row - 1, cell.col - 1);
            }
            setSelectedCellKeys(new Set([`${cell.row}:${cell.col}`]));
            beginDragSelect(cell);
          }}
          onMouseEnter={() => updateDragSelect(cell)}
          onClick={handleClick}
          onContextMenu={(e) => handleCellContextMenu(side, cell, e)}
          title={`地址: ${cell.address}\nbase: ${cell.baseValue ?? ''}\nours: ${cell.oursValue ?? ''}\ntheirs: ${cell.theirsValue ?? ''}`}
          style={{
            width: '100%',
            height: '100%',
            boxSizing: 'border-box',
            backgroundColor: 'transparent',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            cursor: 'pointer',
            userSelect: 'none',
          }}
        >
          {value === null ? '' : String(value)}
        </div>
      );
    };

  const makeGetCellStyle = (side: 'ours' | 'theirs') =>
    (cell: MergeCell | null, ctx: VirtualGridRenderCtx): React.CSSProperties => {
      const style: React.CSSProperties = {};
      if (cell) {
        const key = `${cell.row}:${cell.col}`;
        if (resolvedCellKeys && resolvedCellKeys.has(key)) {
          // 已确认合并：两侧都用浅灰表示“处理过”
          style.backgroundColor = MERGED_COLOR;
        } else if (cell.status !== 'unchanged') {
          style.backgroundColor = getBackgroundColor(cell.status, side);
        } else if (ctx.isFrozenRow || ctx.isFrozenCol) {
          style.backgroundColor = FROZEN_COLOR;
        }
      } else if (ctx.isFrozenRow || ctx.isFrozenCol) {
        style.backgroundColor = FROZEN_COLOR;
      }

      if (cell) {
        const key = `${cell.row}:${cell.col}`;
        if (selectedCellKeys.has(key) && selectionBounds) {
          const { minRow, maxRow, minCol, maxCol } = selectionBounds;
          const shadows: string[] = [];
          if (cell.row === minRow) shadows.push('inset 0 2px 0 0 #ff8000');
          if (cell.row === maxRow) shadows.push('inset 0 -2px 0 0 #ff8000');
          if (cell.col === minCol) shadows.push('inset 2px 0 0 0 #ff8000');
          if (cell.col === maxCol) shadows.push('inset -2px 0 0 0 #ff8000');
          if (shadows.length > 0) {
            style.boxShadow = shadows.join(', ');
            style.position = 'relative';
            style.zIndex = 6;
          }
        }
      }

      if (cell && selected) {
        const sourceRowIndex = cell.row - 1;
        const sourceColIndex = cell.col - 1;
        const isSelected =
          selected.rowIndex === sourceRowIndex && selected.colIndex === sourceColIndex;
        if (isSelected) {
          style.outline = '2px solid #ff8000';
          style.outlineOffset = '-2px';
          style.position = 'relative';
          style.zIndex = 6;
        }
      }

      return style;
    };

  const renderDiffHeaderCell = (colIndex: number) => {
    const colNumber = displayColumns[colIndex];
    return colNumberToLabel(colNumber);
  };
  const renderFullHeaderCell = (colIndex: number) => colNumberToLabel(colIndex + 1);

  const oursRenderCell: (cell: any, ctx: VirtualGridRenderCtx) => React.ReactNode = useFullTables
    ? makeFullRenderCell('ours')
    : makeRenderCell('ours');
  const theirsRenderCell: (cell: any, ctx: VirtualGridRenderCtx) => React.ReactNode = useFullTables
    ? makeFullRenderCell('theirs')
    : makeRenderCell('theirs');
  const oursGetCellStyle: (cell: any, ctx: VirtualGridRenderCtx) => React.CSSProperties | undefined = useFullTables
    ? getFullCellStyle
    : makeGetCellStyle('ours');
  const theirsGetCellStyle: (cell: any, ctx: VirtualGridRenderCtx) => React.CSSProperties | undefined = useFullTables
    ? getFullCellStyle
    : makeGetCellStyle('theirs');
  const renderHeaderCell = useFullTables ? renderFullHeaderCell : renderDiffHeaderCell;
  const renderOursRowHeader = useFullTables ? renderFullRowHeader : makeRowHeaderRenderer('ours');
  const renderTheirsRowHeader = useFullTables ? renderFullRowHeader : makeRowHeaderRenderer('theirs');
  const oursRowsForGrid = useFullTables ? fullOursGrid : gridRows;
  const theirsRowsForGrid = useFullTables ? fullTheirsGrid : gridRows;

  // 当选择变化时，确保两侧都滚动到能看到该单元格
  const scrollToCell = useMemo(() => {
    if (!selected) return null;
    if (useFullTables) {
      return { rowIndex: selected.rowIndex, colIndex: selected.colIndex };
    }
    const targetRowNumber = selected.rowIndex + 1;
    const targetColNumber = selected.colIndex + 1;
    const gridRowIndex = diffRowNumbers.indexOf(targetRowNumber);
    const gridColIndex = displayColumns.indexOf(targetColNumber);
    if (gridRowIndex < 0 || gridColIndex < 0) return null;
    return { rowIndex: gridRowIndex, colIndex: gridColIndex };
  }, [selected, diffRowNumbers, displayColumns, useFullTables]);

  return (
    !useFullTables && !hasDiffData ? (
      <div>没有检测到任何差异。</div>
    ) : (
      <div
        style={{
          border: '1px solid #ccc',
          padding: 8,
          height: '100%',
          minHeight: 0,
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden',
          gap: 4,
        }}
      >
      <div
        style={{
          display: 'flex',
          gap: 16,
          fontSize: 12,
          color: '#444',
          alignItems: 'center',
          minHeight: 18,
        }}
      >
        <div style={{ flex: 1, minWidth: 0, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
          ours{oursPath ? `: ${oursPath}` : ''}
        </div>
        <div
          style={{
            maxWidth: '40%',
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            textAlign: 'center',
            flexShrink: 0,
          }}
        >
          base{basePath ? `: ${basePath}` : ''}
        </div>
        <div
          style={{
            flex: 1,
            minWidth: 0,
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            textAlign: 'right',
          }}
        >
          theirs{theirsPath ? `: ${theirsPath}` : ''}
        </div>
      </div>
      <div style={{ display: 'flex', gap: 16, flex: 1, minHeight: 0 }}>
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
          <VirtualGrid<any>
            rows={oursRowsForGrid}
            rowHeight={ROW_HEIGHT}
            overscanRows={OVERSCAN_ROWS}
            frozenRowCount={frozenRowCount}
            frozenColCount={0}
            rowHeaderWidth={120}
            showRowHeader
            renderRowHeader={renderOursRowHeader as any}
            onRowHeaderContextMenu={
              useFullTables ? undefined : (rowIndex, e) => handleRowHeaderContextMenu('ours', rowIndex, e)
            }
            renderCell={oursRenderCell}
            getCellStyle={oursGetCellStyle}
            renderHeaderCell={renderHeaderCell}
            defaultColWidth={DATA_COL_WIDTH}
            columnWidths={columnWidths}
            onColumnWidthsChange={setColumnWidths}
            containerRef={oursScrollRef as React.RefObject<HTMLDivElement>}
            onScrollXChange={(left) => syncScrollX('ours', left)}
            onScrollYChange={(top) => syncScrollY('ours', top)}
            scrollToCell={scrollToCell}
          />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
          <VirtualGrid<any>
            rows={theirsRowsForGrid}
            rowHeight={ROW_HEIGHT}
            overscanRows={OVERSCAN_ROWS}
            frozenRowCount={frozenRowCount}
            frozenColCount={0}
            rowHeaderWidth={120}
            showRowHeader
            renderRowHeader={renderTheirsRowHeader as any}
            onRowHeaderContextMenu={
              useFullTables ? undefined : (rowIndex, e) => handleRowHeaderContextMenu('theirs', rowIndex, e)
            }
            renderCell={theirsRenderCell}
            getCellStyle={theirsGetCellStyle}
            renderHeaderCell={renderHeaderCell}
            defaultColWidth={DATA_COL_WIDTH}
            columnWidths={columnWidths}
            onColumnWidthsChange={setColumnWidths}
            containerRef={theirsScrollRef as React.RefObject<HTMLDivElement>}
            onScrollXChange={(left) => syncScrollX('theirs', left)}
            onScrollYChange={(top) => syncScrollY('theirs', top)}
            scrollToCell={scrollToCell}
          />
        </div>
        {contextMenu && (
          <div
            style={{
              position: 'fixed',
              left: contextMenu.x,
              top: contextMenu.y,
              background: 'white',
              border: '1px solid #ccc',
              boxShadow: '0 2px 10px rgba(0,0,0,0.15)',
              zIndex: 9999,
              fontSize: 12,
              minWidth: 180,
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ padding: '6px 10px', borderBottom: '1px solid #eee', color: '#666' }}>
              {contextMenu.type === 'row'
                ? `行 ${contextMenu.rowNumber}`
                : `单元格 ${colNumberToLabel(contextMenu.colNumber)}${contextMenu.rowNumber}`}
              （来源：{contextMenu.source}）
            </div>

            {contextMenu.type === 'row' && (
              <button
                type="button"
                style={{ width: '100%', textAlign: 'left', padding: '6px 10px', border: 'none', background: 'white', cursor: 'pointer' }}
                onClick={() => {
                  if (onApplyRowChoice) onApplyRowChoice(contextMenu.rowNumber, contextMenu.source);
                  setContextMenu(null);
                }}
              >
                使用整行单元格数据
              </button>
            )}

            {contextMenu.type === 'cell' && (
              <button
                type="button"
                style={{ width: '100%', textAlign: 'left', padding: '6px 10px', border: 'none', background: 'white', cursor: 'pointer' }}
                onClick={() => {
                  if (onApplyCellChoice) {
                    onApplyCellChoice(contextMenu.rowNumber, contextMenu.colNumber, contextMenu.source);
                  }
                  setContextMenu(null);
                }}
              >
                使用本单元格的值
              </button>
            )}

            {contextMenu.type === 'cells' && (
              <button
                type="button"
                style={{ width: '100%', textAlign: 'left', padding: '6px 10px', border: 'none', background: 'white', cursor: 'pointer' }}
                onClick={() => {
                  if (onApplyCellsChoice) {
                    const keys = Array.from(selectedCellKeys.values()).map((k) => {
                      const [r, c] = k.split(':').map((x) => Number(x));
                      return { rowNumber: r, colNumber: c };
                    });
                    onApplyCellsChoice(keys, contextMenu.source);
                  }
                  setContextMenu(null);
                }}
              >
                使用选中单元格的数据
              </button>
            )}
          </div>
        )}
      </div>
    </div>
    )
  );
};

export const MergeSideBySide = React.memo(MergeSideBySideComponent);
