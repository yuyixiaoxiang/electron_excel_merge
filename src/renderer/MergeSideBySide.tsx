import React, { useEffect, useMemo, useRef, useState } from 'react';
import type { MergeCell } from '../main/preload';
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
  selected?: { rowIndex: number; colIndex: number } | null;
  onSelectCell?: (rowIndex: number, colIndex: number) => void;
  /** 冻结在顶部展示的行数，可配置 */
  frozenRowCount?: number;
}

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
      return '#fff6bf'; // both sides yellow
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
  selected,
  onSelectCell,
  frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS,
}) => {
  if (cells.length === 0) return <div>没有检测到任何差异。</div>;

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

  // 差异行号（1-based Excel 行号）
  const diffRowNumbers = useMemo(() => {
    const rs = new Set<number>();
    cells.forEach((cell) => rs.add(cell.row));
    return Array.from(rs).sort((a, b) => a - b);
  }, [cells]);

  if (diffColumns.length === 0 || diffRowNumbers.length === 0) {
    return <div>没有检测到任何差异。</div>;
  }

  // 将原始 rows + diffColumns/diffRowNumbers 转成 VirtualGrid 需要的矩阵
  const gridRowNumbers = useMemo(() => diffRowNumbers, [diffRowNumbers]);

  const gridRows: (MergeCell | null)[][] = useMemo(
    () =>
      gridRowNumbers.map((rowNumber) =>
        diffColumns.map((colNumber) => cellMap.get(`${rowNumber}:${colNumber}`) ?? null),
      ),
    [cellMap, gridRowNumbers, diffColumns],
  );

  // 两侧共享列宽，避免左右/表头/内容出现 1px 累积偏差或拖拽后不同步
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  useEffect(() => {
    const count = diffColumns.length;
    setColumnWidths((prev) => {
      if (prev.length === count) return prev;
      return Array(count).fill(DATA_COL_WIDTH);
    });
  }, [diffColumns.length]);

  const renderRowHeader = (_gridRowIndex: number, rowCells: (MergeCell | null)[]) => {
    const anyCell = rowCells.find((c) => c != null) ?? undefined;
    return anyCell?.row ?? '';
  };

  const makeRenderCell = (side: 'ours' | 'theirs') =>
    (cell: MergeCell | null, ctx: VirtualGridRenderCtx) => {
      if (!cell) return null;
      const value = side === 'ours' ? cell.oursValue : cell.theirsValue;
      const sourceRowIndex = cell.row - 1;
      const sourceColIndex = cell.col - 1;
      const handleClick = () => {
        if (onSelectCell) {
          onSelectCell(sourceRowIndex, sourceColIndex);
        }
      };

      return (
        <div
          onClick={handleClick}
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
          }}
        >
          {value === null ? '' : String(value)}
        </div>
      );
    };

  const makeGetCellStyle = (side: 'ours' | 'theirs') =>
    (cell: MergeCell | null, ctx: VirtualGridRenderCtx): React.CSSProperties => {
      const style: React.CSSProperties = {};
      if (ctx.isFrozenRow) {
        style.backgroundColor = '#f5f5f5';
      } else if (cell) {
        style.backgroundColor = getBackgroundColor(cell.status, side);
      }

      if (cell && selected) {
        const sourceRowIndex = cell.row - 1;
        const sourceColIndex = cell.col - 1;
        const isSelected =
          selected.rowIndex === sourceRowIndex && selected.colIndex === sourceColIndex;
        if (isSelected) {
          style.border = '2px solid #ff8000';
        }
      }

      return style;
    };

  const renderHeaderCell = (colIndex: number) => {
    const colNumber = diffColumns[colIndex];
    return colNumberToLabel(colNumber);
  };

  const oursRenderCell = makeRenderCell('ours');
  const theirsRenderCell = makeRenderCell('theirs');
  const oursGetCellStyle = makeGetCellStyle('ours');
  const theirsGetCellStyle = makeGetCellStyle('theirs');

  // 当选择变化时，确保两侧都滚动到能看到该单元格
  const scrollToCell = useMemo(() => {
    if (!selected) return null;
    const targetRowNumber = selected.rowIndex + 1;
    const targetColNumber = selected.colIndex + 1;
    const gridRowIndex = diffRowNumbers.indexOf(targetRowNumber);
    const gridColIndex = diffColumns.indexOf(targetColNumber);
    if (gridRowIndex < 0 || gridColIndex < 0) return null;
    return { rowIndex: gridRowIndex, colIndex: gridColIndex };
  }, [selected, diffRowNumbers, diffColumns]);

  return (
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
      <div style={{ display: 'flex', gap: 16, flex: 1, minHeight: 0 }}>
        <VirtualGrid<MergeCell | null>
          rows={gridRows}
          rowHeight={ROW_HEIGHT}
          overscanRows={OVERSCAN_ROWS}
          frozenRowCount={frozenRowCount}
          frozenColCount={0}
          showRowHeader
          renderRowHeader={renderRowHeader}
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
        <VirtualGrid<MergeCell | null>
          rows={gridRows}
          rowHeight={ROW_HEIGHT}
          overscanRows={OVERSCAN_ROWS}
          frozenRowCount={frozenRowCount}
          frozenColCount={0}
          showRowHeader
          renderRowHeader={renderRowHeader}
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
    </div>
  );
};

export const MergeSideBySide = React.memo(MergeSideBySideComponent);
