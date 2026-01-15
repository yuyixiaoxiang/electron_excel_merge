import React, { UIEvent, useEffect, useMemo, useRef, useState } from 'react';
import type { MergeCell } from '../main/preload';

const ROW_HEIGHT = 24; // px, approximate row height for virtualization
const OVERSCAN_ROWS = 8; // render a few extra rows above/below viewport for smooth scroll

export interface MergeSideBySideProps {
  rows: MergeCell[][];
  selected?: { rowIndex: number; colIndex: number } | null;
  onSelectCell?: (rowIndex: number, colIndex: number) => void;
}

const getBackgroundColor = (status: MergeCell['status'], side: 'ours' | 'theirs'): string => {
  switch (status) {
    case 'unchanged':
      return 'white';
    case 'ours-changed':
      return side === 'ours' ? '#d4f8d4' : 'white';
    case 'theirs-changed':
      return side === 'theirs' ? '#d4e8ff' : 'white';
    case 'both-changed-same':
      return '#fff6bf'; // both sides yellow
    case 'conflict':
      return '#ffc8c8'; // both sides red
    default:
      return 'white';
  }
};

const getColumnLabels = (cols: number[]): string[] => {
  // 按列索引生成简单的 A,B,C... 标记，仅用于显示
  return cols.map((_, i) => String.fromCharCode('A'.charCodeAt(0) + (i % 26)));
};

const DATA_COL_WIDTH = 160; // px, keep ours/theirs columns visually aligned

const MergeSideBySideComponent: React.FC<MergeSideBySideProps> = ({
  rows,
  selected,
  onSelectCell,
}) => {
  if (rows.length === 0) return null;

  // 外层滚动容器引用 & 虚拟滚动状态
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(400);

  // 水平滚动同步：左右表格各自有横向滚动条，但滚动位置保持一致
  const oursTableRef = useRef<HTMLDivElement | null>(null);
  const theirsTableRef = useRef<HTMLDivElement | null>(null);
  const isSyncingHorizontalRef = useRef(false);

  useEffect(() => {
    const updateViewportHeight = () => {
      if (containerRef.current) {
        const h = containerRef.current.clientHeight;
        if (h > 0) {
          setViewportHeight(h);
        }
      }
    };

    updateViewportHeight();
    window.addEventListener('resize', updateViewportHeight);
    return () => {
      window.removeEventListener('resize', updateViewportHeight);
    };
  }, []);

  const handleScroll = (e: UIEvent<HTMLDivElement>) => {
    setScrollTop(e.currentTarget.scrollTop);
  };

  const handleHorizontalScroll = (side: 'ours' | 'theirs') => (e: UIEvent<HTMLDivElement>) => {
    const current = e.currentTarget;
    const other = side === 'ours' ? theirsTableRef.current : oursTableRef.current;
    if (!other) return;
    if (isSyncingHorizontalRef.current) return;

    isSyncingHorizontalRef.current = true;
    other.scrollLeft = current.scrollLeft;
    requestAnimationFrame(() => {
      isSyncingHorizontalRef.current = false;
    });
  };

  // 只展示有差异的行/列（status !== 'unchanged'）
  const diffColumns = useMemo(() => {
    const cols = new Set<number>();
    rows.forEach((row) => {
      row.forEach((cell) => {
        if (cell.status !== 'unchanged') {
          cols.add(cell.col);
        }
      });
    });
    return Array.from(cols).sort((a, b) => a - b);
  }, [rows]);

  const diffRowNumbers = useMemo(() => {
    const rowsWithDiff = new Set<number>();
    rows.forEach((row) => {
      if (row.some((cell) => cell.status !== 'unchanged')) {
        // MergeCell.row 从 1 开始，与 main.ts 中一致
        if (row.length > 0) {
          rowsWithDiff.add(row[0].row);
        }
      }
    });
    return Array.from(rowsWithDiff).sort((a, b) => a - b);
  }, [rows]);

  if (diffColumns.length === 0 || diffRowNumbers.length === 0) {
    return <div>没有检测到任何差异。</div>;
  }

  const columnLabels = getColumnLabels(diffColumns);

  // 按 diff 行号做虚拟滚动（两侧共享同一套可见行）
  const totalRows = diffRowNumbers.length;
  const visibleRowCount = Math.ceil(viewportHeight / ROW_HEIGHT) + OVERSCAN_ROWS * 2;
  const firstVisibleIndex = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN_ROWS);
  const lastVisibleIndex = Math.min(totalRows, firstVisibleIndex + visibleRowCount);
  const visibleRowNumbers = diffRowNumbers.slice(firstVisibleIndex, lastVisibleIndex);
  const topSpacerHeight = firstVisibleIndex * ROW_HEIGHT;
  const bottomSpacerHeight = (totalRows - lastVisibleIndex) * ROW_HEIGHT;

  const renderTable = (side: 'ours' | 'theirs') => (
    <div
      ref={side === 'ours' ? oursTableRef : theirsTableRef}
      onScroll={handleHorizontalScroll(side)}
      style={{ flex: 1, overflowX: 'auto', overflowY: 'hidden' }}
    >
      <div style={{ marginBottom: 4, fontWeight: 'bold', fontSize: 12 }}>
        {side === 'ours' ? 'ours (当前分支)' : 'theirs (合并分支)'}
      </div>
      <table style={{ borderCollapse: 'collapse', width: '100%' }}>
        <thead>
          <tr>
            <th
              style={{
                border: '1px solid #ddd',
                padding: 2,
                textAlign: 'right',
                userSelect: 'none',
                backgroundColor: '#f0f0f0',
                fontSize: 12,
                width: 40,
                minWidth: 40,
              }}
            >
              行
            </th>
            {columnLabels.map((label, colIndex) => (
              <th
                key={colIndex}
                style={{
                  border: '1px solid #ddd',
                  padding: 2,
                  textAlign: 'center',
                  userSelect: 'none',
                  backgroundColor: '#f0f0f0',
                  fontSize: 12,
                  width: DATA_COL_WIDTH,
                  minWidth: DATA_COL_WIDTH,
                }}
              >
                {label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {topSpacerHeight > 0 && (
            <tr style={{ height: topSpacerHeight }}>
              <td colSpan={columnLabels.length + 1} />
            </tr>
          )}
          {visibleRowNumbers.map((rowNumber) => {
            const row = rows[rowNumber - 1];
            return (
              <tr key={rowNumber} style={{ height: ROW_HEIGHT }}>
                <td
                  style={{
                    border: '1px solid #ddd',
                    padding: 2,
                    textAlign: 'right',
                    userSelect: 'none',
                    backgroundColor: '#f7f7f7',
                    fontSize: 12,
                    width: 40,
                    minWidth: 40,
                  }}
                >
                  {rowNumber}
                </td>
                {diffColumns.map((colNumber) => {
                  const cell = row[colNumber - 1];
                  if (!cell) {
                    // 保持与有内容的单元格相同的宽度，避免一侧为空时列变窄
                    return (
                      <td
                        key={`${rowNumber}-${colNumber}`}
                        style={{
                          border: '1px solid #ddd',
                          padding: 2,
                          fontSize: 12,
                          width: DATA_COL_WIDTH,
                          minWidth: DATA_COL_WIDTH,
                        }}
                      />
                    );
                  }

                  const sourceRowIndex = cell.row - 1;
                  const sourceColIndex = cell.col - 1;

                  const isSelected =
                    selected &&
                    selected.rowIndex === sourceRowIndex &&
                    selected.colIndex === sourceColIndex;

                  const value =
                    side === 'ours'
                      ? cell.oursValue
                      : cell.theirsValue;

                  return (
                    <td
                      key={cell.address}
                      style={{
                        border: isSelected ? '2px solid #ff8000' : '1px solid #ddd',
                        padding: 2,
                        backgroundColor: getBackgroundColor(cell.status, side),
                        fontSize: 12,
                        width: DATA_COL_WIDTH,
                        minWidth: DATA_COL_WIDTH,
                      }}
                      title={`地址: ${cell.address}\nbase: ${cell.baseValue ?? ''}\nours: ${cell.oursValue ?? ''}\ntheirs: ${cell.theirsValue ?? ''}`}
                      onClick={() => onSelectCell && onSelectCell(sourceRowIndex, sourceColIndex)}
                    >
                      {value === null ? '' : String(value)}
                    </td>
                  );
                })}
              </tr>
            );
          })}
          {bottomSpacerHeight > 0 && (
            <tr style={{ height: bottomSpacerHeight }}>
              <td colSpan={columnLabels.length + 1} />
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );

  return (
    <div
      ref={containerRef}
      onScroll={handleScroll}
      style={{
        display: 'flex',
        gap: 16,
        border: '1px solid #ccc',
        padding: 8,
        maxHeight: '70vh',
        overflowY: 'auto',
        overflowX: 'hidden',
      }}
    >
      {renderTable('ours')}
      {renderTable('theirs')}
    </div>
  );
};

export const MergeSideBySide = React.memo(MergeSideBySideComponent);
