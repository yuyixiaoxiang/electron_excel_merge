import React, { ChangeEvent, UIEvent, useEffect, useRef, useState } from 'react';
import type { SheetCell } from '../main/preload';

const ROW_HEIGHT = 24; // 估算单行高度，用于虚拟滚动
const OVERSCAN_ROWS = 10; // 视窗上下各多渲染几行，滚动更平滑

interface ExcelTableProps {
  rows: SheetCell[][];
  onCellChange: (address: string, newValue: string) => void;
}

const ExcelTableComponent: React.FC<ExcelTableProps> = ({ rows, onCellChange }) => {
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const dragFrameRequestedRef = useRef(false);
  const lastClientXRef = useRef<number | null>(null);

  // 虚拟滚动：容器引用、滚动偏移和视口高度
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(400);

  // 初始化列宽（根据首行列数）
  useEffect(() => {
    if (rows.length === 0) return;
    const colCount = rows[0].length;
    setColumnWidths((prev) => {
      if (prev.length === colCount) return prev;
      // 默认每列 120 像素宽
      return Array(colCount).fill(120);
    });
  }, [rows]);

  // 初始化和更新视口高度，用于计算需要渲染的行数
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

  const handleChange = (cell: SheetCell) => (e: ChangeEvent<HTMLInputElement>) => {
    onCellChange(cell.address, e.target.value);
  };

  const handleMouseDownOnResizer = (
    e: React.MouseEvent<HTMLDivElement>,
    colIndex: number,
  ) => {
    e.preventDefault();
    e.stopPropagation();

    const startX = e.clientX;
    const startWidth = columnWidths[colIndex] ?? 120;

    const handleMouseMove = (moveEvent: MouseEvent) => {
      lastClientXRef.current = moveEvent.clientX;

      if (dragFrameRequestedRef.current) {
        return;
      }

      dragFrameRequestedRef.current = true;

      requestAnimationFrame(() => {
        dragFrameRequestedRef.current = false;
        const clientX = lastClientXRef.current;
        if (clientX == null) return;

        setColumnWidths((prev) => {
          const next = [...prev];
          const delta = clientX - startX;
          const newWidth = Math.max(40, startWidth + delta); // 最小 40px
          next[colIndex] = newWidth;
          return next;
        });
      });
    };

    const handleMouseUp = () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  };

  const getColWidth = (colIndex: number) => columnWidths[colIndex] ?? 120;

  const handleScroll = (e: UIEvent<HTMLDivElement>) => {
    setScrollTop(e.currentTarget.scrollTop);
  };

  if (rows.length === 0) {
    return null;
  }

  const colCount = rows[0].length;
  // 简单列头 A,B,C... 主要用来放拖拽条
  const columnLabels = Array.from({ length: colCount }, (_, i) =>
    String.fromCharCode('A'.charCodeAt(0) + (i % 26)),
  );

  // 计算需要渲染的行窗口（虚拟滚动）
  const totalRows = rows.length;
  const visibleRowCount = Math.ceil(viewportHeight / ROW_HEIGHT) + OVERSCAN_ROWS * 2;
  const firstVisibleRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN_ROWS);
  const lastVisibleRow = Math.min(totalRows, firstVisibleRow + visibleRowCount);
  const visibleRows = rows.slice(firstVisibleRow, lastVisibleRow);
  const topSpacerHeight = firstVisibleRow * ROW_HEIGHT;
  const bottomSpacerHeight = (totalRows - lastVisibleRow) * ROW_HEIGHT;

  return (
    <div
      ref={containerRef}
      onScroll={handleScroll}
      style={{ overflow: 'auto', maxHeight: '70vh', border: '1px solid #ccc' }}
    >
      <table style={{ borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            <th
              style={{
                border: '1px solid #ddd',
                padding: 2,
                textAlign: 'right',
                userSelect: 'none',
                width: 40,
                minWidth: 40,
                backgroundColor: '#f0f0f0',
                fontSize: 12,
              }}
            >
              行
            </th>
            {columnLabels.map((label, colIndex) => (
              <th
                key={colIndex}
                style={{
                  position: 'relative',
                  border: '1px solid #ddd',
                  padding: 2,
                  textAlign: 'center',
                  userSelect: 'none',
                  width: getColWidth(colIndex),
                  minWidth: getColWidth(colIndex),
                }}
              >
                {label}
                <div
                  onMouseDown={(e) => handleMouseDownOnResizer(e, colIndex)}
                  style={{
                    position: 'absolute',
                    right: 0,
                    top: 0,
                    bottom: 0,
                    width: 4,
                    cursor: 'col-resize',
                    backgroundColor: 'transparent',
                  }}
                />
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {topSpacerHeight > 0 && (
            <tr style={{ height: topSpacerHeight }}>
              <td colSpan={colCount + 1} />
            </tr>
          )}
          {visibleRows.map((row, localRowIndex) => {
            const rowIndex = firstVisibleRow + localRowIndex;
            const excelRowNumber = row[0]?.row ?? rowIndex + 1;
            return (
              <tr key={rowIndex}>
                <td
                  style={{
                    border: '1px solid #ddd',
                    padding: 2,
                    textAlign: 'right',
                    userSelect: 'none',
                    width: 40,
                    minWidth: 40,
                    backgroundColor: '#f7f7f7',
                    fontSize: 12,
                  }}
                >
                  {excelRowNumber}
                </td>
              {row.map((cell, colIndex) => (
                <td
                  key={cell.address}
                  style={{
                    border: '1px solid #ddd',
                    padding: 2,
                    width: getColWidth(colIndex),
                    minWidth: getColWidth(colIndex),
                    height: ROW_HEIGHT,
                  }}
                >
                  <input
                    style={{
                      width: '100%',
                      boxSizing: 'border-box',
                      border: 'none',
                      outline: 'none',
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                    }}
                    value={cell.value === null ? '' : String(cell.value)}
                    onChange={handleChange(cell)}
                  />
                </td>
              ))}
            </tr>
          );
        })}
          {bottomSpacerHeight > 0 && (
            <tr style={{ height: bottomSpacerHeight }}>
              <td colSpan={colCount + 1} />
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
};

export const ExcelTable = React.memo(ExcelTableComponent);
