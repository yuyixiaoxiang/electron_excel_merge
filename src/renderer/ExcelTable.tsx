import React, { ChangeEvent, useEffect, useRef, useState } from 'react';
import type { SheetCell } from '../main/preload';

interface ExcelTableProps {
  rows: SheetCell[][];
  onCellChange: (address: string, newValue: string) => void;
}

export const ExcelTable: React.FC<ExcelTableProps> = ({ rows, onCellChange }) => {
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const dragFrameRequestedRef = useRef(false);
  const lastClientXRef = useRef<number | null>(null);

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

  if (rows.length === 0) {
    return null;
  }

  const colCount = rows[0].length;
  // 简单列头 A,B,C... 主要用来放拖拽条
  const columnLabels = Array.from({ length: colCount }, (_, i) =>
    String.fromCharCode('A'.charCodeAt(0) + (i % 26)),
  );

  return (
    <div style={{ overflow: 'auto', maxHeight: '70vh', border: '1px solid #ccc' }}>
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
          {rows.map((row, rowIndex) => {
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
        </tbody>
      </table>
    </div>
  );
};
