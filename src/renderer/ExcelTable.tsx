import React, { ChangeEvent, UIEvent, useEffect, useRef, useState } from 'react';
import type { SheetCell } from '../main/preload';

const ROW_HEIGHT = 24; // 估算单行高度，用于虚拟滚动
const OVERSCAN_ROWS = 10; // 视窗上下各多渲染几行，滚动更平滑

/**
 * 单文件编辑视图中使用的表格组件。
 *
 * 只渲染可见的行（基于简单的虚拟滚动实现），
 * 以便在行数较多时保持较好的滚动与编辑性能。
 */
interface ExcelTableProps {
  rows: SheetCell[][];
  onCellChange: (address: string, newValue: string) => void;
  /** 当用户点击/聚焦某个单元格时，通知上层当前选中的单元格 */
  onCellSelect?: (cell: SheetCell) => void;
  /** 当前选中单元格的地址，用于高亮显示 */
  selectedAddress?: string | null;
  /** 固定在顶部的行数（类似 Excel 冻结窗格），默认 0 表示不固定 */
  frozenRowCount?: number;
  /** 固定在左侧的列数（不包含最左侧行号列），默认 0 表示不固定 */
  frozenColCount?: number;
}

const ExcelTableComponent: React.FC<ExcelTableProps> = ({
  rows,
  onCellChange,
  onCellSelect,
  selectedAddress,
  frozenRowCount = 0,
  frozenColCount = 0,
}) => {
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const dragFrameRequestedRef = useRef(false);
  const lastClientXRef = useRef<number | null>(null);

  // 虚拟滚动：容器引用、滚动偏移和视口高度
  const containerRef = useRef<HTMLDivElement | null>(null);
  // 表头 + 冻结行所在的水平滚动容器，用来同步左右滚动位置
  const headerScrollRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(400);
  // 竖向滚动条宽度，用于让上方表头区域预留出同样的空间，避免列错位
  const [scrollbarWidth, setScrollbarWidth] = useState(0);

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
    const updateLayout = () => {
      if (containerRef.current) {
        const el = containerRef.current;
        const h = el.clientHeight;
        if (h > 0) {
          setViewportHeight(h);
        }
        const sw = el.offsetWidth - el.clientWidth;
        if (sw >= 0) {
          setScrollbarWidth(sw);
        }
      }
    };

    updateLayout();
    window.addEventListener('resize', updateLayout);
    return () => {
      window.removeEventListener('resize', updateLayout);
    };
  }, []);

  const handleChange = (cell: SheetCell) => (e: ChangeEvent<HTMLInputElement>) => {
    onCellChange(cell.address, e.target.value);
  };

  const handleSelect = (cell: SheetCell) => {
    if (onCellSelect) {
      onCellSelect(cell);
    }
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
    const target = e.currentTarget;
    setScrollTop(target.scrollTop);
    // 同步表头 + 冻结行的水平滚动位置
    if (headerScrollRef.current) {
      headerScrollRef.current.scrollLeft = target.scrollLeft;
    }
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
  const safeFrozenRowCount = Math.max(0, Math.min(frozenRowCount, totalRows));
  const frozenRows = rows.slice(0, safeFrozenRowCount);
  const scrollRows = rows.slice(safeFrozenRowCount);

  const totalScrollRows = scrollRows.length;
  const visibleRowCount = Math.ceil(viewportHeight / ROW_HEIGHT) + OVERSCAN_ROWS * 2;
  const firstVisibleRow = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - OVERSCAN_ROWS);
  const lastVisibleRow = Math.min(totalScrollRows, firstVisibleRow + visibleRowCount);
  const visibleRows = scrollRows.slice(firstVisibleRow, lastVisibleRow);
  const topSpacerHeight = firstVisibleRow * ROW_HEIGHT;
  const bottomSpacerHeight = (totalScrollRows - lastVisibleRow) * ROW_HEIGHT;

  // 冻结列（不含最左侧行号列）的位置偏移，用于 sticky left
  const safeFrozenColCount = Math.max(0, Math.min(frozenColCount, colCount));
  const frozenColOffsets: number[] = [];
  {
    let acc = 40; // 行号列固定宽度
    for (let colIndex = 0; colIndex < safeFrozenColCount; colIndex += 1) {
      frozenColOffsets[colIndex] = acc;
      acc += getColWidth(colIndex);
    }
  }

  const renderRow = (row: SheetCell[], rowIndex: number) => {
    const excelRowNumber = row[0]?.row ?? rowIndex + 1;
    const isFrozenRow = rowIndex < safeFrozenRowCount;
    return (
      <tr key={rowIndex}>
        <td
          style={{
            position: 'sticky',
            left: 0,
            zIndex: 4,
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
        {Array.from({ length: colCount }, (_, colIndex) => {
          const cell = row[colIndex];
          const isFrozenCol = colIndex < safeFrozenColCount;
          const isFrozenCell = isFrozenRow || isFrozenCol;
          const stickyLeft = isFrozenCol ? frozenColOffsets[colIndex] ?? 40 : undefined;

          // 行中没有这个 SheetCell：渲染一个纯空白单元格，只负责显示网格
          if (!cell) {
            return (
              <td
                key={`empty-${rowIndex}-${colIndex}`}
                style={{
                  position: isFrozenCol ? 'sticky' : 'static',
                  left: stickyLeft,
                  zIndex: isFrozenCol ? 3 : 1,
                  border: '1px solid #ddd',
                  padding: 2,
                  width: getColWidth(colIndex),
                  minWidth: getColWidth(colIndex),
                  height: ROW_HEIGHT,
                  backgroundColor: isFrozenCell ? '#f5f5f5' : '#fff',
                }}
              />
            );
          }

          const isSelected = selectedAddress === cell.address;
          return (
            <td
              key={cell.address}
              style={{
                position: isFrozenCol ? 'sticky' : 'static',
                left: stickyLeft,
                zIndex: isFrozenCol ? 3 : 1,
                border: isSelected ? '2px solid #00aa00' : '1px solid #ddd',
                padding: 2,
                width: getColWidth(colIndex),
                minWidth: getColWidth(colIndex),
                height: ROW_HEIGHT,
                backgroundColor: isFrozenCell ? '#f5f5f5' : undefined,
              }}
              onClick={() => handleSelect(cell)}
            >
              <input
                onFocus={() => handleSelect(cell)}
                style={{
                  width: '100%',
                  boxSizing: 'border-box',
                  border: 'none',
                  outline: 'none',
                  backgroundColor: 'transparent',
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                }}
                value={cell.value === null ? '' : String(cell.value)}
                onChange={handleChange(cell)}
              />
            </td>
          );
        })}
      </tr>
    );
  };

  return (
    <div
      style={{
        border: '1px solid #ccc',
        height: '100%',
        display: 'flex',
        flexDirection: 'column',
        overflow: 'hidden',
      }}
    >
      {/* 表头 + 冻结行：不随滚动条上下滚动；通过 scrollLeft 与下方可滚动区域保持水平同步 */}
      <div
        ref={headerScrollRef}
        style={{ overflowX: 'hidden', overflowY: 'hidden', paddingRight: scrollbarWidth }}
      >
        <table style={{ borderCollapse: 'collapse', width: 'max-content' }}>
          <thead>
          <tr>
            <th
              style={{
                position: 'sticky',
                left: 0,
                zIndex: 5,
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
            {columnLabels.map((label, colIndex) => {
              const isFrozenCol = colIndex < safeFrozenColCount;
              const stickyLeft = isFrozenCol ? frozenColOffsets[colIndex] ?? 40 : undefined;
              return (
                <th
                  key={colIndex}
                  style={{
                    position: isFrozenCol ? 'sticky' : 'relative',
                    left: stickyLeft,
                    zIndex: isFrozenCol ? 4 : 1,
                    border: '1px solid #ddd',
                    padding: 2,
                    textAlign: 'center',
                    userSelect: 'none',
                    width: getColWidth(colIndex),
                    minWidth: getColWidth(colIndex),
                    backgroundColor: isFrozenCol ? '#ccd8ff' : '#f0f0f0',
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
            );
            })}
          </tr>
        </thead>
        <tbody>
          {frozenRows.map((row, rowIndex) => renderRow(row, rowIndex))}
        </tbody>
      </table>
      </div>

      {/* 可滚动部分：不包含已冻结的前几行 */}
      <div
        ref={containerRef}
        onScroll={handleScroll}
        style={{ flex: 1, minHeight: 0, overflowY: 'auto', overflowX: 'auto' }}
      >
        <table style={{ borderCollapse: 'collapse', width: 'max-content' }}>
          <tbody>
            {topSpacerHeight > 0 && (
              <tr style={{ height: topSpacerHeight }}>
                <td colSpan={colCount + 1} />
              </tr>
            )}
            {visibleRows.map((row, localRowIndex) => {
              const globalRowIndex = safeFrozenRowCount + firstVisibleRow + localRowIndex;
              return renderRow(row, globalRowIndex);
            })}
            {bottomSpacerHeight > 0 && (
              <tr style={{ height: bottomSpacerHeight }}>
                <td colSpan={colCount + 1} />
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export const ExcelTable = React.memo(ExcelTableComponent);
