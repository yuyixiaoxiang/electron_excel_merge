import React, { CSSProperties, UIEvent, useEffect, useMemo, useRef, useState } from 'react';

// 通用虚拟表格组件：只负责滚动、冻结行列、行号列和列宽拖拽，不关心具体单元格内容

export interface VirtualGridRenderCtx {
  rowIndex: number; // 全局行下标（0-based）
  colIndex: number;
  isFrozenRow: boolean;
  isFrozenCol: boolean;
}

export interface VirtualGridProps<Cell> {
  rows: Cell[][];
  rowHeight?: number;
  overscanRows?: number;
  frozenRowCount?: number;
  frozenColCount?: number; // 不包含最左侧行号列
  // 行号列
  showRowHeader?: boolean;
  renderRowHeader?: (rowIndex: number, row: Cell[]) => React.ReactNode;
  // 单元格内容
  renderCell: (cell: Cell | null, ctx: VirtualGridRenderCtx) => React.ReactNode;
  // 单元格样式（背景色、边框等，返回的 style 会 merge 到内部样式之后）
  getCellStyle?: (cell: Cell | null, ctx: VirtualGridRenderCtx) => CSSProperties | undefined;
  // 列头
  renderHeaderCell?: (colIndex: number) => React.ReactNode;
  // 初始列宽（像素），如果没提供则用 120
  defaultColWidth?: number;
  // 可选：暴露内部滚动容器，便于外部同步 scrollLeft
  containerRef?: React.RefObject<HTMLDivElement>;
  // 可选：每次水平滚动时通知外部当前 scrollLeft，用于左右表格联动
  onScrollXChange?: (scrollLeft: number) => void;
}

const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_OVERSCAN_ROWS = 10;
const MIN_COL_WIDTH = 40;

export function VirtualGrid<Cell>(props: VirtualGridProps<Cell>) {
  const {
    rows,
    rowHeight = DEFAULT_ROW_HEIGHT,
    overscanRows = DEFAULT_OVERSCAN_ROWS,
    frozenRowCount = 0,
    frozenColCount = 0,
    showRowHeader = true,
    renderRowHeader,
    renderCell,
    getCellStyle,
    renderHeaderCell,
    defaultColWidth = 120,
  } = props;

  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const dragFrameRequestedRef = useRef(false);
  const lastClientXRef = useRef<number | null>(null);

  const internalContainerRef = useRef<HTMLDivElement | null>(null);
  const containerRef = props.containerRef ?? internalContainerRef;
  const headerScrollRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(400);
  const [scrollbarWidth, setScrollbarWidth] = useState(0);

  if (rows.length === 0) {
    return null;
  }

  const colCount = rows[0].length;

  // 初始化列宽
  useEffect(() => {
    if (rows.length === 0) return;
    const count = rows[0].length;
    setColumnWidths((prev) => {
      if (prev.length === count) return prev;
      return Array(count).fill(defaultColWidth);
    });
  }, [rows, defaultColWidth]);

  // 初始化和更新视口高度 + 竖向滚动条宽度
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

  const getColWidth = (colIndex: number) => columnWidths[colIndex] ?? defaultColWidth;

  const handleScroll = (e: UIEvent<HTMLDivElement>) => {
    const target = e.currentTarget;
    setScrollTop(target.scrollTop);
    if (headerScrollRef.current) {
      headerScrollRef.current.scrollLeft = target.scrollLeft;
    }
    if (props.onScrollXChange) {
      props.onScrollXChange(target.scrollLeft);
    }
  };

  const handleMouseDownOnResizer = (
    e: React.MouseEvent<HTMLDivElement>,
    colIndex: number,
  ) => {
    e.preventDefault();
    e.stopPropagation();

    const startX = e.clientX;
    const startWidth = getColWidth(colIndex);

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
          const newWidth = Math.max(MIN_COL_WIDTH, startWidth + delta);
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

  // 虚拟滚动窗口
  const totalRows = rows.length;
  const safeFrozenRowCount = Math.max(0, Math.min(frozenRowCount, totalRows));
  const frozenRows = rows.slice(0, safeFrozenRowCount);
  const scrollRows = rows.slice(safeFrozenRowCount);

  const totalScrollRows = scrollRows.length;
  const visibleRowCount = Math.ceil(viewportHeight / rowHeight) + overscanRows * 2;
  const firstVisibleRow = Math.max(0, Math.floor(scrollTop / rowHeight) - overscanRows);
  const lastVisibleRow = Math.min(totalScrollRows, firstVisibleRow + visibleRowCount);
  const visibleRows = scrollRows.slice(firstVisibleRow, lastVisibleRow);
  const topSpacerHeight = firstVisibleRow * rowHeight;
  const bottomSpacerHeight = (totalScrollRows - lastVisibleRow) * rowHeight;

  // 冻结列（不含最左侧行号列）偏移
  const safeFrozenColCount = Math.max(0, Math.min(frozenColCount, colCount));
  const frozenColOffsets: number[] = useMemo(() => {
    const offsets: number[] = [];
    let acc = showRowHeader ? 40 : 0;
    for (let colIndex = 0; colIndex < safeFrozenColCount; colIndex += 1) {
      offsets[colIndex] = acc;
      acc += getColWidth(colIndex);
    }
    return offsets;
  }, [safeFrozenColCount, columnWidths, showRowHeader]);

  const renderRow = (row: Cell[], rowIndex: number) => {
    const isFrozenRow = rowIndex < safeFrozenRowCount;
    const ctxBase = { isFrozenRow };

    const rowHeader = showRowHeader ? (
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
        {renderRowHeader ? renderRowHeader(rowIndex, row) : rowIndex + 1}
      </td>
    ) : null;

    return (
      <tr key={rowIndex}>
        {rowHeader}
        {Array.from({ length: colCount }, (_, colIndex) => {
          const cell = row[colIndex] ?? null;
          const isFrozenCol = colIndex < safeFrozenColCount;
          const isFrozenCell = isFrozenRow || isFrozenCol;
          const stickyLeft = isFrozenCol ? frozenColOffsets[colIndex] ?? (showRowHeader ? 40 : 0) : undefined;

          const ctx: VirtualGridRenderCtx = {
            rowIndex,
            colIndex,
            isFrozenRow,
            isFrozenCol,
          };

          const externalStyle = getCellStyle ? getCellStyle(cell, ctx) : undefined;

          return (
            <td
              key={colIndex}
              style={{
                position: isFrozenCol ? 'sticky' : 'static',
                left: stickyLeft,
                zIndex: isFrozenCol ? 3 : 1,
                border: '1px solid #ddd',
                padding: 2,
                width: getColWidth(colIndex),
                minWidth: getColWidth(colIndex),
                height: rowHeight,
                ...(externalStyle || {}),
              }}
            >
              {renderCell(cell, ctx)}
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
              {showRowHeader && (
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
              )}
              {Array.from({ length: colCount }, (_, colIndex) => {
                const isFrozenCol = colIndex < safeFrozenColCount;
                const stickyLeft = isFrozenCol ? frozenColOffsets[colIndex] ?? (showRowHeader ? 40 : 0) : undefined;
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
                      backgroundColor: '#f0f0f0',
                    }}
                  >
                    {renderHeaderCell ? renderHeaderCell(colIndex) : String.fromCharCode('A'.charCodeAt(0) + (colIndex % 26))}
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
                <td colSpan={(showRowHeader ? 1 : 0) + colCount} />
              </tr>
            )}
            {visibleRows.map((row, localRowIndex) => {
              const globalRowIndex = safeFrozenRowCount + firstVisibleRow + localRowIndex;
              return renderRow(row, globalRowIndex);
            })}
            {bottomSpacerHeight > 0 && (
              <tr style={{ height: bottomSpacerHeight }}>
                <td colSpan={(showRowHeader ? 1 : 0) + colCount} />
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
