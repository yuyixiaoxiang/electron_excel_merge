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
  rowHeaderWidth?: number; // 行号列宽度
  // 行号列
  showRowHeader?: boolean;
  renderRowHeader?: (rowIndex: number, row: Cell[]) => React.ReactNode;
  onRowHeaderContextMenu?: (rowIndex: number, e: React.MouseEvent<HTMLTableCellElement>) => void;
  // 单元格内容
  renderCell: (cell: Cell | null, ctx: VirtualGridRenderCtx) => React.ReactNode;
  // 单元格样式（背景色、边框等，返回的 style 会 merge 到内部样式之后）
  getCellStyle?: (cell: Cell | null, ctx: VirtualGridRenderCtx) => CSSProperties | undefined;
  // 列头
  renderHeaderCell?: (colIndex: number) => React.ReactNode;
  onHeaderContextMenu?: (colIndex: number, e: React.MouseEvent<HTMLTableCellElement>) => void;
  // 初始列宽（像素），如果没提供则用 120
  defaultColWidth?: number;
  // 可选：外部受控列宽（用于左右表格列宽一致）
  columnWidths?: number[];
  onColumnWidthsChange?: (widths: number[]) => void;
  // 可选：外部受控横向滚动位置（用于共享一个横向滚动条）
  scrollLeft?: number | null;
  disableHorizontalScroll?: boolean;
  // 可选：暴露内部滚动容器，便于外部同步 scrollLeft/scrollTop
  containerRef?: React.RefObject<HTMLDivElement>;
  // 可选：每次水平滚动时通知外部当前 scrollLeft，用于左右表格联动
  onScrollXChange?: (scrollLeft: number) => void;
  // 可选：每次竖向滚动时通知外部当前 scrollTop，用于左右表格联动
  onScrollYChange?: (scrollTop: number) => void;
  // 可选：滚动使指定单元格出现在视口中（rowIndex/colIndex 均为 0-based grid 索引）
  scrollToCell?: { rowIndex: number; colIndex: number } | null;
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
    rowHeaderWidth = 40,
    showRowHeader = true,
    renderRowHeader,
    onRowHeaderContextMenu,
    renderCell,
    getCellStyle,
    renderHeaderCell,
    onHeaderContextMenu,
    defaultColWidth = 120,
  } = props;

  const [internalColumnWidths, setInternalColumnWidths] = useState<number[]>([]);
  const isControlledWidths = Array.isArray(props.columnWidths);
  const columnWidths = (props.columnWidths ?? internalColumnWidths) as number[];
  const dragFrameRequestedRef = useRef(false);
  const lastClientXRef = useRef<number | null>(null);

  const internalContainerRef = useRef<HTMLDivElement | null>(null);
  const containerRef = props.containerRef ?? internalContainerRef;
  const headerScrollRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const lastScrollTopRef = useRef(0);
  const scrollRafRequestedRef = useRef(false);
  const lastScrollLeftRef = useRef(0);
  const [viewportHeight, setViewportHeight] = useState(400);
  const [scrollbarWidth, setScrollbarWidth] = useState(0);

  const hasRows = rows.length > 0;
  const colCount = hasRows ? rows[0].length : 0;

  // 初始化列宽
  useEffect(() => {
    if (rows.length === 0) return;
    const count = rows[0].length;
    if (isControlledWidths) {
      if ((props.columnWidths?.length ?? 0) !== count && props.onColumnWidthsChange) {
        props.onColumnWidthsChange(Array(count).fill(defaultColWidth));
      }
      return;
    }
    setInternalColumnWidths((prev) => {
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

    // scrollLeft：用于同步表头和左右表格，直接读写不触发 React render
    lastScrollLeftRef.current = target.scrollLeft;
    if (headerScrollRef.current) {
      headerScrollRef.current.scrollLeft = target.scrollLeft;
    }
    if (props.onScrollXChange) {
      props.onScrollXChange(target.scrollLeft);
    }

    // scrollTop：会影响虚拟列表窗口，使用 rAF 节流避免滚动时频繁 setState 卡顿
    lastScrollTopRef.current = target.scrollTop;
    if (props.onScrollYChange) {
      props.onScrollYChange(target.scrollTop);
    }

    if (scrollRafRequestedRef.current) return;
    scrollRafRequestedRef.current = true;

    requestAnimationFrame(() => {
      scrollRafRequestedRef.current = false;
      setScrollTop((prev) => {
        const next = lastScrollTopRef.current;
        return prev === next ? prev : next;
      });
    });
  };

  // 外部受控 scrollLeft：用于“只保留一个横向滚动条”的场景
  useEffect(() => {
    const el = containerRef.current;
    const left = props.scrollLeft;
    if (!el || left == null) return;
    if (el.scrollLeft === left) return;
    el.scrollLeft = left;
    if (headerScrollRef.current) {
      headerScrollRef.current.scrollLeft = left;
    }
  }, [props.scrollLeft]);

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

        const apply = (prev: number[]) => {
          const next = [...prev];
          const delta = clientX - startX;
          const newWidth = Math.max(MIN_COL_WIDTH, startWidth + delta);
          next[colIndex] = newWidth;
          return next;
        };

        if (isControlledWidths) {
          if (props.onColumnWidthsChange) {
            props.onColumnWidthsChange(apply(columnWidths));
          }
          return;
        }

        setInternalColumnWidths((prev) => apply(prev));
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

  const getFrozenColsWidth = () => {
    let w = showRowHeader ? rowHeaderWidth : 0;
    for (let i = 0; i < safeFrozenColCount; i += 1) {
      w += getColWidth(i);
    }
    return w;
  };

  const getColLeft = (colIndex: number) => {
    let left = showRowHeader ? rowHeaderWidth : 0;
    for (let i = 0; i < colIndex; i += 1) {
      left += getColWidth(i);
    }
    return left;
  };

  const frozenColOffsets: number[] = useMemo(() => {
    const offsets: number[] = [];
    let acc = showRowHeader ? rowHeaderWidth : 0;
    for (let colIndex = 0; colIndex < safeFrozenColCount; colIndex += 1) {
      offsets[colIndex] = acc;
      acc += getColWidth(colIndex);
    }
    return offsets;
  }, [safeFrozenColCount, columnWidths, showRowHeader, rowHeaderWidth]);

  // 当外部指定 scrollToCell 时，自动滚动让该单元格出现在视口中
  useEffect(() => {
    const target = props.scrollToCell;
    const el = containerRef.current;
    if (!target || !el) return;

    const { rowIndex, colIndex } = target;

    // 竖向：冻结行无需滚动
    if (rowIndex >= safeFrozenRowCount) {
      const scrollRowIndex = rowIndex - safeFrozenRowCount;
      const rowTop = scrollRowIndex * rowHeight;
      const rowBottom = rowTop + rowHeight;
      const viewTop = el.scrollTop;
      const viewBottom = viewTop + el.clientHeight;

      if (rowTop < viewTop) {
        el.scrollTop = rowTop;
      } else if (rowBottom > viewBottom) {
        el.scrollTop = Math.max(0, rowBottom - el.clientHeight);
      }
    }

    // 横向：冻结列无需滚动
    if (colIndex >= safeFrozenColCount) {
      const colLeft = getColLeft(colIndex);
      const colRight = colLeft + getColWidth(colIndex);
      const frozenW = getFrozenColsWidth();
      const viewLeft = el.scrollLeft + frozenW;
      const viewRight = el.scrollLeft + el.clientWidth;

      if (colLeft < viewLeft) {
        el.scrollLeft = Math.max(0, colLeft - frozenW);
      } else if (colRight > viewRight) {
        el.scrollLeft = Math.max(0, colRight - el.clientWidth);
      }
    }
    // 有些环境下 programmatic scroll 不一定触发 scroll 事件，这里显式通知外部同步
    if (props.onScrollXChange) props.onScrollXChange(el.scrollLeft);
    if (props.onScrollYChange) props.onScrollYChange(el.scrollTop);
  }, [props.scrollToCell, safeFrozenRowCount, safeFrozenColCount, rowHeight, viewportHeight, columnWidths]);

  const renderRow = (row: Cell[], rowIndex: number) => {
    const isFrozenRow = rowIndex < safeFrozenRowCount;
    const ctxBase = { isFrozenRow };

    const rowHeader = showRowHeader ? (
      <td
        onContextMenu={(e) => {
          if (onRowHeaderContextMenu) {
            onRowHeaderContextMenu(rowIndex, e);
          }
        }}
        style={{
          position: 'sticky',
          left: 0,
          zIndex: 4,
          border: '1px solid #ddd',
          padding: 2,
          textAlign: 'right',
          userSelect: 'none',
          width: rowHeaderWidth,
          minWidth: rowHeaderWidth,
          backgroundColor: '#f7f7f7',
          fontSize: 12,
          cursor: onRowHeaderContextMenu ? 'context-menu' : 'default',
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
          const stickyLeft = isFrozenCol
            ? frozenColOffsets[colIndex] ?? (showRowHeader ? rowHeaderWidth : 0)
            : undefined;

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
                maxWidth: getColWidth(colIndex),
                height: rowHeight,
                overflow: 'hidden',
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

  const totalTableWidth = useMemo(() => {
    let w = showRowHeader ? rowHeaderWidth : 0;
    for (let i = 0; i < colCount; i += 1) {
      w += getColWidth(i);
    }
    return w;
  }, [showRowHeader, colCount, columnWidths, defaultColWidth, rowHeaderWidth]);

  const ColGroup = useMemo(() => {
    return (
      <colgroup>
        {showRowHeader && <col style={{ width: rowHeaderWidth }} />}
        {Array.from({ length: colCount }, (_, colIndex) => (
          <col key={colIndex} style={{ width: getColWidth(colIndex) }} />
        ))}
      </colgroup>
    );
  }, [showRowHeader, colCount, columnWidths, defaultColWidth, rowHeaderWidth]);

  if (!hasRows) {
    return null;
  }

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
        <table style={{ borderCollapse: 'collapse', tableLayout: 'fixed', width: totalTableWidth }}>
          {ColGroup}
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
                    width: rowHeaderWidth,
                    minWidth: rowHeaderWidth,
                    backgroundColor: '#f0f0f0',
                    fontSize: 12,
                  }}
                >
                  行
                </th>
              )}
              {Array.from({ length: colCount }, (_, colIndex) => {
                const isFrozenCol = colIndex < safeFrozenColCount;
                const stickyLeft = isFrozenCol ? frozenColOffsets[colIndex] ?? (showRowHeader ? rowHeaderWidth : 0) : undefined;
                return (
                  <th
                    key={colIndex}
                    onContextMenu={(e) => {
                      if (onHeaderContextMenu) {
                        onHeaderContextMenu(colIndex, e);
                      }
                    }}
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
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                      cursor: onHeaderContextMenu ? 'context-menu' : 'default',
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
        style={{
          flex: 1,
          minHeight: 0,
          overflowY: 'auto',
          overflowX: props.disableHorizontalScroll ? 'hidden' : 'auto',
        }}
      >
        <table style={{ borderCollapse: 'collapse', tableLayout: 'fixed', width: totalTableWidth }}>
          {ColGroup}
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
