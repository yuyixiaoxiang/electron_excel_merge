import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';
import type { MergeColumnMeta, MergeRowMeta, RowStatus, SheetCell } from '../main/preload';
import { VirtualGrid, VirtualGridRenderCtx } from './VirtualGrid';

const ROW_HEIGHT = 24;
const OVERSCAN_ROWS = 8;
const DEFAULT_FROZEN_HEADER_ROWS = 3;
const DEFAULT_COL_WIDTH = 160;
const FROZEN_COLOR = '#f5f5f5';
const LEFT_DIFF_COLOR = '#d4f8d4';
const RIGHT_DIFF_COLOR = '#ffc8c8';
const MISSING_COLOR = '#f2f2f2';

type DiffSide = 'left' | 'right';

export interface DiffCellData {
  alignedRowNumber: number;
  alignedColNumber: number;
  address: string | null;
  value: string | number | null;
  otherValue: string | number | null;
  sourceRowNumber: number | null;
  sourceColNumber: number | null;
  isDifferent: boolean;
}

interface DiffSideBySideProps {
  leftPath?: string | null;
  rightPath?: string | null;
  leftRows: SheetCell[][];
  rightRows: SheetCell[][];
  rowsMeta?: MergeRowMeta[];
  columnsMeta?: MergeColumnMeta[];
  frozenRowCount?: number;
  selected?: { rowIndex: number; colIndex: number } | null;
  onSelectCell?: (rowIndex: number, colIndex: number) => void;
  onCellChange?: (side: DiffSide, cell: DiffCellData, newValue: string) => void;
  onApplyOtherSideCell?: (side: DiffSide, cell: DiffCellData) => void;
  onApplyOtherSideRow?: (side: DiffSide, cell: DiffCellData) => void;
  onDeleteRow?: (side: DiffSide, cell: DiffCellData) => void;
}

const normalizeComparableValue = (value: string | number | null): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return String(value);
  return String(value).trim();
};

const sameComparableValue = (a: string | number | null, b: string | number | null) =>
  normalizeComparableValue(a) === normalizeComparableValue(b);

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

const DiffSideBySideComponent: React.FC<DiffSideBySideProps> = ({
  leftPath,
  rightPath,
  leftRows,
  rightRows,
  rowsMeta,
  columnsMeta,
  frozenRowCount = DEFAULT_FROZEN_HEADER_ROWS,
  selected,
  onSelectCell,
  onCellChange,
  onApplyOtherSideCell,
  onApplyOtherSideRow,
  onDeleteRow,
}) => {
  const leftScrollRef = useRef<HTMLDivElement | null>(null);
  const rightScrollRef = useRef<HTMLDivElement | null>(null);
  const isSyncingHorizontalRef = useRef(false);
  const isSyncingVerticalRef = useRef(false);
  const [columnWidths, setColumnWidths] = useState<number[]>([]);
  const [editingCell, setEditingCell] = useState<{
    side: DiffSide;
    rowIndex: number;
    colIndex: number;
  } | null>(null);
  const [draftValue, setDraftValue] = useState('');
  const [contextMenu, setContextMenu] = useState<
    | {
        type: 'cell' | 'row';
        side: DiffSide;
        x: number;
        y: number;
        rowIndex: number;
        colIndex: number;
        cell: DiffCellData;
      }
    | null
  >(null);

  const syncScrollX = (from: DiffSide, scrollLeft: number) => {
    const otherRef = from === 'left' ? rightScrollRef : leftScrollRef;
    if (!otherRef.current) return;
    if (isSyncingHorizontalRef.current) return;
    isSyncingHorizontalRef.current = true;
    otherRef.current.scrollLeft = scrollLeft;
    requestAnimationFrame(() => {
      isSyncingHorizontalRef.current = false;
    });
  };

  const syncScrollY = (from: DiffSide, scrollTop: number) => {
    const otherRef = from === 'left' ? rightScrollRef : leftScrollRef;
    if (!otherRef.current) return;
    if (isSyncingVerticalRef.current) return;
    isSyncingVerticalRef.current = true;
    otherRef.current.scrollTop = scrollTop;
    requestAnimationFrame(() => {
      isSyncingVerticalRef.current = false;
    });
  };

  const displayRowsMeta = useMemo(() => {
    if (rowsMeta && rowsMeta.length > 0) {
      return [...rowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
    }
    const maxRows = Math.max(leftRows.length, rightRows.length);
    return Array.from({ length: maxRows }, (_, idx) => ({
      visualRowNumber: idx + 1,
      baseRowNumber: idx + 1,
      oursRowNumber: leftRows[idx] ? idx + 1 : null,
      theirsRowNumber: rightRows[idx] ? idx + 1 : null,
      oursStatus: 'unchanged' as RowStatus,
      theirsStatus: 'unchanged' as RowStatus,
    }));
  }, [rowsMeta, leftRows, rightRows]);

  const displayColumnsMeta = useMemo(() => {
    if (columnsMeta && columnsMeta.length > 0) {
      return [...columnsMeta].sort((a, b) => a.col - b.col);
    }
    const leftColCount = leftRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);
    const rightColCount = rightRows.reduce((max, row) => Math.max(max, row?.length ?? 0), 0);
    const maxCols = Math.max(leftColCount, rightColCount);
    return Array.from({ length: maxCols }, (_, idx) => ({
      col: idx + 1,
      baseCol: idx + 1,
      oursCol: idx + 1,
      theirsCol: idx + 1,
    }));
  }, [columnsMeta, leftRows, rightRows]);

  useEffect(() => {
    const count = displayColumnsMeta.length;
    setColumnWidths((prev) => {
      if (prev.length === count) return prev;
      return Array(count).fill(DEFAULT_COL_WIDTH);
    });
  }, [displayColumnsMeta.length]);

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

  const buildCell = (
    side: DiffSide,
    rowMeta: MergeRowMeta,
    columnMeta: MergeColumnMeta,
  ): DiffCellData => {
    const leftRowNumber = rowMeta.oursRowNumber ?? null;
    const rightRowNumber = rowMeta.theirsRowNumber ?? null;
    const leftColNumber = columnMeta.oursCol ?? null;
    const rightColNumber = columnMeta.theirsCol ?? null;
    const leftCell =
      leftRowNumber && leftColNumber ? leftRows[leftRowNumber - 1]?.[leftColNumber - 1] ?? null : null;
    const rightCell =
      rightRowNumber && rightColNumber ? rightRows[rightRowNumber - 1]?.[rightColNumber - 1] ?? null : null;
    const leftValue = leftCell?.value ?? null;
    const rightValue = rightCell?.value ?? null;
    const isDifferent = !sameComparableValue(leftValue, rightValue);

    if (side === 'left') {
      return {
        alignedRowNumber: rowMeta.visualRowNumber,
        alignedColNumber: columnMeta.col,
        address: leftCell?.address ?? null,
        value: leftValue,
        otherValue: rightValue,
        sourceRowNumber: leftRowNumber,
        sourceColNumber: leftColNumber,
        isDifferent,
      };
    }

    return {
      alignedRowNumber: rowMeta.visualRowNumber,
      alignedColNumber: columnMeta.col,
      address: rightCell?.address ?? null,
      value: rightValue,
      otherValue: leftValue,
      sourceRowNumber: rightRowNumber,
      sourceColNumber: rightColNumber,
      isDifferent,
    };
  };

  const leftGridRows = useMemo(
    () =>
      displayRowsMeta.map((rowMeta) =>
        displayColumnsMeta.map((columnMeta) => buildCell('left', rowMeta, columnMeta)),
      ),
    [displayRowsMeta, displayColumnsMeta, leftRows, rightRows],
  );

  const rightGridRows = useMemo(
    () =>
      displayRowsMeta.map((rowMeta) =>
        displayColumnsMeta.map((columnMeta) => buildCell('right', rowMeta, columnMeta)),
      ),
    [displayRowsMeta, displayColumnsMeta, leftRows, rightRows],
  );

  const handleSelect = (rowIndex: number, colIndex: number) => {
    if (onSelectCell) onSelectCell(rowIndex, colIndex);
  };
  const handleRowHeaderContextMenu = (side: DiffSide, rowIndex: number, e: React.MouseEvent<HTMLTableCellElement>) => {
    e.preventDefault();
    e.stopPropagation();
    const rowMeta = displayRowsMeta[rowIndex];
    if (!rowMeta) return;
    const firstColumnMeta =
      displayColumnsMeta[0] ?? {
        col: 1,
        baseCol: 1,
        oursCol: 1,
        theirsCol: 1,
      };
    const cell = buildCell(side, rowMeta, firstColumnMeta);
    handleSelect(rowIndex, 0);
    setContextMenu({
      type: 'row',
      side,
      x: e.clientX,
      y: e.clientY,
      rowIndex,
      colIndex: 0,
      cell,
    });
  };

  const beginEdit = (side: DiffSide, rowIndex: number, colIndex: number, cell: DiffCellData) => {
    if (!cell.address) return;
    setEditingCell({ side, rowIndex, colIndex });
    setDraftValue(cell.value == null ? '' : String(cell.value));
    handleSelect(rowIndex, colIndex);
  };

  const commitEdit = (side: DiffSide, rowIndex: number, colIndex: number, cell: DiffCellData) => {
    if (!onCellChange || !cell.address) {
      setEditingCell(null);
      return;
    }
    onCellChange(side, cell, draftValue);
    setEditingCell(null);
  };

  const makeRowHeaderRenderer =
    (side: DiffSide) =>
    (rowIndex: number) => {
      const meta = displayRowsMeta[rowIndex];
      if (!meta) return rowIndex + 1;
      const indicator = getRowStatusIndicator(side === 'left' ? meta.oursStatus : meta.theirsStatus);
      const rowNumber =
        side === 'left'
          ? meta.oursRowNumber ?? meta.baseRowNumber ?? meta.visualRowNumber
          : meta.theirsRowNumber ?? meta.baseRowNumber ?? meta.visualRowNumber;
      return (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'flex-end',
            gap: 4,
            overflow: 'hidden',
          }}
        >
          {indicator.symbol && (
            <span style={{ color: indicator.color, fontWeight: 700 }}>{indicator.symbol}</span>
          )}
          <span>{rowNumber ?? ''}</span>
        </div>
      );
    };

  const makeHeaderRenderer =
    (side: DiffSide) =>
    (colIndex: number) => {
      const meta = displayColumnsMeta[colIndex];
      if (!meta) return colNumberToLabel(colIndex + 1);
      const sourceColNumber = side === 'left' ? meta.oursCol : meta.theirsCol;
      return sourceColNumber ? colNumberToLabel(sourceColNumber) : '—';
    };

  const makeRenderCell =
    (side: DiffSide) =>
    (cell: DiffCellData | null, ctx: VirtualGridRenderCtx) => {
      if (!cell) return null;
      const isEditing =
        editingCell?.side === side &&
        editingCell.rowIndex === ctx.rowIndex &&
        editingCell.colIndex === ctx.colIndex;
      const displayValue = cell.value == null ? '' : String(cell.value);

      if (isEditing) {
        return (
          <input
            autoFocus
            value={draftValue}
            onFocus={() => handleSelect(ctx.rowIndex, ctx.colIndex)}
            onChange={(e: ChangeEvent<HTMLInputElement>) => setDraftValue(e.target.value)}
            onBlur={() => commitEdit(side, ctx.rowIndex, ctx.colIndex, cell)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                commitEdit(side, ctx.rowIndex, ctx.colIndex, cell);
              }
              if (e.key === 'Escape') {
                e.preventDefault();
                setEditingCell(null);
              }
            }}
            style={{
              width: '100%',
              boxSizing: 'border-box',
              border: 'none',
              outline: 'none',
              backgroundColor: 'transparent',
            }}
          />
        );
      }

      return (
        <div
          onMouseDown={() => handleSelect(ctx.rowIndex, ctx.colIndex)}
          onDoubleClick={() => beginEdit(side, ctx.rowIndex, ctx.colIndex, cell)}
          onContextMenu={(e) => {
            e.preventDefault();
            e.stopPropagation();
            handleSelect(ctx.rowIndex, ctx.colIndex);
            setContextMenu({
              type: 'cell',
              side,
              x: e.clientX,
              y: e.clientY,
              rowIndex: ctx.rowIndex,
              colIndex: ctx.colIndex,
              cell,
            });
          }}
          title={
            `${cell.address ?? '无对应单元格'}\n` +
            `当前: ${displayValue}\n` +
            `另一侧: ${cell.otherValue == null ? '' : String(cell.otherValue)}`
          }
          style={{
            width: '100%',
            height: '100%',
            boxSizing: 'border-box',
            backgroundColor: 'transparent',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
            cursor: cell.address ? 'text' : 'default',
            userSelect: 'none',
            color: cell.address ? '#111' : '#888',
          }}
        >
          {displayValue}
        </div>
      );
    };

  const makeCellStyle =
    (side: DiffSide) =>
    (cell: DiffCellData | null, ctx: VirtualGridRenderCtx): React.CSSProperties => {
      const style: React.CSSProperties = {};
      if (!cell) return style;

      if (cell.isDifferent) {
        if (!cell.address) {
          style.backgroundColor = MISSING_COLOR;
        } else {
          style.backgroundColor = side === 'left' ? LEFT_DIFF_COLOR : RIGHT_DIFF_COLOR;
        }
      } else if (ctx.isFrozenRow || ctx.isFrozenCol) {
        style.backgroundColor = FROZEN_COLOR;
      }

      if (selected && selected.rowIndex === ctx.rowIndex && selected.colIndex === ctx.colIndex) {
        style.outline = '2px solid #ff8000';
        style.outlineOffset = '-2px';
        style.position = 'relative';
        style.zIndex = 6;
      }

      return style;
    };

  const scrollToCell = useMemo(() => {
    if (!selected) return null;
    return { rowIndex: selected.rowIndex, colIndex: selected.colIndex };
  }, [selected]);

  const hasData = leftGridRows.length > 0 || rightGridRows.length > 0;

  if (!hasData) {
    return <div>请选择左右两个 Excel 文件。</div>;
  }

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
        <div
          style={{
            flex: 1,
            minWidth: 0,
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
          }}
        >
          left{leftPath ? `: ${leftPath}` : ''}
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
          right{rightPath ? `: ${rightPath}` : ''}
        </div>
      </div>
      <div style={{ display: 'flex', gap: 16, flex: 1, minHeight: 0 }}>
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
          <VirtualGrid<DiffCellData>
            rows={leftGridRows}
            rowHeight={ROW_HEIGHT}
            overscanRows={OVERSCAN_ROWS}
            frozenRowCount={frozenRowCount}
            frozenColCount={0}
            rowHeaderWidth={64}
            showRowHeader
            renderRowHeader={makeRowHeaderRenderer('left')}
            onRowHeaderContextMenu={(rowIndex, e) => handleRowHeaderContextMenu('left', rowIndex, e)}
            renderCell={makeRenderCell('left')}
            getCellStyle={makeCellStyle('left')}
            renderHeaderCell={makeHeaderRenderer('left')}
            defaultColWidth={DEFAULT_COL_WIDTH}
            columnWidths={columnWidths}
            onColumnWidthsChange={setColumnWidths}
            containerRef={leftScrollRef as React.RefObject<HTMLDivElement>}
            onScrollXChange={(left) => syncScrollX('left', left)}
            onScrollYChange={(top) => syncScrollY('left', top)}
            scrollToCell={scrollToCell}
          />
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minWidth: 0 }}>
          <VirtualGrid<DiffCellData>
            rows={rightGridRows}
            rowHeight={ROW_HEIGHT}
            overscanRows={OVERSCAN_ROWS}
            frozenRowCount={frozenRowCount}
            frozenColCount={0}
            rowHeaderWidth={64}
            showRowHeader
            renderRowHeader={makeRowHeaderRenderer('right')}
            onRowHeaderContextMenu={(rowIndex, e) => handleRowHeaderContextMenu('right', rowIndex, e)}
            renderCell={makeRenderCell('right')}
            getCellStyle={makeCellStyle('right')}
            renderHeaderCell={makeHeaderRenderer('right')}
            defaultColWidth={DEFAULT_COL_WIDTH}
            columnWidths={columnWidths}
            onColumnWidthsChange={setColumnWidths}
            containerRef={rightScrollRef as React.RefObject<HTMLDivElement>}
            onScrollXChange={(left) => syncScrollX('right', left)}
            onScrollYChange={(top) => syncScrollY('right', top)}
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
              minWidth: 220,
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ padding: '6px 10px', borderBottom: '1px solid #eee', color: '#666' }}>
              {contextMenu.type === 'row'
                ? `${contextMenu.side === 'left' ? 'left' : 'right'} 行 ${contextMenu.cell.alignedRowNumber}`
                : `${contextMenu.side === 'left' ? 'left' : 'right'} ${contextMenu.cell.address ?? '无对应单元格'}`}
            </div>
            {contextMenu.type === 'cell' && (
              <button
                type="button"
                style={{
                  width: '100%',
                  textAlign: 'left',
                  padding: '6px 10px',
                  border: 'none',
                  background: 'white',
                  cursor: 'pointer',
                }}
                onClick={() => {
                  if (onApplyOtherSideCell) {
                    onApplyOtherSideCell(contextMenu.side, contextMenu.cell);
                  }
                  setContextMenu(null);
                }}
              >
                使用另一边相同位置的单元格
              </button>
            )}
            <button
              type="button"
              style={{
                width: '100%',
                textAlign: 'left',
                padding: '6px 10px',
                border: 'none',
                background: 'white',
                cursor: 'pointer',
              }}
              onClick={() => {
                if (onApplyOtherSideRow) {
                  onApplyOtherSideRow(contextMenu.side, contextMenu.cell);
                }
                setContextMenu(null);
              }}
            >
              使用另一边整行
            </button>
            <button
              type="button"
              style={{
                width: '100%',
                textAlign: 'left',
                padding: '6px 10px',
                border: 'none',
                background: 'white',
                cursor: 'pointer',
                color: '#b00020',
              }}
              onClick={() => {
                if (onDeleteRow) {
                  onDeleteRow(contextMenu.side, contextMenu.cell);
                }
                setContextMenu(null);
              }}
            >
              删除本行
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

export const DiffSideBySide = React.memo(DiffSideBySideComponent);
