import React, { ChangeEvent, useEffect, useState } from 'react';
import type { SheetCell } from '../main/preload';
import { VirtualGrid, VirtualGridRenderCtx } from './VirtualGrid';

/**
 * 单文件编辑视图中使用的表格组件。
 *
 * 基于通用 VirtualGrid 实现虚拟滚动与冻结行/列。
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
  const [editingAddress, setEditingAddress] = useState<string | null>(null);
  const [draftValue, setDraftValue] = useState<string>('');

  // 当外部选中变化时，如果正在编辑其他单元格，则结束编辑
  useEffect(() => {
    if (editingAddress && selectedAddress && editingAddress !== selectedAddress) {
      setEditingAddress(null);
    }
  }, [selectedAddress, editingAddress]);

  const commitEdit = (cell: SheetCell) => {
    onCellChange(cell.address, draftValue);
    setEditingAddress(null);
  };

  const handleSelect = (cell: SheetCell) => {
    if (onCellSelect) {
      onCellSelect(cell);
    }
  };

  const renderCell = (cell: SheetCell | null, ctx: VirtualGridRenderCtx) => {
    if (!cell) return null;

    const isEditing = editingAddress === cell.address;
    const displayValue = cell.value === null ? '' : String(cell.value);

    if (isEditing) {
      return (
        <input
          autoFocus
          onFocus={() => handleSelect(cell)}
          style={{
            width: '100%',
            boxSizing: 'border-box',
            border: 'none',
            outline: 'none',
            backgroundColor: 'transparent',
          }}
          value={draftValue}
          onChange={(e: ChangeEvent<HTMLInputElement>) => setDraftValue(e.target.value)}
          onBlur={() => commitEdit(cell)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              commitEdit(cell);
            }
            if (e.key === 'Escape') {
              e.preventDefault();
              setEditingAddress(null);
            }
          }}
        />
      );
    }

    return (
      <div
        onMouseDown={() => handleSelect(cell)}
        onDoubleClick={() => {
          setEditingAddress(cell.address);
          setDraftValue(displayValue);
        }}
        title={displayValue}
        style={{
          width: '100%',
          height: '100%',
          boxSizing: 'border-box',
          backgroundColor: 'transparent',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          cursor: 'text',
          userSelect: 'none',
        }}
      >
        {displayValue}
      </div>
    );
  };

  const getCellStyle = (cell: SheetCell | null, ctx: VirtualGridRenderCtx) => {
    const base: React.CSSProperties = {};
    const isFrozenCell = ctx.isFrozenRow || ctx.isFrozenCol;

    if (isFrozenCell) {
      base.backgroundColor = '#f5f5f5';
    }

    if (cell && selectedAddress === cell.address) {
      base.border = '2px solid #00aa00';
    }

    return base;
  };

  const renderRowHeader = (rowIndex: number, row: SheetCell[]) => {
    const excelRowNumber = row[0]?.row ?? rowIndex + 1;
    return excelRowNumber;
  };

  const renderHeaderCell = (colIndex: number) => {
    let n = colIndex + 1;
    let s = '';
    while (n > 0) {
      n -= 1;
      s = String.fromCharCode('A'.charCodeAt(0) + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  };

  return (
    <VirtualGrid<SheetCell>
      rows={rows}
      rowHeight={24}
      overscanRows={10}
      frozenRowCount={frozenRowCount}
      frozenColCount={frozenColCount}
      showRowHeader
      renderRowHeader={renderRowHeader}
      renderCell={renderCell}
      getCellStyle={getCellStyle}
      renderHeaderCell={renderHeaderCell}
      defaultColWidth={120}
    />
  );
};

export const ExcelTable = React.memo(ExcelTableComponent);
