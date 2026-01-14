import React, { useMemo } from 'react';
import type { MergeCell } from '../main/preload';

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

export const MergeSideBySide: React.FC<MergeSideBySideProps> = ({
  rows,
  selected,
  onSelectCell,
}) => {
  if (rows.length === 0) return null;

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

  const renderTable = (side: 'ours' | 'theirs') => (
    <div style={{ flex: 1, overflow: 'auto' }}>
      <div style={{ marginBottom: 4, fontWeight: 'bold', fontSize: 12 }}>
        {side === 'ours' ? 'ours (当前分支)' : 'theirs (合并分支)'}
      </div>
      <table style={{ borderCollapse: 'collapse' }}>
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
                }}
              >
                {label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {diffRowNumbers.map((rowNumber) => {
            const row = rows[rowNumber - 1];
            return (
              <tr key={rowNumber}>
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
                    return <td key={`${rowNumber}-${colNumber}`} />;
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
        </tbody>
      </table>
    </div>
  );

  return (
    <div
      style={{
        display: 'flex',
        gap: 16,
        border: '1px solid #ccc',
        padding: 8,
        maxHeight: '70vh',
        overflow: 'hidden',
      }}
    >
      {renderTable('ours')}
      {renderTable('theirs')}
    </div>
  );
};
