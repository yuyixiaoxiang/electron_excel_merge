import React from 'react';
import type { MergeCell } from '../main/preload';

export interface MergeTableProps {
  rows: MergeCell[][];
  selected?: { rowIndex: number; colIndex: number } | null;
  onSelectCell?: (rowIndex: number, colIndex: number) => void;
}

const getBackgroundColor = (status: MergeCell['status']): string => {
  switch (status) {
    case 'unchanged':
      return 'white';
    case 'ours-changed':
      return '#d4f8d4'; // light green
    case 'theirs-changed':
      return '#d4e8ff'; // light blue
    case 'both-changed-same':
      return '#fff6bf'; // light yellow
    case 'conflict':
      return '#ffc8c8'; // light red
    default:
      return 'white';
  }
};

export const MergeTable: React.FC<MergeTableProps> = ({ rows, selected, onSelectCell }) => {
  if (rows.length === 0) return null;

  const colCount = rows[0].length;
  const columnLabels = Array.from({ length: colCount }, (_, i) =>
    String.fromCharCode('A'.charCodeAt(0) + (i % 26)),
  );

  return (
    <div style={{ overflow: 'auto', maxHeight: '70vh', border: '1px solid #ccc' }}>
      <table style={{ borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            {columnLabels.map((label, colIndex) => (
              <th
                key={colIndex}
                style={{
                  border: '1px solid #ddd',
                  padding: 2,
                  textAlign: 'center',
                  userSelect: 'none',
                  backgroundColor: '#f0f0f0',
                }}
              >
                {label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {row.map((cell, colIndex) => {
                const isSelected =
                  selected && selected.rowIndex === rowIndex && selected.colIndex === colIndex;

                return (
                  <td
                    key={cell.address}
                    style={{
                      border: isSelected ? '2px solid #ff8000' : '1px solid #ddd',
                      padding: 2,
                      backgroundColor: getBackgroundColor(cell.status),
                      fontSize: 12,
                    }}
                    title={`地址: ${cell.address}\nbase: ${cell.baseValue ?? ''}\nours: ${cell.oursValue ?? ''}\ntheirs: ${cell.theirsValue ?? ''}`}
                    onClick={() => onSelectCell && onSelectCell(rowIndex, colIndex)}
                  >
                    {cell.mergedValue === null ? '' : String(cell.mergedValue)}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
