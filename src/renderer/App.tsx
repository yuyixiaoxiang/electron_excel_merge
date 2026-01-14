import React, { useCallback, useEffect, useMemo, useState } from 'react';
import type {
  CellChange,
  CliThreeWayInfo,
  MergeCell,
  MergeSheetData,
  OpenResult,
  SaveMergeRequest,
  SheetCell,
  SheetData,
  ThreeWayOpenResult,
} from '../main/preload';
import { ExcelTable } from './ExcelTable';
import { MergeSideBySide } from './MergeSideBySide';

type ViewMode = 'single' | 'merge';

export const App: React.FC = () => {
  const [mode, setMode] = useState<ViewMode>('single');

  // 单文件编辑状态
  const [filePath, setFilePath] = useState<string | null>(null);
  const [sheetName, setSheetName] = useState<string | null>(null);
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [selectedSheetIndex, setSelectedSheetIndex] = useState<number>(0);
  const [rows, setRows] = useState<SheetCell[][]>([]);
  const [changes, setChanges] = useState<Map<string, CellChange>>(new Map());
  const [saving, setSaving] = useState(false);

  // 三方 diff 状态
  const [mergeSheets, setMergeSheets] = useState<MergeSheetData[]>([]);
  const [selectedMergeSheetIndex, setSelectedMergeSheetIndex] = useState<number>(0);
  const [mergeRows, setMergeRows] = useState<MergeCell[][]>([]);
  const [mergeInfo, setMergeInfo] = useState<{
    basePath: string;
    oursPath: string;
    theirsPath: string;
    sheetName: string;
  } | null>(null);
  const [cliInfo, setCliInfo] = useState<CliThreeWayInfo | null>(null);
  const [selectedMergeCell, setSelectedMergeCell] = useState<{
    rowIndex: number;
    colIndex: number;
  } | null>(null);

  const handleOpen = useCallback(async () => {
    const result: OpenResult | null = await window.excelAPI.openFile();
    if (!result) return;

    setMode('single');
    setFilePath(result.filePath);
    const allSheets = result.sheets && result.sheets.length > 0 ? result.sheets : [result.sheet];
    setSheets(allSheets);
    setSelectedSheetIndex(0);
    setSheetName(allSheets[0]?.sheetName ?? null);
    setRows(allSheets[0]?.rows ?? []);
    setChanges(new Map());
  }, []);

  const handleOpenThreeWay = useCallback(async () => {
    const result: ThreeWayOpenResult | null = await window.excelAPI.openThreeWay();
    if (!result) return;

    setMode('merge');
    const allMergeSheets = result.sheets && result.sheets.length > 0 ? result.sheets : [result.sheet];
    setMergeSheets(allMergeSheets);
    setSelectedMergeSheetIndex(0);
    setMergeRows(allMergeSheets[0]?.rows ?? []);
    setMergeInfo({
      basePath: result.basePath,
      oursPath: result.oursPath,
      theirsPath: result.theirsPath,
      sheetName: allMergeSheets[0]?.sheetName ?? result.sheet.sheetName,
    });
    setSelectedMergeCell(null);
  }, []);

  // 如果是 git/Fork 调用并传入了 CLI three-way 参数，启动后自动进入 merge 视图
  useEffect(() => {
    (async () => {
      try {
        const info = await window.excelAPI.getCliThreeWayInfo();
        if (info) {
          setCliInfo(info);
          await handleOpenThreeWay();
        }
      } catch {
        // 忽略错误，保持交互式模式可用
      }
    })();
  }, [handleOpenThreeWay]);

  const handleCellChange = useCallback(
    (address: string, newValue: string) => {
      setRows((prev) =>
        prev.map((row) =>
          row.map((cell) =>
            cell.address === address
              ? {
                  ...cell,
                  value: newValue === '' ? null : newValue,
                }
              : cell,
          ),
        ),
      );

      setChanges((prev) => {
        const next = new Map(prev);
        next.set(address, {
          address,
          newValue: newValue === '' ? null : newValue,
        });
        return next;
      });
    },
    [],
  );

  const handleSave = useCallback(async () => {
    if (!filePath || changes.size === 0) return;
    setSaving(true);
    try {
      const changeList = Array.from(changes.values());
      await window.excelAPI.saveChanges(changeList);
      setChanges(new Map());
      // 不需要刷新格式，只要值正确写回即可
    } finally {
      setSaving(false);
    }
  }, [changes, filePath]);

  const hasData = useMemo(() => rows.length > 0, [rows]);
  const hasMergeData = useMemo(() => mergeRows.length > 0, [mergeRows]);

  const handleSelectMergeCell = useCallback((rowIndex: number, colIndex: number) => {
    setSelectedMergeCell({ rowIndex, colIndex });
  }, []);

  const handleApplyMergeChoice = useCallback(
    (source: 'base' | 'ours' | 'theirs') => {
      if (!selectedMergeCell) return;

      const { rowIndex, colIndex } = selectedMergeCell;
      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newRows = sheet.rows.map((row, rIdx) =>
            row.map((cell, cIdx) => {
              if (rIdx !== rowIndex || cIdx !== colIndex) return cell;
              let value: string | number | null;
              if (source === 'base') value = cell.baseValue;
              else if (source === 'ours') value = cell.oursValue;
              else value = cell.theirsValue;
              return { ...cell, mergedValue: value };
            }),
          );
          return { ...sheet, rows: newRows };
        }),
      );

      // 同步当前视图的 rows
      setMergeRows((prev) =>
        prev.map((row, rIdx) =>
          row.map((cell, cIdx) => {
            if (rIdx !== rowIndex || cIdx !== colIndex) return cell;
            let value: string | number | null;
            if (source === 'base') value = cell.baseValue;
            else if (source === 'ours') value = cell.oursValue;
            else value = cell.theirsValue;
            return { ...cell, mergedValue: value };
          }),
        ),
      );
    },
    [selectedMergeCell, selectedMergeSheetIndex],
  );

  const handleSaveMergeToFile = useCallback(async () => {
    if (!mergeInfo || mergeSheets.length === 0) return;

    // 生成本次合并的概要信息：统计所有 sheet 中 status !== 'unchanged' 的单元格
    const changedCells: { sheetName: string; address: string; ours: any; theirs: any; merged: any }[] = [];
    mergeSheets.forEach((sheet) => {
      sheet.rows.forEach((row: MergeCell[]) => {
        row.forEach((cell: MergeCell) => {
          if (cell.status !== 'unchanged') {
            changedCells.push({
              sheetName: sheet.sheetName,
              address: cell.address,
              ours: cell.oursValue,
              theirs: cell.theirsValue,
              merged: cell.mergedValue,
            });
          }
        });
      });
    });

    const formatVal = (v: any) => (v === null || v === undefined ? '' : String(v));

    const maxLines = 100;
    const lines = changedCells.slice(0, maxLines).map((c) =>
      `[${c.sheetName}] 单元格 ${c.address}: ours="${formatVal(c.ours)}"  |  theirs="${formatVal(
        c.theirs,
      )}"  |  合并="${formatVal(c.merged)}"`,
    );

    if (changedCells.length > maxLines) {
      lines.push(`…… 还有 ${changedCells.length - maxLines} 个单元格未展示`);
    }

    const preview =
      `本次合并将影响 ${changedCells.length} 个单元格（覆盖所有工作表）：` +
      (lines.length ? `\n\n${lines.join('\n')}` : '\n(无差异单元格——仅写回了当前值)') +
      '\n\n注意：保存时会将所有工作表的合并结果一并写入目标 Excel 文件。' +
      '\n\n确认要将以上结果写入 Excel 文件吗？';

    const confirmed = window.confirm(preview);
    if (!confirmed) return;

    const cells = mergeSheets.flatMap((sheet: MergeSheetData) =>
      sheet.rows.flatMap((row: MergeCell[]) =>
        row.map((cell: MergeCell) => ({
          sheetName: sheet.sheetName,
          address: cell.address,
          value: cell.mergedValue,
        })),
      ),
    );

    const payload: SaveMergeRequest = {
      templatePath: mergeInfo.oursPath,
      cells,
    };

    try {
      const result = await window.excelAPI.saveMergeResult(payload);
      if (!result.success || result.cancelled) {
        const msg = result.errorMessage ?? '未知错误，可能是目标文件被占用或没有写入权限。';
        alert(`保存合并结果失败：${msg}`);
        return;
      }

      alert(`合并结果已保存到: ${result.filePath ?? ''}`);
    } catch (e) {
      alert(`保存合并结果失败：${String(e)}`);
    }
  }, [mergeInfo, mergeRows]);

  return (
    <div style={{ padding: 16, fontFamily: 'sans-serif' }}>
      <h1>Excel Viewer / Merge Tool (Electron + React + TypeScript)</h1>
      <div style={{ marginBottom: 12 }}>
        <button onClick={handleOpen}>打开单个 Excel 文件</button>
        <button
          onClick={handleSave}
          disabled={mode !== 'single' || !filePath || changes.size === 0 || saving}
          style={{ marginLeft: 8 }}
        >
          {saving ? '保存中…' : '保存修改到原 Excel'}
        </button>
        <button onClick={handleOpenThreeWay} style={{ marginLeft: 16 }}>
          打开三方 Merge/Diff（base / ours / theirs）
        </button>
        {mode === 'merge' && hasMergeData && mergeInfo && (
          <>
            <button onClick={handleSaveMergeToFile} style={{ marginLeft: 8 }}>
              {cliInfo?.mode === 'merge'
                ? '将合并结果写回 Git 合并文件（MERGED，解决冲突）'
                : cliInfo?.mode === 'diff'
                ? '将合并结果覆盖 ours（当前分支）文件'
                : '保存合并结果为新的 Excel 文件（以 ours 为格式模板）'}
            </button>
            <span style={{ marginLeft: 8, fontSize: 12, color: '#666' }}>
              {cliInfo
                ? '（本次操作会将所有工作表的合并结果写入 Git 传入的目标文件，保存后回到 Git 执行 git add 即可完成冲突解决）'
                : '（注意：保存时会将所有工作表的合并结果一并写入目标文件）'}
            </span>
          </>
        )}
      </div>

      {mode === 'single' && filePath && (
        <div style={{ marginBottom: 8 }}>
          <div>当前文件: {filePath}</div>
          <div style={{ display: 'flex', alignItems: 'center', marginTop: 4 }}>
            <span>工作表:</span>
            <div
              style={{
                display: 'inline-flex',
                marginLeft: 4,
                borderBottom: '1px solid #ccc',
                gap: 4,
              }}
            >
              {sheets.map((s, idx) => {
                const isActive = idx === selectedSheetIndex;
                return (
                  <button
                    key={s.sheetName || idx}
                    type="button"
                    onClick={() => {
                      setSelectedSheetIndex(idx);
                      const sheet = sheets[idx];
                      setSheetName(sheet?.sheetName ?? null);
                      setRows(sheet?.rows ?? []);
                      setChanges(new Map());
                    }}
                    style={{
                      padding: '2px 8px',
                      fontSize: 12,
                      borderRadius: '4px 4px 0 0',
                      border: '1px solid #ccc',
                      borderBottom: isActive ? '2px solid white' : '1px solid #ccc',
                      backgroundColor: isActive ? '#ffffff' : '#f5f5f5',
                      cursor: 'pointer',
                    }}
                  >
                    {s.sheetName || `Sheet${idx + 1}`}
                  </button>
                );
              })}
            </div>
          </div>
          <div>已修改单元格数量: {changes.size}</div>
        </div>
      )}

      {mode === 'single' && (
        hasData ? <ExcelTable rows={rows} onCellChange={handleCellChange} /> : <div>请先打开一个 .xlsx 文件。</div>
      )}

      {mode === 'merge' && (
        hasMergeData && mergeInfo ? (
          <div>
            <div style={{ marginBottom: 8 }}>
              {/* diff 模式（仅传 LOCAL/REMOTE）时不显示 base 行 */}
              {!(cliInfo && cliInfo.mode === 'diff') && <div>base: {mergeInfo.basePath}</div>}
              <div>ours: {mergeInfo.oursPath}</div>
              <div>theirs: {mergeInfo.theirsPath}</div>
              {cliInfo?.mergedPath && (
                <div>merged(写回目标): {cliInfo.mergedPath}</div>
              )}
              <div style={{ display: 'flex', alignItems: 'center', marginTop: 4 }}>
                <span>工作表:</span>
                <div
                  style={{
                    display: 'inline-flex',
                    marginLeft: 4,
                    borderBottom: '1px solid #ccc',
                    gap: 4,
                  }}
                >
                  {mergeSheets.map((s, idx) => {
                    const isActive = idx === selectedMergeSheetIndex;
                    return (
                      <button
                        key={s.sheetName || idx}
                        type="button"
                        onClick={() => {
                          setSelectedMergeSheetIndex(idx);
                          const sheet = mergeSheets[idx];
                          setMergeInfo((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  sheetName: sheet?.sheetName ?? prev.sheetName,
                                }
                              : prev,
                          );
                          setMergeRows(sheet?.rows ?? []);
                          setSelectedMergeCell(null);
                        }}
                        style={{
                          padding: '2px 8px',
                          fontSize: 12,
                          borderRadius: '4px 4px 0 0',
                          border: '1px solid #ccc',
                          borderBottom: isActive ? '2px solid white' : '1px solid #ccc',
                          backgroundColor: isActive ? '#ffffff' : '#f5f5f5',
                          cursor: 'pointer',
                        }}
                      >
                        {s.sheetName || `Sheet${idx + 1}`}
                      </button>
                    );
                  })}
                </div>
              </div>
              <div style={{ marginTop: 4, fontSize: 12 }}>
                颜色说明：绿色 = 只在 ours 改变，蓝色 = 只在 theirs 改变，黄色 = 双方都改成相同值，红色 = 冲突（双方修改成不同值）。比较时只看单元格值，忽略格式。
              </div>
            </div>
            <MergeSideBySide
              rows={mergeRows}
              selected={selectedMergeCell}
              onSelectCell={handleSelectMergeCell}
            />
            {selectedMergeCell && (() => {
              const cell = mergeRows[selectedMergeCell.rowIndex]?.[selectedMergeCell.colIndex];
              if (!cell) return null;
              return (
                <div style={{ marginTop: 8, padding: 8, border: '1px solid #ccc', fontSize: 12 }}>
                  <div>当前单元格: {cell.address}</div>
                  {/* diff 模式下不显示 base */}
                  {!(cliInfo && cliInfo.mode === 'diff') && (
                    <div>base: {cell.baseValue ?? ''}</div>
                  )}
                  <div>ours: {cell.oursValue ?? ''}</div>
                  <div>theirs: {cell.theirsValue ?? ''}</div>
                  <div>当前合并值: {cell.mergedValue ?? ''}</div>
                  <div style={{ marginTop: 4 }}>
                    <button onClick={() => handleApplyMergeChoice('base')}>用 base</button>
                    <button onClick={() => handleApplyMergeChoice('ours')} style={{ marginLeft: 4 }}>
                      用 ours
                    </button>
                    <button onClick={() => handleApplyMergeChoice('theirs')} style={{ marginLeft: 4 }}>
                      用 theirs
                    </button>
                  </div>
                </div>
              );
            })()}
          </div>
        ) : (
          <div>请先选择 base / ours / theirs 三个 Excel 文件。</div>
        )
      )}
    </div>
  );
};
