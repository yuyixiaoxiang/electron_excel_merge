import React, { useCallback, useEffect, useMemo, useState } from 'react';
import type {
  CellChange,
  CliThreeWayInfo,
  MergeCell,
  MergeRowMeta,
  MergeSheetData,
  OpenResult,
  SaveMergeRequest,
  SheetCell,
  SheetData,
  ThreeWayDiffRequest,
  ThreeWayOpenResult,
  ThreeWayRowResult,
} from '../main/preload';
import { ExcelTable } from './ExcelTable';
import { MergeSideBySide } from './MergeSideBySide';
import { VirtualGrid } from './VirtualGrid';

/**
 * 应用根组件：
 * - single 模式：单个 Excel 文件的查看与轻量编辑；
 * - merge 模式：base / ours / theirs 三方合并与结果写回。
 */
type ViewMode = 'single' | 'merge';

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

const makeAddress = (colNumber: number, rowNumber: number): string => {
  return `${colNumberToLabel(colNumber)}${rowNumber}`;
};

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
  // 当前单文件模式下选中的单元格（用于顶部“公式栏”显示）
  const [selectedSingleCell, setSelectedSingleCell] = useState<SheetCell | null>(null);
  // 固定在顶部的首行数，默认 3 行
  const [frozenRowCount, setFrozenRowCount] = useState<number>(3);
  // 固定在左侧的列数（不含最左侧行号列），默认 0 列
  const [frozenColCount, setFrozenColCount] = useState<number>(0);
  // merge/diff 视图中固定在顶部展示的行数，默认 3 行
  const [mergeFrozenRowCount, setMergeFrozenRowCount] = useState<number>(3);
  const [rowSimilarityThreshold, setRowSimilarityThreshold] = useState<number>(0.9);

  // 三方 diff 状态
  const [mergeSheets, setMergeSheets] = useState<MergeSheetData[]>([]);
  const [selectedMergeSheetIndex, setSelectedMergeSheetIndex] = useState<number>(0);
  const [mergeCells, setMergeCells] = useState<MergeCell[]>([]);
  const [mergeRowsMeta, setMergeRowsMeta] = useState<MergeRowMeta[]>([]);
  const [primaryKeyCol, setPrimaryKeyCol] = useState<number>(1);
  const [autoHasPrimaryKey, setAutoHasPrimaryKey] = useState<boolean>(true);
  const [lastPrimaryKeyCol, setLastPrimaryKeyCol] = useState<number>(1);
  const [primaryKeyHint, setPrimaryKeyHint] = useState<string>('');
  // 记录“用户已确认合并”的单元格（resolved），按 sheetIndex 分组，key="row:col"（1-based）
  const [resolvedBySheet, setResolvedBySheet] = useState<Map<number, Set<string>>>(new Map());
  // 当前选中行的三方原始值（用于构建 merged 行视图；目前以 ours 为基准覆盖 mergedValue）
  const [selectedThreeWayRow, setSelectedThreeWayRow] = useState<ThreeWayRowResult | null>(null);
  const [mergedRowColumnWidths, setMergedRowColumnWidths] = useState<number[]>([]);

  // 底部 merged 行视图：当列数变化时重置列宽
  useEffect(() => {
    const colCount = selectedThreeWayRow?.colCount ?? 0;
    if (!colCount) {
      setMergedRowColumnWidths([]);
      return;
    }
    setMergedRowColumnWidths((prev) => (prev.length === colCount ? prev : Array(colCount).fill(120)));
  }, [selectedThreeWayRow?.colCount]);
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

  /**
   * 交互式选择一个 Excel 文件并进入单文件编辑模式。
   */
  const handleOpen = useCallback(async () => {
    const result: OpenResult | null = await window.excelAPI.openFile();
    if (!result) return;

    setMode('single');
    setFilePath(result.filePath);
    setSelectedSingleCell(null);
    const allSheets = result.sheets && result.sheets.length > 0 ? result.sheets : [result.sheet];
    setSheets(allSheets);
    setSelectedSheetIndex(0);
    setSheetName(allSheets[0]?.sheetName ?? null);
    setRows(allSheets[0]?.rows ?? []);
    setChanges(new Map());
  }, []);

  /**
   * 交互式选择 base / ours / theirs（三方 diff），并切换到 merge 视图。
   *
   * 如果是通过 git/Fork CLI 启动，则在 useEffect 中自动调用，无需用户再次点按钮。
   */
  const handleOpenThreeWay = useCallback(async () => {
    const result: ThreeWayOpenResult | null = await window.excelAPI.openThreeWay();
    if (!result) return;

    setMode('merge');
    setSelectedSingleCell(null);
    const allMergeSheets =
      result.sheets && result.sheets.length > 0
        ? result.sheets
        : result.sheet
          ? [result.sheet]
          : [];

    setMergeSheets(allMergeSheets);
    setSelectedMergeSheetIndex(0);
    setMergeCells(allMergeSheets[0]?.cells ?? []);
    setMergeRowsMeta(allMergeSheets[0]?.rowsMeta ?? []);
    setPrimaryKeyCol(1);
    setAutoHasPrimaryKey(true);
    setLastPrimaryKeyCol(1);
    setPrimaryKeyHint('');
    setRowSimilarityThreshold(0.9);
    setResolvedBySheet(new Map());
    setMergeInfo({
      basePath: result.basePath,
      oursPath: result.oursPath,
      theirsPath: result.theirsPath,
      sheetName: allMergeSheets[0]?.sheetName ?? result.sheet?.sheetName ?? '',
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

  // 当主键列设置变化时，重新计算三方 diff（避免重开文件）
  useEffect(() => {
    if (mode !== 'merge' || !mergeInfo) return;
    let cancelled = false;
    (async () => {
      const req: ThreeWayDiffRequest = {
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        primaryKeyCol,
        frozenRowCount: mergeFrozenRowCount,
        rowSimilarityThreshold,
      };
      const result = await window.excelAPI.computeThreeWayDiff(req);
      if (!result || cancelled) return;
      const allMergeSheets =
        result.sheets && result.sheets.length > 0
          ? result.sheets
          : result.sheet
            ? [result.sheet]
            : [];
      const nextIndex = Math.min(selectedMergeSheetIndex, Math.max(0, allMergeSheets.length - 1));
      setMergeSheets(allMergeSheets);
      setSelectedMergeSheetIndex(nextIndex);
      setMergeCells(allMergeSheets[nextIndex]?.cells ?? []);
      setMergeRowsMeta(allMergeSheets[nextIndex]?.rowsMeta ?? []);
      setResolvedBySheet(new Map());
      setSelectedMergeCell(null);
      setSelectedThreeWayRow(null);
    })();
    return () => {
      cancelled = true;
    };
  }, [
    primaryKeyCol,
    mergeFrozenRowCount,
    rowSimilarityThreshold,
    mergeInfo?.basePath,
    mergeInfo?.oursPath,
    mergeInfo?.theirsPath,
    mode,
  ]);

  // 自动判断是否存在主键（基于 rowsMeta 的 key 覆盖率，带滞回避免抖动）
  useEffect(() => {
    if (mode !== 'merge') return;
    if (!mergeRowsMeta || mergeRowsMeta.length === 0) return;
    const nonEmptyKeyCount = mergeRowsMeta.filter((m) => m.key != null && String(m.key).trim() !== '').length;
    const ratio = nonEmptyKeyCount / mergeRowsMeta.length;

    let nextHas = autoHasPrimaryKey;
    if (autoHasPrimaryKey) {
      if (ratio < 0.4) nextHas = false;
    } else {
      if (ratio > 0.6) nextHas = true;
    }

    if (nextHas) {
      setPrimaryKeyHint('自动识别：有主键');
      if (primaryKeyCol === -1) {
        const restored = Math.max(1, Math.floor(lastPrimaryKeyCol || 1));
        setPrimaryKeyCol(restored);
      }
      if (primaryKeyCol >= 1) {
        setLastPrimaryKeyCol(primaryKeyCol);
      }
    } else {
      setPrimaryKeyHint('自动识别：无主键（主键列空值较多）');
      if (primaryKeyCol !== -1) setPrimaryKeyCol(-1);
    }

    if (nextHas !== autoHasPrimaryKey) {
      setAutoHasPrimaryKey(nextHas);
    }
  }, [mode, mergeRowsMeta, autoHasPrimaryKey, primaryKeyCol, lastPrimaryKeyCol]);

  /**
   * 单文件编辑模式下，当用户修改某个输入框时：
   * - 更新内存中的 rows；
   * - 在 changes Map 中记录此单元格修改，供后续一次性保存。
   */
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

  /**
   * 将单文件编辑模式下所有修改过的单元格一次性写回原 Excel。
   */
  const handleSave = useCallback(async () => {
    if (!filePath || changes.size === 0) return;
    setSaving(true);
    try {
      const changeList = Array.from(changes.values());
      await window.excelAPI.saveChanges(changeList);
      setChanges(new Map());
      // 不需要刷新格式，只要值正确写回即可
    } catch (e) {
      alert(`保存失败：${(e as any)?.message ?? String(e)}`);
    } finally {
      setSaving(false);
    }
  }, [changes, filePath]);

  const hasData = useMemo(() => rows.length > 0, [rows]);
  const hasMergeData = useMemo(() => mergeCells.length > 0, [mergeCells]);
  const mergeCellKeySet = useMemo(
    () => new Set(mergeCells.map((c) => `${c.row}:${c.col}`)),
    [mergeCells],
  );

  // 顶部“公式栏”当前要展示的单元格信息
  const selectedMergeCellData = useMemo(() => {
    if (mode !== 'merge' || !selectedMergeCell) return null;
    const rowNumber = selectedMergeCell.rowIndex + 1;
    const colNumber = selectedMergeCell.colIndex + 1;
    const key = `${rowNumber}:${colNumber}`;
    const hit = mergeCells.find((c) => `${c.row}:${c.col}` === key);
    if (hit) return hit;
    const keyCol =
      typeof primaryKeyCol === 'number' && primaryKeyCol >= 1 ? Math.floor(primaryKeyCol) : -1;
    if (keyCol > 0 && colNumber === keyCol) {
      const meta = mergeRowsMeta.find((m) => m.visualRowNumber === rowNumber);
      if (!meta) return null;
      const value = meta.key ?? null;
      const addressRow = meta.oursRowNumber ?? meta.baseRowNumber ?? meta.theirsRowNumber ?? rowNumber;
      return {
        address: makeAddress(colNumber, addressRow),
        row: rowNumber,
        col: colNumber,
        baseValue: value,
        oursValue: value,
        theirsValue: value,
        status: 'unchanged',
        mergedValue: value,
      };
    }
    return null;
  }, [mode, selectedMergeCell, mergeCells, primaryKeyCol, mergeRowsMeta]);

  const selectedMergeRowMeta = useMemo(() => {
    if (mode !== 'merge' || !selectedMergeCell) return null;
    const visualRowNumber = selectedMergeCell.rowIndex + 1;
    return mergeRowsMeta.find((m) => m.visualRowNumber === visualRowNumber) ?? null;
  }, [mode, selectedMergeCell, mergeRowsMeta]);

  const mergedPath = useMemo(() => {
    if (!mergeInfo) return null;
    if (cliInfo?.mode === 'merge') {
      return cliInfo.mergedPath ?? mergeInfo.oursPath;
    }
    if (cliInfo?.mode === 'diff') {
      return mergeInfo.oursPath;
    }
    return null;
  }, [mergeInfo, cliInfo]);

  // 当选中单元格变化时，按需读取该“整行”的 base/ours/theirs 值，用于底部行级对比视图
  useEffect(() => {
    let cancelled = false;
    (async () => {
      if (mode !== 'merge' || !mergeInfo || !selectedMergeCell) {
        setSelectedThreeWayRow(null);
        return;
      }

      const visualRowNumber = selectedMergeCell.rowIndex + 1;
      const rowMeta = mergeRowsMeta.find((m) => m.visualRowNumber === visualRowNumber);
      const result = await window.excelAPI.getThreeWayRow({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        rowNumber: visualRowNumber,
        baseRowNumber: rowMeta?.baseRowNumber ?? null,
        oursRowNumber: rowMeta?.oursRowNumber ?? null,
        theirsRowNumber: rowMeta?.theirsRowNumber ?? null,
      });

      if (cancelled) return;
      setSelectedThreeWayRow(result);
    })();

    return () => {
      cancelled = true;
    };
  }, [mode, mergeInfo, selectedMergeCell, selectedMergeSheetIndex, mergeRowsMeta]);

  // 顶部“公式栏”当前要展示的单元格坐标和值（single / merge 共用）
  let currentCellAddress = '';
  let currentCellValue = '';

  if (mode === 'single' && selectedSingleCell) {
    currentCellAddress = selectedSingleCell.address;
    currentCellValue = selectedSingleCell.value === null ? '' : String(selectedSingleCell.value);
  } else if (mode === 'merge' && selectedMergeCellData) {
    currentCellAddress = selectedMergeCellData.address;
    // merge 模式下不再用一个“当前值”展示；此字段保留给 single 模式
    currentCellValue = '';
  }

  const handleSelectMergeCell = useCallback((rowIndex: number, colIndex: number) => {
    setSelectedMergeCell({ rowIndex, colIndex });
  }, []);

  /**
   * merge 模式下，在右侧详情中点击“用 base / ours / theirs”按钮时：
   * - 更新 mergeSheets 中对应单元格的 mergedValue；
   * - 同步更新当前正在展示的 mergeRows；
   *   这样列表与详情都能立即反映最新选择。
   */
  const markResolvedKeys = useCallback(
    (sheetIndex: number, keys: string[]) => {
      if (keys.length === 0) return;
      setResolvedBySheet((prev) => {
        const next = new Map(prev);
        const current = next.get(sheetIndex) ?? new Set<string>();
        const merged = new Set(current);
        keys.forEach((k) => merged.add(k));
        next.set(sheetIndex, merged);
        return next;
      });
    },
    [],
  );

  const handleApplyMergeChoice = useCallback(
    (source: 'base' | 'ours' | 'theirs') => {
      if (!selectedMergeCell) return;

      const { rowIndex, colIndex } = selectedMergeCell;
      const key = `${rowIndex + 1}:${colIndex + 1}`;
      if (!mergeCellKeySet.has(key)) return;
      // 只标记用户显式操作过的单元格
      markResolvedKeys(selectedMergeSheetIndex, [key]);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.row - 1 !== rowIndex || cell.col - 1 !== colIndex) return cell;
            let value: string | number | null;
            if (source === 'base') value = cell.baseValue;
            else if (source === 'ours') value = cell.oursValue;
            else value = cell.theirsValue;
            return { ...cell, mergedValue: value };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      // 同步当前视图的 cells
      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.row - 1 !== rowIndex || cell.col - 1 !== colIndex) return cell;
          let value: string | number | null;
          if (source === 'base') value = cell.baseValue;
          else if (source === 'ours') value = cell.oursValue;
          else value = cell.theirsValue;
          return { ...cell, mergedValue: value };
        }),
      );
    },
    [selectedMergeCell, selectedMergeSheetIndex, markResolvedKeys, mergeCellKeySet],
  );

  const handleApplyMergeRowChoice = useCallback(
    (rowNumber: number, source: 'ours' | 'theirs') => {
      const valueFrom = (cell: MergeCell) => (source === 'ours' ? cell.oursValue : cell.theirsValue);

      // 标记这一行所有差异单元格为 resolved
      const keys = mergeCells
        .filter((c) => c.row === rowNumber)
        .map((c) => `${c.row}:${c.col}`);
      markResolvedKeys(selectedMergeSheetIndex, keys);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.row !== rowNumber) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      // 同步当前视图的 cells
      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.row !== rowNumber) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );
    },
    [selectedMergeSheetIndex, mergeCells, markResolvedKeys],
  );

  const handleApplyMergeCellChoice = useCallback(
    (rowNumber: number, colNumber: number, source: 'ours' | 'theirs') => {
      const valueFrom = (cell: MergeCell) => (source === 'ours' ? cell.oursValue : cell.theirsValue);
      const key = `${rowNumber}:${colNumber}`;
      if (!mergeCellKeySet.has(key)) return;

      markResolvedKeys(selectedMergeSheetIndex, [`${rowNumber}:${colNumber}`]);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.row !== rowNumber || cell.col !== colNumber) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.row !== rowNumber || cell.col !== colNumber) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );
    },
    [selectedMergeSheetIndex, markResolvedKeys, mergeCellKeySet],
  );

  const handleApplyMergeCellsChoice = useCallback(
    (keys: { rowNumber: number; colNumber: number }[], source: 'ours' | 'theirs') => {
      if (!keys.length) return;
      const valueFrom = (cell: MergeCell) => (source === 'ours' ? cell.oursValue : cell.theirsValue);
      const filtered = keys.filter((k) => mergeCellKeySet.has(`${k.rowNumber}:${k.colNumber}`));
      if (!filtered.length) return;
      const keySet = new Set(filtered.map((k) => `${k.rowNumber}:${k.colNumber}`));
      markResolvedKeys(selectedMergeSheetIndex, Array.from(keySet));

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            const k = `${cell.row}:${cell.col}`;
            if (!keySet.has(k)) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      setMergeCells((prev) =>
        prev.map((cell) => {
          const k = `${cell.row}:${cell.col}`;
          if (!keySet.has(k)) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );
    },
    [selectedMergeSheetIndex, markResolvedKeys, mergeCellKeySet],
  );

  /**
   * merge 模式下，将所有工作表的 mergedValue 写回一个目标 Excel 文件。
   *
   * 为了避免误操作，这里会先统计所有发生变化的单元格，
   * 构造一个预览字符串通过 window.confirm 让用户二次确认。
   */
  const handleSaveMergeToFile = useCallback(async () => {
    if (!mergeInfo || mergeSheets.length === 0) return;

    // 生成本次合并的概要信息：mergeSheets.cells 本身就是差异单元格列表
    const changedCells: { sheetName: string; address: string; ours: any; theirs: any; merged: any }[] = [];
    let skippedCells = 0;
    mergeSheets.forEach((sheet) => {
      const rowMetaMap = new Map<number, MergeRowMeta>();
      (sheet.rowsMeta ?? []).forEach((m) => rowMetaMap.set(m.visualRowNumber, m));
      const hasRowMeta = (sheet.rowsMeta ?? []).length > 0;
      sheet.cells.forEach((cell: MergeCell) => {
        const meta = rowMetaMap.get(cell.row);
        const targetRowNumber = meta?.oursRowNumber ?? null;
        if (hasRowMeta && !targetRowNumber) {
          skippedCells += 1;
          return;
        }
        const address = targetRowNumber ? makeAddress(cell.col, targetRowNumber) : cell.address;
        changedCells.push({
          sheetName: sheet.sheetName,
          address,
          ours: cell.oursValue,
          theirs: cell.theirsValue,
          merged: cell.mergedValue,
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
    if (skippedCells > 0) {
      lines.push(`（提示：有 ${skippedCells} 个单元格因 ours 侧缺少对应行而未写入）`);
    }

    const preview =
      `本次合并将影响 ${changedCells.length} 个单元格（覆盖所有工作表）：` +
      (lines.length ? `\n\n${lines.join('\n')}` : '\n(无差异单元格——仅写回了当前值)') +
      '\n\n注意：保存时会将所有工作表的合并结果一并写入目标 Excel 文件。' +
      '\n\n确认要将以上结果写入 Excel 文件吗？';

    const confirmed = window.confirm(preview);
    if (!confirmed) return;

    const cells = mergeSheets.flatMap((sheet: MergeSheetData) => {
      const rowMetaMap = new Map<number, MergeRowMeta>();
      (sheet.rowsMeta ?? []).forEach((m) => rowMetaMap.set(m.visualRowNumber, m));
      const hasRowMeta = (sheet.rowsMeta ?? []).length > 0;
      return sheet.cells
        .map((cell: MergeCell) => {
          const meta = rowMetaMap.get(cell.row);
          const targetRowNumber = meta?.oursRowNumber ?? null;
          if (hasRowMeta && !targetRowNumber) return null;
          const address = targetRowNumber ? makeAddress(cell.col, targetRowNumber) : cell.address;
          return {
            sheetName: sheet.sheetName,
            address,
            value: cell.mergedValue,
          };
        })
        .filter(Boolean) as { sheetName: string; address: string; value: string | number | null }[];
    });

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
  }, [mergeInfo, mergeSheets]);

  return (
    <div
      style={{
        padding: 16,
        fontFamily: 'sans-serif',
        height: '100vh',
        boxSizing: 'border-box',
        display: 'flex',
        flexDirection: 'column',
        overflow: 'hidden',
      }}
    >
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


      {/* 主内容：表格 / 三方 Merge，占用剩余空间，由内部自己滚动 */}
      <div
        style={{
          flex: 1,
          minHeight: 0,
          overflow: 'hidden',
          display: 'flex',
          flexDirection: 'column',
        }}
      >

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
                      setSelectedSingleCell(null);
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
          <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 4 }}>
            <span>固定首行数:</span>
            <input
              type="number"
              min={0}
              value={frozenRowCount}
              onChange={(e) => {
                const v = Number(e.target.value);
                if (Number.isNaN(v)) return;
                setFrozenRowCount(Math.max(0, Math.floor(v)));
              }}
              style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
            />
            <span style={{ fontSize: 12, color: '#666' }}>（例如 3 表示固定前 3 行）</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 4 }}>
            <span>固定首列数:</span>
            <input
              type="number"
              min={0}
              value={frozenColCount}
              onChange={(e) => {
                const v = Number(e.target.value);
                if (Number.isNaN(v)) return;
                setFrozenColCount(Math.max(0, Math.floor(v)));
              }}
              style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
            />
            <span style={{ fontSize: 12, color: '#666' }}>（例如 1 表示固定 A 列）</span>
          </div>

          {/* 公式栏：移到文件/工作表信息下方 */}
          <div
            style={{
              display: 'flex',
              alignItems: 'flex-start',
              gap: 12,
              marginTop: 8,
              flexWrap: 'wrap',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              <span style={{ fontSize: 12 }}>单元格地址</span>
              <input
                readOnly
                value={currentCellAddress}
                placeholder="例如 A1"
                style={{ width: 90, padding: '2px 6px', boxSizing: 'border-box' }}
              />
            </div>
            <div style={{ display: 'flex', flex: 1, alignItems: 'flex-start', gap: 4 }}>
              <span style={{ fontSize: 12, whiteSpace: 'nowrap' }}>当前值</span>
              <textarea
                readOnly
                value={currentCellValue}
                placeholder="当前单元格值"
                rows={1}
                style={{
                  flex: 1,
                  minWidth: 260,
                  maxWidth: '100%',
                  padding: '2px 6px',
                  boxSizing: 'border-box',
                  height: 24,
                  resize: 'none',
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-all',
                }}
              />
            </div>
          </div>
        </div>
      )}

      {mode === 'single' && (
        hasData ? (
          <div style={{ flex: 1, minHeight: 0 }}>
            <ExcelTable
              rows={rows}
              onCellChange={handleCellChange}
              onCellSelect={setSelectedSingleCell}
              selectedAddress={selectedSingleCell?.address ?? null}
              frozenRowCount={frozenRowCount}
              frozenColCount={frozenColCount}
            />
          </div>
        ) : (
          <div>请先打开一个 .xlsx 文件。</div>
        )
      )}

      {mode === 'merge' && (
        mergeInfo && mergeSheets.length === 0 ? (
          <div>
            没有可对比的工作表（base / ours / theirs 中没有任何“同名工作表”的交集）。
          </div>
        ) : mergeInfo ? (
          <div style={{ flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
            <div style={{ marginBottom: 8 }}>
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
                    const hasDiff =
                      typeof s.hasExactDiff === 'boolean' ? s.hasExactDiff : (s.cells?.length ?? 0) > 0;
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
                          setMergeCells(sheet?.cells ?? []);
                          setMergeRowsMeta(sheet?.rowsMeta ?? []);
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
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: 6,
                        }}
                      >
                        {hasDiff && (
                          <span
                            title="该工作表有内容变动"
                            style={{
                              width: 8,
                              height: 8,
                              backgroundColor: '#d32f2f',
                              borderRadius: 2,
                              display: 'inline-block',
                            }}
                          />
                        )}
                        {s.sheetName || `Sheet${idx + 1}`}
                      </button>
                    );
                  })}
                </div>
              </div>
              <div style={{ marginTop: 4, fontSize: 12, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <span>颜色说明（只比较单元格值，忽略格式）：</span>
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ width: 10, height: 10, backgroundColor: '#d4f8d4', border: '1px solid #bbb', display: 'inline-block' }} />
                  <span>ours 侧：ours 有改动 / 冲突时 ours</span>
                </span>
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ width: 10, height: 10, backgroundColor: '#ffc8c8', border: '1px solid #bbb', display: 'inline-block' }} />
                  <span>theirs 侧：theirs 有改动 / 冲突时 theirs</span>
                </span>
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ width: 10, height: 10, backgroundColor: '#fafafa', border: '1px solid #bbb', display: 'inline-block' }} />
                  <span>浅灰：双方都改且改成相同值 / 已合并</span>
                </span>
                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                  <span style={{ width: 10, height: 10, backgroundColor: '#ffffff', border: '1px solid #bbb', display: 'inline-block' }} />
                  <span>白色：无差异</span>
                </span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 4 }}>
                <span>merge/diff 冻结行数:</span>
                <input
                  type="number"
                  min={0}
                  value={mergeFrozenRowCount}
                  onChange={(e) => {
                    const v = Number(e.target.value);
                    if (Number.isNaN(v)) return;
                    setMergeFrozenRowCount(Math.max(0, Math.floor(v)));
                  }}
                  style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
                />
                <span style={{ fontSize: 12, color: '#666' }}>（例如 3 表示固定前 3 行）</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 4 }}>
                <span>行相似度阈值:</span>
                <input
                  type="number"
                  min={0}
                  max={1}
                  step={0.01}
                  value={rowSimilarityThreshold}
                  onChange={(e) => {
                    const v = Number(e.target.value);
                    if (Number.isNaN(v)) return;
                    setRowSimilarityThreshold(Math.min(1, Math.max(0, v)));
                  }}
                  style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
                />
                <span style={{ fontSize: 12, color: '#666' }}>（0~1，越大越严格）</span>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 4 }}>
                {autoHasPrimaryKey && (
                  <>
                    <span>主键列:</span>
                    <input
                      type="number"
                      min={1}
                      value={primaryKeyCol}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        if (Number.isNaN(v)) return;
                        const next = Math.max(1, Math.floor(v));
                        setPrimaryKeyCol(next);
                        setLastPrimaryKeyCol(next);
                      }}
                      style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
                    />
                    <span style={{ fontSize: 12, color: '#666' }}>（1=A 列，2=B 列…）</span>
                  </>
                )}
                {!autoHasPrimaryKey && (
                  <span style={{ fontSize: 12, color: '#666' }}>（无主键：使用序列对齐算法）</span>
                )}
                {primaryKeyHint && (
                  <span style={{ fontSize: 12, color: '#b00020' }}>{primaryKeyHint}</span>
                )}
              </div>

              {/* 公式栏：移到路径/工作表信息下方 */}
              <div
                style={{
                  display: 'flex',
                  alignItems: 'flex-start',
                  gap: 12,
                  marginTop: 8,
                  flexWrap: 'wrap',
                }}
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span style={{ fontSize: 12 }}>单元格地址</span>
                  <input
                    readOnly
                    value={currentCellAddress}
                    placeholder="例如 A1"
                    style={{ width: 90, padding: '2px 6px', boxSizing: 'border-box' }}
                  />
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span style={{ fontSize: 12, whiteSpace: 'nowrap' }}>base</span>
                  <input
                    readOnly
                    value={selectedMergeCellData?.baseValue == null ? '' : String(selectedMergeCellData.baseValue)}
                    style={{ width: 220, padding: '2px 6px', boxSizing: 'border-box' }}
                  />
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span style={{ fontSize: 12, whiteSpace: 'nowrap' }}>ours</span>
                  <input
                    readOnly
                    value={selectedMergeCellData?.oursValue == null ? '' : String(selectedMergeCellData.oursValue)}
                    style={{ width: 220, padding: '2px 6px', boxSizing: 'border-box' }}
                  />
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span style={{ fontSize: 12, whiteSpace: 'nowrap' }}>theirs</span>
                  <input
                    readOnly
                    value={selectedMergeCellData?.theirsValue == null ? '' : String(selectedMergeCellData.theirsValue)}
                    style={{ width: 220, padding: '2px 6px', boxSizing: 'border-box' }}
                  />
                </div>
              </div>
            </div>
            <div style={{ flex: 1, minHeight: 0 }}>
                <MergeSideBySide
                  cells={mergeCells}
                  rowsMeta={mergeRowsMeta}
                  selected={selectedMergeCell}
                  onSelectCell={handleSelectMergeCell}
                  onApplyRowChoice={handleApplyMergeRowChoice}
                  onApplyCellChoice={handleApplyMergeCellChoice}
                  onApplyCellsChoice={handleApplyMergeCellsChoice}
                  resolvedCellKeys={resolvedBySheet.get(selectedMergeSheetIndex)}
                  frozenRowCount={mergeFrozenRowCount}
                  primaryKeyCol={primaryKeyCol}
                  oursPath={mergeInfo?.oursPath ?? null}
                  basePath={mergeInfo?.basePath ?? null}
                  theirsPath={mergeInfo?.theirsPath ?? null}
                />
            </div>
            {selectedMergeCell && selectedThreeWayRow && (
              <div style={{ marginTop: 8, border: '1px solid #ccc', overflow: 'hidden' }}>
                <div
                  style={{
                    padding: 6,
                    borderBottom: '1px solid #eee',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 8,
                    flexWrap: 'wrap',
                    fontSize: 12,
                    backgroundColor: '#fafafa',
                  }}
                >
                  {mergedPath ? (
                    <span
                      title={mergedPath}
                      style={{ maxWidth: 520, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                    >
                      merged: {mergedPath}
                    </span>
                  ) : (
                    <span>没有设置 merged 路径</span>
                  )}
                </div>

                {(() => {
                  const colCount = selectedThreeWayRow.colCount;
                  const selectedColIndex = selectedMergeCell.colIndex; // 0-based
                  const visualRowNumber = selectedMergeCell.rowIndex + 1;

                  // 基于 ours 行作为模板，覆盖本行所有 diff cell 的 mergedValue
                  const mergedRowValues = Array.from({ length: colCount }, (_v, i) =>
                    selectedThreeWayRow.ours[i] ?? null,
                  );
                  mergeCells.forEach((c) => {
                    if (c.row === visualRowNumber && c.col >= 1 && c.col <= colCount) {
                      mergedRowValues[c.col - 1] = c.mergedValue ?? null;
                    }
                  });

                  const renderValueCell = (v: string | number | null) => (
                    <div
                      title={v == null ? '' : String(v)}
                      style={{
                        width: '100%',
                        height: '100%',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap',
                      }}
                    >
                      {v == null ? '' : String(v)}
                    </div>
                  );

                  const getCellStyle = (_v: any, ctx: any): React.CSSProperties => {
                    const style: React.CSSProperties = {};
                    if (ctx.colIndex === selectedColIndex) {
                      style.border = '2px solid #ff8000';
                    }
                    return style;
                  };

                  const scrollToCell = { rowIndex: 0, colIndex: selectedColIndex };

                  return (
                    <div style={{ padding: 6 }}>
                      <div style={{ height: 84 }}>
                        <VirtualGrid<(string | number | null)>
                          rows={[mergedRowValues]}
                          showRowHeader
                          renderRowHeader={() => visualRowNumber}
                          renderCell={(cell) => renderValueCell(cell)}
                          getCellStyle={getCellStyle}
                          frozenRowCount={0}
                          frozenColCount={0}
                          columnWidths={mergedRowColumnWidths}
                          onColumnWidthsChange={setMergedRowColumnWidths}
                          scrollToCell={scrollToCell}
                        />
                      </div>
                    </div>
                  );
                })()}
              </div>
            )}
          </div>
        ) : (
          <div>请先选择 base / ours / theirs 三个 Excel 文件。</div>
        )
      )}
      </div>
    </div>
  );
};

