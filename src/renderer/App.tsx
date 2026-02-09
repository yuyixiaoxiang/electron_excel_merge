import React, { useCallback, useEffect, useMemo, useState } from 'react';
import type {
  CellChange,
  CliThreeWayInfo,
  MergeCell,
  MergeColumnMeta,
  MergeRowMeta,
  MergeSheetData,
  OpenResult,
  SaveMergeColOp,
  SaveMergeRowOp,
  SaveMergeRequest,
  SheetCell,
  SheetData,
  ThreeWayDiffRequest,
  ThreeWayOpenResult,
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
  const [mergeColumnsMeta, setMergeColumnsMeta] = useState<MergeColumnMeta[]>([]);
  const [primaryKeyCol, setPrimaryKeyCol] = useState<number>(1);
  const [autoHasPrimaryKey, setAutoHasPrimaryKey] = useState<boolean>(true);
  const [lastPrimaryKeyCol, setLastPrimaryKeyCol] = useState<number>(1);
  const [primaryKeyHint, setPrimaryKeyHint] = useState<string>('');
  // 记录“用户已确认合并”的单元格（resolved），按 sheetIndex 分组，key="row:col"（1-based）
  const [resolvedBySheet, setResolvedBySheet] = useState<Map<number, Set<string>>>(new Map());
  const [mergeRowOpsBySheet, setMergeRowOpsBySheet] = useState<Map<number, Map<number, SaveMergeRowOp>>>(new Map());
  const [mergeColOpsBySheet, setMergeColOpsBySheet] = useState<Map<number, Map<number, SaveMergeColOp>>>(new Map());
  const [mergedPreviewMinRows, setMergedPreviewMinRows] = useState<number>(5);
  const [mergedPreviewRows, setMergedPreviewRows] = useState<(string | number | null)[][]>([]);
  const [mergedPreviewRowVisuals, setMergedPreviewRowVisuals] = useState<(number | null)[]>([]);
  const [showFullTables, setShowFullTables] = useState<boolean>(false);
  const [fullOursRows, setFullOursRows] = useState<(string | number | null)[][]>([]);
  const [fullTheirsRows, setFullTheirsRows] = useState<(string | number | null)[][]>([]);
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
  const displayPrimaryKeyCol = useMemo(() => {
    if (typeof primaryKeyCol !== 'number' || primaryKeyCol < 1) return primaryKeyCol;
    const hit = mergeColumnsMeta.find((c) => c.oursCol === primaryKeyCol);
    return hit ? hit.col : primaryKeyCol;
  }, [primaryKeyCol, mergeColumnsMeta]);

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
    setMergeColumnsMeta(allMergeSheets[0]?.columnsMeta ?? []);
    setPrimaryKeyCol(1);
    setAutoHasPrimaryKey(true);
    setLastPrimaryKeyCol(1);
    setPrimaryKeyHint('');
    setRowSimilarityThreshold(0.9);
    setResolvedBySheet(new Map());
    setMergeRowOpsBySheet(new Map());
    setMergeColOpsBySheet(new Map());
    setMergedPreviewRows([]);
    setMergedPreviewRowVisuals([]);
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
      setMergeColumnsMeta(allMergeSheets[nextIndex]?.columnsMeta ?? []);
      setResolvedBySheet(new Map());
      setMergeRowOpsBySheet(new Map());
      setMergeColOpsBySheet(new Map());
      setMergedPreviewRows([]);
      setMergedPreviewRowVisuals([]);
      setSelectedMergeCell(null);
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
      await window.excelAPI.saveChanges({
        changes: changeList,
        sheetName: sheetName ?? undefined,
        sheetIndex: selectedSheetIndex,
      });
      setChanges(new Map());
      // 不需要刷新格式，只要值正确写回即可
    } catch (e) {
      alert(`保存失败：${(e as any)?.message ?? String(e)}`);
    } finally {
      setSaving(false);
    }
  }, [changes, filePath, sheetName, selectedSheetIndex]);

  const hasData = useMemo(() => rows.length > 0, [rows]);
  const hasMergeData = useMemo(() => mergeCells.length > 0, [mergeCells]);
  const mergeCellKeySet = useMemo(
    () => new Set(mergeCells.map((c) => `${c.row}:${c.col}`)),
    [mergeCells],
  );
  useEffect(() => {
    if (mode !== 'merge' || !mergeInfo || !showFullTables) {
      setFullOursRows([]);
      setFullTheirsRows([]);
      return;
    }
    let cancelled = false;
    (async () => {
      const [oursSheet, theirsSheet] = await Promise.all([
        window.excelAPI.getSheetData({
          path: mergeInfo.oursPath,
          sheetName: mergeInfo.sheetName,
          sheetIndex: selectedMergeSheetIndex,
        }),
        window.excelAPI.getSheetData({
          path: mergeInfo.theirsPath,
          sheetName: mergeInfo.sheetName,
          sheetIndex: selectedMergeSheetIndex,
        }),
      ]);
      if (cancelled) return;
      const oursRows = (oursSheet?.rows ?? []).map((row: SheetCell[]) =>
        row.map((c: SheetCell) => c.value ?? null),
      );
      const theirsRows = (theirsSheet?.rows ?? []).map((row: SheetCell[]) =>
        row.map((c: SheetCell) => c.value ?? null),
      );
      setFullOursRows(oursRows);
      setFullTheirsRows(theirsRows);
    })();
    return () => {
      cancelled = true;
    };
  }, [
    mode,
    showFullTables,
    mergeInfo?.oursPath,
    mergeInfo?.theirsPath,
    mergeInfo?.sheetName,
    selectedMergeSheetIndex,
  ]);

  const mergeCellsByRow = useMemo(() => {
    const m = new Map<number, MergeCell[]>();
    mergeCells.forEach((cell) => {
      if (!m.has(cell.row)) m.set(cell.row, []);
      m.get(cell.row)!.push(cell);
    });
    return m;
  }, [mergeCells]);

  // 顶部"公式栏"当前要展示的单元格信息
  const selectedMergeCellData = useMemo(() => {
    if (mode !== 'merge' || !selectedMergeCell) return null;
    const rowNumber = selectedMergeCell.rowIndex + 1;
    const colNumber = selectedMergeCell.colIndex + 1;
    const rowCells = mergeCellsByRow.get(rowNumber);
    const hit = rowCells?.find((c) => c.col === colNumber) ?? null;
    if (hit) return hit;
    const keyCol =
      typeof displayPrimaryKeyCol === 'number' && displayPrimaryKeyCol >= 1
        ? Math.floor(displayPrimaryKeyCol)
        : -1;
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
  }, [mode, selectedMergeCell, mergeCellsByRow, displayPrimaryKeyCol, mergeRowsMeta]);


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
  const currentRowOps = useMemo(
    () => mergeRowOpsBySheet.get(selectedMergeSheetIndex) ?? new Map<number, SaveMergeRowOp>(),
    [mergeRowOpsBySheet, selectedMergeSheetIndex],
  );
  const currentColOps = useMemo(
    () => mergeColOpsBySheet.get(selectedMergeSheetIndex) ?? new Map<number, SaveMergeColOp>(),
    [mergeColOpsBySheet, selectedMergeSheetIndex],
  );
  useEffect(() => {
    if (mode !== 'merge' || !mergeInfo) {
      setMergedPreviewRows([]);
      setMergedPreviewRowVisuals([]);
      return;
    }
    let cancelled = false;
    (async () => {
      const metas = [...mergeRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      const minRows = Math.max(1, Math.floor(mergedPreviewMinRows));
      if (metas.length === 0) {
        if (!cancelled) {
          setMergedPreviewRows(Array.from({ length: minRows }, () => []));
          setMergedPreviewRowVisuals(Array.from({ length: minRows }, () => null));
        }
        return;
      }
      const rowsReq = metas.map((m) => ({
        rowNumber: m.visualRowNumber,
        baseRowNumber: m.baseRowNumber,
        oursRowNumber: m.oursRowNumber,
        theirsRowNumber: m.theirsRowNumber,
      }));
      const result = await window.excelAPI.getThreeWayRows({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        frozenRowCount: mergeFrozenRowCount,
        rows: rowsReq,
      });
      if (!result || cancelled) return;
      const rawColCount = result.colCount ?? 0;
      // Build effective column list considering col ops
      const deletedAlignedCols = new Set<number>();
      const insertedAlignedCols: number[] = [];
      currentColOps.forEach((op, alignedCol) => {
        if (op.action === 'delete') deletedAlignedCols.add(alignedCol);
        else if (op.action === 'insert') insertedAlignedCols.push(alignedCol);
      });
      // Map aligned col -> ours col for non-deleted columns
      // IMPORTANT: 只包含 ours 模板中存在的列（oursCol 非空），
      // theirs-only 列只有在用户显式选择"插入"后才通过下方 insert 逻辑加入，
      // 否则 merged 预览中会出现重复列。
      const effectiveColMap: { alignedCol: number; oursCol: number | null }[] = [];
      for (let c = 1; c <= rawColCount; c += 1) {
        if (deletedAlignedCols.has(c)) continue;
        const meta = mergeColumnsMeta.find((m) => m.col === c);
        if (!meta?.oursCol) continue;
        effectiveColMap.push({ alignedCol: c, oursCol: meta.oursCol });
      }
      // Add inserted columns (theirs-only)
      insertedAlignedCols.sort((a, b) => a - b);
      // IMPORTANT: 先收集所有插入位置，然后从后往前插入，避免索引错乱
      const insertions: Array<{ idx: number; col: number }> = [];
      for (const ac of insertedAlignedCols) {
        const meta = mergeColumnsMeta.find((m) => m.col === ac);
        if (meta && !meta.oursCol && meta.theirsCol) {
          // Find insertion position
          let insertIdx = effectiveColMap.length;
          for (let i = 0; i < effectiveColMap.length; i += 1) {
            if (effectiveColMap[i].alignedCol > ac) {
              insertIdx = i;
              break;
            }
          }
          insertions.push({ idx: insertIdx, col: ac });
        }
      }
      // 从后往前插入，避免每次 splice 改变后续索引
      // 同一位置的多个插入按 col 降序处理，确保最终顺序正确
      insertions.sort((a, b) => b.idx - a.idx || b.col - a.col);
      for (const ins of insertions) {
        effectiveColMap.splice(ins.idx, 0, { alignedCol: ins.col, oursCol: null });
      }
      const colCount = effectiveColMap.length;
      const mergedRows: (string | number | null)[][] = [];
      const mergedVisuals: (number | null)[] = [];
      result.rows.forEach((rowRes: any, idx: number) => {
        const meta = metas[idx];
        const visualRowNumber = meta?.visualRowNumber ?? rowRes.rowNumber ?? idx + 1;
        const op = currentRowOps.get(visualRowNumber);
        const oursMissing = !meta?.oursRowNumber;
        if (oursMissing && op?.action !== 'insert') return;
        if (!oursMissing && op?.action === 'delete') return;
        // Build merged row based on effective columns
        const mergedRow: (string | number | null)[] = [];
        for (let i = 0; i < effectiveColMap.length; i += 1) {
          const colInfo = effectiveColMap[i];
          const alignedCol = colInfo.alignedCol;
          const colMeta = mergeColumnsMeta.find((m) => m.col === alignedCol);
          // Check if there's a diff cell override
          const diffCell = (mergeCellsByRow.get(visualRowNumber) ?? []).find((c) => c.col === alignedCol);
          if (diffCell) {
            mergedRow.push(diffCell.mergedValue ?? null);
            continue;
          }
          // Otherwise get from ours/theirs raw data
          if (op?.action === 'insert' && op.values) {
            // IMPORTANT: 用 alignedCol 索引而非循环索引 i——effectiveColMap 会跳过已删除的列和未插入的 theirs-only 列，
            // 但 op.values 始终按原始 aligned 列顺序排列，所以必须用 alignedCol - 1 取值。
            mergedRow.push(op.values[alignedCol - 1] ?? null);
          } else if (colMeta?.oursCol && rowRes.ours) {
            mergedRow.push(rowRes.ours[alignedCol - 1] ?? null);
          } else if (colMeta?.theirsCol && rowRes.theirs) {
            // For theirs-only columns that are being inserted
            mergedRow.push(rowRes.theirs[alignedCol - 1] ?? null);
          } else {
            mergedRow.push(null);
          }
        }
        mergedRows.push(mergedRow);
        mergedVisuals.push(visualRowNumber);
      });
      while (mergedRows.length < minRows) {
        mergedRows.push(Array(colCount).fill(null));
        mergedVisuals.push(null);
      }
      if (cancelled) return;
      setMergedPreviewRows(mergedRows);
      setMergedPreviewRowVisuals(mergedVisuals);
    })();
    return () => {
      cancelled = true;
    };
  }, [
    mode,
    mergeInfo,
    mergeRowsMeta,
    mergeCellsByRow,
    mergedPreviewMinRows,
    mergeFrozenRowCount,
    selectedMergeSheetIndex,
    currentRowOps,
    currentColOps,
    mergeColumnsMeta,
  ]);


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
  const updateRowOpForSheet = useCallback(
    (sheetIndex: number, visualRowNumber: number, op: SaveMergeRowOp | null) => {
      setMergeRowOpsBySheet((prev) => {
        const next = new Map(prev);
        const sheetOps = new Map(next.get(sheetIndex) ?? new Map<number, SaveMergeRowOp>());
        if (op) sheetOps.set(visualRowNumber, op);
        else sheetOps.delete(visualRowNumber);
        if (sheetOps.size === 0) next.delete(sheetIndex);
        else next.set(sheetIndex, sheetOps);
        return next;
      });
    },
    [],
  );
  const updateColOpForSheet = useCallback(
    (sheetIndex: number, alignedColNumber: number, op: SaveMergeColOp | null) => {
      setMergeColOpsBySheet((prev) => {
        const next = new Map(prev);
        const sheetOps = new Map(next.get(sheetIndex) ?? new Map<number, SaveMergeColOp>());
        if (op) sheetOps.set(alignedColNumber, op);
        else sheetOps.delete(alignedColNumber);
        if (sheetOps.size === 0) next.delete(sheetIndex);
        else next.set(sheetIndex, sheetOps);
        return next;
      });
    },
    [],
  );
  const computeInsertTargetColNumber = useCallback(
    (alignedColNumber: number) => {
      if (!mergeColumnsMeta || mergeColumnsMeta.length === 0) return 1;
      const metaMap = new Map<number, MergeColumnMeta>();
      mergeColumnsMeta.forEach((m) => metaMap.set(m.col, m));
      for (let c = alignedColNumber - 1; c >= 1; c -= 1) {
        const meta = metaMap.get(c);
        if (meta?.oursCol) return meta.oursCol + 1;
      }
      for (let c = alignedColNumber + 1; c <= metaMap.size; c += 1) {
        const meta = metaMap.get(c);
        if (meta?.oursCol) return meta.oursCol;
      }
      return 1;
    },
    [mergeColumnsMeta],
  );
  const computeInsertTargetRowNumber = useCallback(
    (visualRowNumber: number) => {
      const list = [...mergeRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      const idx = list.findIndex((m) => m.visualRowNumber === visualRowNumber);
      if (idx < 0) return 1;
      for (let i = idx - 1; i >= 0; i -= 1) {
        const r = list[i].oursRowNumber;
        if (r) return r + 1;
      }
      for (let i = idx + 1; i < list.length; i += 1) {
        const r = list[i].oursRowNumber;
        if (r) return r;
      }
      return 1;
    },
    [mergeRowsMeta],
  );
  const buildMergedRowValues = useCallback(
    async (visualRowNumber: number, rowMeta: MergeRowMeta) => {
      if (!mergeInfo) return null;
      const result = await window.excelAPI.getThreeWayRow({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        frozenRowCount: mergeFrozenRowCount,
        rowNumber: visualRowNumber,
        baseRowNumber: rowMeta.baseRowNumber ?? null,
        oursRowNumber: rowMeta.oursRowNumber ?? null,
        theirsRowNumber: rowMeta.theirsRowNumber ?? null,
      });
      if (!result) return null;
      const colCount = result.colCount;
      let baseRow: (string | number | null)[] = [];
      if (rowMeta.oursRowNumber) baseRow = result.ours;
      else if (rowMeta.baseRowNumber) baseRow = result.base;
      else if (rowMeta.theirsRowNumber) baseRow = result.theirs;
      else baseRow = Array(colCount).fill(null);
      const mergedRow = baseRow.slice(0, colCount);
      if (mergedRow.length < colCount) {
        mergedRow.push(...Array(colCount - mergedRow.length).fill(null));
      }
      const diffCells = mergeCellsByRow.get(visualRowNumber) ?? [];
      diffCells.forEach((cell) => {
        if (cell.col >= 1 && cell.col <= colCount) {
          mergedRow[cell.col - 1] = cell.mergedValue ?? null;
        }
      });
      return mergedRow;
    },
    [mergeInfo, selectedMergeSheetIndex, mergeCellsByRow, mergeFrozenRowCount],
  );

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
    async (rowNumber: number, source: 'ours' | 'theirs') => {
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
      const rowMeta = mergeRowsMeta.find((m) => m.visualRowNumber === rowNumber);
      if (!rowMeta || !mergeInfo) return;
      const oursRowNumber = rowMeta.oursRowNumber ?? null;
      const theirsRowNumber = rowMeta.theirsRowNumber ?? null;
      let op: SaveMergeRowOp | null = null;
      if (!oursRowNumber && theirsRowNumber) {
        if (source === 'theirs') {
          const values = await buildMergedRowValues(rowNumber, rowMeta);
          if (values) {
            op = {
              sheetName: mergeInfo.sheetName,
              action: 'insert',
              targetRowNumber: computeInsertTargetRowNumber(rowNumber),
              values,
              visualRowNumber: rowNumber,
            };
          }
        }
      } else if (oursRowNumber && !theirsRowNumber) {
        if (source === 'theirs') {
          op = {
            sheetName: mergeInfo.sheetName,
            action: 'delete',
            targetRowNumber: oursRowNumber,
            visualRowNumber: rowNumber,
          };
        }
      }
      if (op || currentRowOps.has(rowNumber)) {
        updateRowOpForSheet(selectedMergeSheetIndex, rowNumber, op);
      }
    },
    [
      selectedMergeSheetIndex,
      mergeCells,
      markResolvedKeys,
      mergeRowsMeta,
      mergeInfo,
      buildMergedRowValues,
      computeInsertTargetRowNumber,
      updateRowOpForSheet,
      currentRowOps,
    ],
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

  const buildMergedColumnValues = useCallback(
    async (colNumber: number) => {
      if (!mergeInfo) return null;
      // Get all rows for this sheet to build column values
      const metas = [...mergeRowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
      if (metas.length === 0) return [];
      
      const rowsReq = metas.map((m) => ({
        rowNumber: m.visualRowNumber,
        baseRowNumber: m.baseRowNumber,
        oursRowNumber: m.oursRowNumber,
        theirsRowNumber: m.theirsRowNumber,
      }));
      const result = await window.excelAPI.getThreeWayRows({
        basePath: mergeInfo.basePath,
        oursPath: mergeInfo.oursPath,
        theirsPath: mergeInfo.theirsPath,
        sheetName: mergeInfo.sheetName,
        sheetIndex: selectedMergeSheetIndex,
        frozenRowCount: mergeFrozenRowCount,
        rows: rowsReq,
      });
      if (!result || !result.rows) return [];
      
      // Extract column values from result
      // IMPORTANT: 不要过滤任何行，必须收集所有行的值
      // 因为保存时列操作在行操作之前，那些行还没被删除
      const columnValues: (string | number | null)[] = [];
      result.rows.forEach((rowRes: any) => {
        const visualRowNumber = rowRes.rowNumber ?? 0;
        
        // Get value from aligned column (colNumber is 1-based aligned col)
        const diffCell = (mergeCellsByRow.get(visualRowNumber) ?? []).find((c) => c.col === colNumber);
        if (diffCell) {
          columnValues.push(diffCell.mergedValue ?? null);
        } else if (rowRes.theirs && colNumber >= 1 && colNumber <= rowRes.theirs.length) {
          columnValues.push(rowRes.theirs[colNumber - 1] ?? null);
        } else {
          columnValues.push(null);
        }
      });
      return columnValues;
    },
    [mergeInfo, selectedMergeSheetIndex, mergeRowsMeta, mergeFrozenRowCount, currentRowOps, mergeCellsByRow],
  );

  const handleApplyMergeColumnChoice = useCallback(
    async (colNumber: number, source: 'ours' | 'theirs') => {
      const valueFrom = (cell: MergeCell) => (source === 'theirs' ? cell.theirsValue : cell.oursValue);
      const keys = mergeCells.filter((c) => c.col === colNumber).map((c) => `${c.row}:${c.col}`);
      markResolvedKeys(selectedMergeSheetIndex, keys);

      setMergeSheets((prev) =>
        prev.map((sheet: MergeSheetData, sIdx: number) => {
          if (sIdx !== selectedMergeSheetIndex) return sheet;
          const newCells = sheet.cells.map((cell) => {
            if (cell.col !== colNumber) return cell;
            return { ...cell, mergedValue: valueFrom(cell) };
          });
          return { ...sheet, cells: newCells };
        }),
      );

      setMergeCells((prev) =>
        prev.map((cell) => {
          if (cell.col !== colNumber) return cell;
          return { ...cell, mergedValue: valueFrom(cell) };
        }),
      );

      if (!mergeInfo) return;
      const meta = mergeColumnsMeta.find((c) => c.col === colNumber);
      // theirs-only column -> insert
      const canInsert = source === 'theirs' && meta && !meta.oursCol && meta.theirsCol;
      // ours-only column but user chooses theirs (which is empty) -> delete
      const canDelete = source === 'theirs' && meta && meta.oursCol && !meta.theirsCol;
      if (canInsert) {
        const targetColNumber = computeInsertTargetColNumber(colNumber);
        const values = await buildMergedColumnValues(colNumber);
        if (values) {
          const op: SaveMergeColOp = {
            sheetName: mergeInfo.sheetName,
            action: 'insert',
            targetColNumber,
            alignedColNumber: colNumber,
            source,
            values,
          };
          updateColOpForSheet(selectedMergeSheetIndex, colNumber, op);
        }
      } else if (canDelete && meta.oursCol) {
        const op: SaveMergeColOp = {
          sheetName: mergeInfo.sheetName,
          action: 'delete',
          targetColNumber: meta.oursCol,
          alignedColNumber: colNumber,
          source,
        };
        updateColOpForSheet(selectedMergeSheetIndex, colNumber, op);
      } else if (currentColOps.has(colNumber)) {
        // Clear any existing op if user changes mind
        updateColOpForSheet(selectedMergeSheetIndex, colNumber, null);
      }
    },
    [
      mergeCells,
      markResolvedKeys,
      selectedMergeSheetIndex,
      mergeInfo,
      mergeColumnsMeta,
      computeInsertTargetColNumber,
      buildMergedColumnValues,
      updateColOpForSheet,
      currentColOps,
    ],
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
        const targetColNumber = cell.oursCol ?? null;
        if (hasRowMeta && !targetRowNumber) {
          skippedCells += 1;
          return;
        }
        if (!targetColNumber) {
          skippedCells += 1;
          return;
        }
        const address = targetRowNumber ? makeAddress(targetColNumber, targetRowNumber) : makeAddress(targetColNumber, cell.row);
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
      lines.push(`（提示：有 ${skippedCells} 个单元格因 ours 侧缺少对应行/列而未写入）`);
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
          const targetColNumber = cell.oursCol ?? null;
          if (!targetColNumber) return null;
          const address = targetRowNumber ? makeAddress(targetColNumber, targetRowNumber) : makeAddress(targetColNumber, cell.row);
          return {
            sheetName: sheet.sheetName,
            address,
            value: cell.mergedValue,
          };
        })
        .filter(Boolean) as { sheetName: string; address: string; value: string | number | null }[];
    });
    // 构建 aligned → physical 列映射：考虑列删除和列插入后，物理工作表的列布局
    const buildPhysicalColMap = (
      colsMeta: MergeColumnMeta[],
      colOpsMap: Map<number, SaveMergeColOp>,
    ): number[] => {
      const rawColCount = colsMeta.reduce((m, c) => Math.max(m, c.col), 0);
      const deletedCols = new Set<number>();
      const insertedCols: number[] = [];
      colOpsMap.forEach((op, ac) => {
        if (op.action === 'delete') deletedCols.add(ac);
        else if (op.action === 'insert') insertedCols.push(ac);
      });
      const map: number[] = [];
      for (let c = 1; c <= rawColCount; c += 1) {
        if (deletedCols.has(c)) continue;
        const m = colsMeta.find((cm) => cm.col === c);
        if (!m?.oursCol) continue;
        map.push(c);
      }
      insertedCols.sort((a, b) => a - b);
      const ins: Array<{ idx: number; col: number }> = [];
      for (const ac of insertedCols) {
        const m = colsMeta.find((cm) => cm.col === ac);
        if (m && !m.oursCol && m.theirsCol) {
          let insertIdx = map.length;
          for (let k = 0; k < map.length; k += 1) {
            if (map[k] > ac) { insertIdx = k; break; }
          }
          ins.push({ idx: insertIdx, col: ac });
        }
      }
      ins.sort((a, b) => b.idx - a.idx || b.col - a.col);
      for (const entry of ins) map.splice(entry.idx, 0, entry.col);
      return map;
    };

    const rowOps = Array.from(mergeRowOpsBySheet.entries()).flatMap(([sheetIndex, opsMap]) => {
      const sheet = mergeSheets[sheetIndex];
      const sheetName = sheet?.sheetName ?? mergeInfo.sheetName;
      const colsMeta = sheet?.columnsMeta ?? [];
      const colOpsForSheet = mergeColOpsBySheet.get(sheetIndex) ?? new Map<number, SaveMergeColOp>();
      // 将 row op 的 values 从 aligned 列空间重映射到物理列空间，
      // 跳过 theirs-only 列（除非用户选择了插入）和已删除的列。
      const colMap = colsMeta.length > 0 ? buildPhysicalColMap(colsMeta, colOpsForSheet) : null;
      return Array.from(opsMap.values()).map((op) => ({
        ...op,
        sheetName: op.sheetName || sheetName,
        values: op.values && colMap
          ? colMap.map((ac) => op.values![ac - 1] ?? null)
          : op.values,
      }));
    });
    const buildMergedColumnValues = (sheet: MergeSheetData, alignedColNumber: number) => {
      const rowsMeta = sheet.rowsMeta ?? [];
      const rowMetaMap = new Map<number, MergeRowMeta>();
      rowsMeta.forEach((m) => rowMetaMap.set(m.visualRowNumber, m));
      const maxRow = rowsMeta.reduce((m, r) => Math.max(m, r.oursRowNumber ?? 0), 0);
      const values: (string | number | null)[] = Array(maxRow).fill(null);
      sheet.cells.forEach((cell) => {
        if (cell.col !== alignedColNumber) return;
        const meta = rowMetaMap.get(cell.row);
        if (!meta?.oursRowNumber) return;
        values[meta.oursRowNumber - 1] = cell.mergedValue ?? null;
      });
      return values;
    };
    const colOps = Array.from(mergeColOpsBySheet.entries()).flatMap(([sheetIndex, opsMap]) => {
      const sheet = mergeSheets[sheetIndex];
      const sheetName = sheet?.sheetName ?? mergeInfo.sheetName;
      return Array.from(opsMap.values()).map((op) => ({
        ...op,
        sheetName: op.sheetName || sheetName,
        values: sheet && op.alignedColNumber ? buildMergedColumnValues(sheet, op.alignedColNumber) : op.values,
      }));
    });

    const payload: SaveMergeRequest = {
      templatePath: mergeInfo.oursPath,
      cells,
      rowOps,
      colOps,
      basePath: mergeInfo.basePath,
      oursPath: mergeInfo.oursPath,
      theirsPath: mergeInfo.theirsPath,
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
  }, [mergeInfo, mergeSheets, mergeRowOpsBySheet, mergeColOpsBySheet]);
  const mergedPreviewScrollToCell = useMemo(() => {
    if (!selectedMergeCell) return null;
    const visualRowNumber = selectedMergeCell.rowIndex + 1;
    const rowIndex = mergedPreviewRowVisuals.indexOf(visualRowNumber);
    if (rowIndex < 0) return null;
    return { rowIndex, colIndex: selectedMergeCell.colIndex };
  }, [selectedMergeCell, mergedPreviewRowVisuals]);
  const renderMergedPreviewRowHeader = (rowIndex: number) => {
    const visual = mergedPreviewRowVisuals[rowIndex];
    return visual == null ? '' : visual;
  };
  const renderMergedPreviewHeaderCell = (colIndex: number) => colNumberToLabel(colIndex + 1);
  const renderMergedPreviewCell = (cell: string | number | null, ctx: any) => {
    const visualRowNumber = mergedPreviewRowVisuals[ctx.rowIndex];
    const value = cell == null ? '' : String(cell);
    return (
      <div
        onMouseDown={() => {
          if (visualRowNumber == null) return;
          setSelectedMergeCell({ rowIndex: visualRowNumber - 1, colIndex: ctx.colIndex });
        }}
        title={value}
        style={{
          width: '100%',
          height: '100%',
          boxSizing: 'border-box',
          backgroundColor: 'transparent',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          cursor: 'pointer',
          userSelect: 'none',
        }}
      >
        {value}
      </div>
    );
  };
  const getMergedPreviewCellStyle = (_cell: any, ctx: any): React.CSSProperties => {
    const style: React.CSSProperties = {};
    if (ctx.isFrozenRow || ctx.isFrozenCol) {
      style.backgroundColor = '#f5f5f5';
    }
    if (selectedMergeCell) {
      const visualRowNumber = mergedPreviewRowVisuals[ctx.rowIndex];
      if (
        visualRowNumber === selectedMergeCell.rowIndex + 1 &&
        ctx.colIndex === selectedMergeCell.colIndex
      ) {
        style.outline = '2px solid #ff8000';
        style.outlineOffset = '-2px';
        style.position = 'relative';
        style.zIndex = 6;
      }
    }
    return style;
  };
  const mergedPreviewSafeRows = useMemo(() => {
    const minRows = Math.max(1, Math.floor(mergedPreviewMinRows));
    if (!mergedPreviewRows || mergedPreviewRows.length === 0) {
      return Array.from({ length: minRows }, () => [null]);
    }
    const first = mergedPreviewRows[0];
    const hasCols = Array.isArray(first) && first.length > 0;
    if (!hasCols) {
      return mergedPreviewRows.map(() => [null]);
    }
    return mergedPreviewRows;
  }, [mergedPreviewRows, mergedPreviewMinRows]);
  const mergedPreviewMinHeight = Math.max(mergedPreviewMinRows, 1) * 24 + 28;

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
                          setMergeColumnsMeta(sheet?.columnsMeta ?? []);
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
              <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 8 }}>
                <label style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontSize: 12 }}>
                  <input
                    type="checkbox"
                    checked={showFullTables}
                    onChange={(e) => setShowFullTables(e.target.checked)}
                  />
                  显示 ours/theirs 全表
                </label>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', marginTop: 4, gap: 4 }}>
                <span>merged 预览最少行数:</span>
                <input
                  type="number"
                  min={1}
                  value={mergedPreviewMinRows}
                  onChange={(e) => {
                    const v = Number(e.target.value);
                    if (Number.isNaN(v)) return;
                    setMergedPreviewMinRows(Math.max(1, Math.floor(v)));
                  }}
                  style={{ width: 60, padding: '2px 6px', boxSizing: 'border-box' }}
                />
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
                  onApplyColumnChoice={handleApplyMergeColumnChoice}
                  resolvedCellKeys={resolvedBySheet.get(selectedMergeSheetIndex)}
                  frozenRowCount={mergeFrozenRowCount}
                  primaryKeyCol={displayPrimaryKeyCol}
                  columnsMeta={mergeColumnsMeta}
                  oursPath={mergeInfo?.oursPath ?? null}
                  basePath={mergeInfo?.basePath ?? null}
                  theirsPath={mergeInfo?.theirsPath ?? null}
                  showFullTables={showFullTables}
                  fullOursRows={fullOursRows}
                  fullTheirsRows={fullTheirsRows}
                />
            </div>
            <div
              style={{
                marginTop: 8,
                border: '1px solid #ccc',
                overflow: 'hidden',
                flexShrink: 0,
                minHeight: mergedPreviewMinHeight + 24,
              }}
            >
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
              <div style={{ padding: 6 }}>
                <div style={{ height: mergedPreviewMinHeight }}>
                  <VirtualGrid<(string | number | null)>
                    rows={mergedPreviewSafeRows}
                    showRowHeader
                    renderRowHeader={renderMergedPreviewRowHeader}
                    renderHeaderCell={renderMergedPreviewHeaderCell}
                    renderCell={renderMergedPreviewCell}
                    getCellStyle={getMergedPreviewCellStyle}
                    frozenRowCount={0}
                    frozenColCount={0}
                    defaultColWidth={120}
                    scrollToCell={mergedPreviewScrollToCell}
                  />
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div>请先选择 base / ours / theirs 三个 Excel 文件。</div>
        )
      )}
      </div>
    </div>
  );
};

