import assert from 'node:assert/strict';
import * as fs from 'node:fs/promises';
import * as os from 'node:os';
import * as path from 'node:path';
import { app } from 'electron';
import { Workbook } from 'exceljs';
import { __testOnly } from '../main';

type SheetRow = Array<string | number | null>;
type SheetFixture = {
  name: string;
  rows: SheetRow[];
};
type StatusKey = 'unchanged' | 'added' | 'deleted' | 'modified' | 'ambiguous';
type MergeStatusKey = 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
type SeededKeyedCase = {
  sheetName: string;
  headers: string[];
  baseRows: SheetRow[];
  oursRows: SheetRow[];
  theirsRows: SheetRow[];
  theirsById: Map<string, SheetRow>;
  expected: {
    alignedRowCount: number;
    oursStats: Partial<Record<StatusKey, number>>;
    theirsStats: Partial<Record<StatusKey, number>>;
    cellStats: Partial<Record<MergeStatusKey, number>>;
  };
};

const HEADER_ROW_COUNT = 3;
const LARGE_ROW_SIMILARITY_THRESHOLD = 0.9;

const pad = (value: number, width = 4) => String(value).padStart(width, '0');

const makeHeaderRows = (columns: string[]): SheetRow[] => [
  columns.map((col) => `catalog:${col}`),
  columns.map((col, index) => `group-${index + 1}`),
  columns,
];

const cloneRows = (rows: SheetRow[]): SheetRow[] => rows.map((row) => [...row]);

const countBy = <T extends string>(values: T[]) =>
  values.reduce<Record<T, number>>((acc, value) => {
    acc[value] = (acc[value] ?? 0) + 1;
    return acc;
  }, {} as Record<T, number>);

const getCount = <T extends string>(stats: Partial<Record<T, number>>, key: T) => stats[key] ?? 0;

const createPrng = (seed: number) => {
  let state = seed >>> 0;
  return () => {
    state += 0x6d2b79f5;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
};

const randomInt = (rng: () => number, min: number, max: number) => Math.floor(rng() * (max - min + 1)) + min;

const shuffle = <T>(items: T[], rng: () => number): T[] => {
  const arr = items.slice();
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(rng() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
};

const insertEvery = (rows: SheetRow[], inserts: SheetRow[], every: number): SheetRow[] => {
  const result: SheetRow[] = [];
  let insertIndex = 0;
  rows.forEach((row, index) => {
    result.push(row);
    if ((index + 1) % every === 0 && insertIndex < inserts.length) {
      result.push(inserts[insertIndex]);
      insertIndex += 1;
    }
  });
  while (insertIndex < inserts.length) {
    result.push(inserts[insertIndex]);
    insertIndex += 1;
  }
  return result;
};

const insertAtRandomPositions = (rows: SheetRow[], inserts: SheetRow[], rng: () => number): SheetRow[] => {
  const result = rows.map((row) => [...row] as SheetRow);
  inserts.forEach((row) => {
    const pos = randomInt(rng, 0, result.length);
    result.splice(pos, 0, [...row]);
  });
  return result;
};

const colNumberToLabel = (col: number): string => {
  let n = col;
  let label = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
};

const makeAddress = (col: number, row: number) => `${colNumberToLabel(col)}${row}`;

const toSimpleCellValue = (value: unknown): string | number | null => {
  if (value == null) return null;
  if (typeof value === 'string' || typeof value === 'number') return value;
  if (typeof value === 'object') {
    const maybeFormula = value as { result?: unknown; text?: unknown; hyperlink?: unknown; richText?: Array<{ text?: string }> };
    if (typeof maybeFormula.result === 'string' || typeof maybeFormula.result === 'number') return maybeFormula.result;
    if (typeof maybeFormula.text === 'string') return maybeFormula.text;
    if (typeof maybeFormula.hyperlink === 'string') return maybeFormula.hyperlink;
    if (Array.isArray(maybeFormula.richText)) {
      return maybeFormula.richText.map((chunk) => chunk.text ?? '').join('');
    }
  }
  return String(value);
};

const readWorkbookRows = async (filePath: string, sheetName: string): Promise<SheetRow[]> => {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(filePath);
  const ws = workbook.getWorksheet(sheetName);
  assert.ok(ws, `Expected output workbook to contain sheet ${sheetName}`);
  const maxRow = (ws.actualRowCount ?? 0) > 0 ? ws.actualRowCount : ws.rowCount;
  const maxCol = (ws.actualColumnCount ?? 0) > 0 ? ws.actualColumnCount : ws.columnCount;
  const rows: SheetRow[] = [];
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const row: SheetRow = [];
    const wsRow = ws.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
      row.push(toSimpleCellValue(wsRow.getCell(colNumber).value));
    }
    rows.push(row);
  }
  return rows;
};

const buildSeededKeyedCase = (seed: number, rowCount: number, sheetName: string): SeededKeyedCase => {
  const rng = createPrng(seed);
  const headers = ['id', 'code', 'qty', 'note', 'owner'];
  const baseRows: SheetRow[] = Array.from({ length: rowCount }, (_, index) => {
    const id = index + 1;
    return [
      `RK-${seed}-${pad(id, 4)}`,
      `code-${(id * 17 + seed) % 97}`,
      1000 + id * 7,
      `note-${seed}-${pad(id, 4)}`,
      `owner-${(id * 11 + seed) % 23}`,
    ];
  });
  const pool = shuffle(
    Array.from({ length: rowCount }, (_, index) => index + 1),
    rng,
  );
  const takeIds = (count: number) => pool.splice(0, Math.max(0, Math.min(count, pool.length)));

  const oursOnlyIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(8, Math.floor(rowCount * 0.08)), Math.max(10, Math.floor(rowCount * 0.11)))),
  );
  const theirsOnlyIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(8, Math.floor(rowCount * 0.07)), Math.max(10, Math.floor(rowCount * 0.1)))),
  );
  const bothSameIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(6, Math.floor(rowCount * 0.06)), Math.max(8, Math.floor(rowCount * 0.09)))),
  );
  const conflictIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(6, Math.floor(rowCount * 0.05)), Math.max(8, Math.floor(rowCount * 0.08)))),
  );
  const oursDeleteIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(5, Math.floor(rowCount * 0.04)), Math.max(6, Math.floor(rowCount * 0.07)))),
  );
  const theirsDeleteIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(5, Math.floor(rowCount * 0.04)), Math.max(6, Math.floor(rowCount * 0.06)))),
  );
  const bothDeleteIds = takeIds(
    Math.min(pool.length, randomInt(rng, Math.max(4, Math.floor(rowCount * 0.03)), Math.max(5, Math.floor(rowCount * 0.05)))),
  );

  const oursById = new Map(baseRows.map((row) => [String(row[0]), [...row] as SheetRow]));
  const theirsById = new Map(baseRows.map((row) => [String(row[0]), [...row] as SheetRow]));

  oursOnlyIds.forEach((id) => {
    const row = oursById.get(`RK-${seed}-${pad(id, 4)}`);
    if (row) row[3] = `ours-note-${seed}-${pad(id, 4)}`;
  });
  theirsOnlyIds.forEach((id) => {
    const row = theirsById.get(`RK-${seed}-${pad(id, 4)}`);
    if (row) row[3] = `theirs-note-${seed}-${pad(id, 4)}`;
  });
  bothSameIds.forEach((id) => {
    const key = `RK-${seed}-${pad(id, 4)}`;
    const oursRow = oursById.get(key);
    const theirsRow = theirsById.get(key);
    const value = `shared-note-${seed}-${pad(id, 4)}`;
    if (oursRow) oursRow[3] = value;
    if (theirsRow) theirsRow[3] = value;
  });
  conflictIds.forEach((id) => {
    const key = `RK-${seed}-${pad(id, 4)}`;
    const oursRow = oursById.get(key);
    const theirsRow = theirsById.get(key);
    if (oursRow) oursRow[3] = `ours-conflict-${seed}-${pad(id, 4)}`;
    if (theirsRow) theirsRow[3] = `theirs-conflict-${seed}-${pad(id, 4)}`;
  });

  [...oursDeleteIds, ...bothDeleteIds].forEach((id) => {
    oursById.delete(`RK-${seed}-${pad(id, 4)}`);
  });
  [...theirsDeleteIds, ...bothDeleteIds].forEach((id) => {
    theirsById.delete(`RK-${seed}-${pad(id, 4)}`);
  });

  const oursInsertCount = randomInt(rng, Math.max(10, Math.floor(rowCount * 0.04)), Math.max(12, Math.floor(rowCount * 0.07)));
  const theirsInsertCount = randomInt(rng, Math.max(10, Math.floor(rowCount * 0.04)), Math.max(12, Math.floor(rowCount * 0.07)));
  const oursInserts: SheetRow[] = Array.from({ length: oursInsertCount }, (_, index) => {
    const seq = index + 1;
    return [
      `RKO-${seed}-${pad(seq, 3)}`,
      `ours-code-${(seq * 13 + seed) % 89}`,
      500000 + seq,
      `ours-insert-note-${seed}-${pad(seq, 3)}`,
      `ours-insert-owner-${seq % 17}`,
    ];
  });
  const theirsInserts: SheetRow[] = Array.from({ length: theirsInsertCount }, (_, index) => {
    const seq = index + 1;
    return [
      `RKT-${seed}-${pad(seq, 3)}`,
      `theirs-code-${(seq * 19 + seed) % 83}`,
      700000 + seq,
      `theirs-insert-note-${seed}-${pad(seq, 3)}`,
      `theirs-insert-owner-${seq % 19}`,
    ];
  });

  const oursRows = insertAtRandomPositions(Array.from(oursById.values()), oursInserts, rng);
  const theirsRows = insertAtRandomPositions(Array.from(theirsById.values()), theirsInserts, rng);
  theirsInserts.forEach((row) => {
    theirsById.set(String(row[0]), [...row]);
  });
  const remainingBase =
    rowCount -
    oursOnlyIds.length -
    theirsOnlyIds.length -
    bothSameIds.length -
    conflictIds.length -
    oursDeleteIds.length -
    theirsDeleteIds.length -
    bothDeleteIds.length;

  return {
    sheetName,
    headers,
    baseRows,
    oursRows,
    theirsRows,
    theirsById,
    expected: {
      alignedRowCount: rowCount + oursInsertCount + theirsInsertCount,
      oursStats: {
        added: oursInsertCount,
        deleted: oursDeleteIds.length + bothDeleteIds.length,
        modified: oursOnlyIds.length + bothSameIds.length + conflictIds.length,
        unchanged: remainingBase + theirsOnlyIds.length + theirsDeleteIds.length + theirsInsertCount,
        ambiguous: 0,
      },
      theirsStats: {
        added: theirsInsertCount,
        deleted: theirsDeleteIds.length + bothDeleteIds.length,
        modified: theirsOnlyIds.length + bothSameIds.length + conflictIds.length,
        unchanged: remainingBase + oursOnlyIds.length + oursDeleteIds.length + oursInsertCount,
        ambiguous: 0,
      },
      cellStats: {
        'ours-changed': oursOnlyIds.length,
        'theirs-changed': theirsOnlyIds.length,
        'both-changed-same': bothSameIds.length,
        conflict: conflictIds.length,
      },
    },
  };
};

const computeInsertTargetRowNumber = (
  rowsMeta: Array<{ visualRowNumber: number; oursRowNumber: number | null }>,
  visualRowNumber: number,
) => {
  const list = [...rowsMeta].sort((a, b) => a.visualRowNumber - b.visualRowNumber);
  const idx = list.findIndex((item) => item.visualRowNumber === visualRowNumber);
  if (idx < 0) return 1;
  for (let i = idx - 1; i >= 0; i -= 1) {
    const rowNumber = list[i].oursRowNumber;
    if (rowNumber) return rowNumber + 1;
  }
  for (let i = idx + 1; i < list.length; i += 1) {
    const rowNumber = list[i].oursRowNumber;
    if (rowNumber) return rowNumber;
  }
  return 1;
};

const writeWorkbook = async (filePath: string, sheets: SheetFixture[]) => {
  const workbook = new Workbook();
  sheets.forEach((sheet) => {
    const ws = workbook.addWorksheet(sheet.name);
    sheet.rows.forEach((row) => ws.addRow(row));
  });
  await workbook.xlsx.writeFile(filePath);
};

const writeWorkbookTriplet = async (
  rootDir: string,
  prefix: string,
  baseSheets: SheetFixture[],
  oursSheets: SheetFixture[],
  theirsSheets: SheetFixture[],
) => {
  const basePath = path.join(rootDir, `${prefix}.base.xlsx`);
  const oursPath = path.join(rootDir, `${prefix}.ours.xlsx`);
  const theirsPath = path.join(rootDir, `${prefix}.theirs.xlsx`);
  await writeWorkbook(basePath, baseSheets);
  await writeWorkbook(oursPath, oursSheets);
  await writeWorkbook(theirsPath, theirsSheets);
  return { basePath, oursPath, theirsPath };
};

const getSheetOrThrow = <T extends { sheetName: string }>(mergeSheets: T[], sheetName: string): T => {
  const sheet = mergeSheets.find((item) => item.sheetName === sheetName);
  assert.ok(sheet, `Expected sheet ${sheetName} to exist`);
  return sheet;
};

const runCellStatusLargeScenario = async (rootDir: string) => {
  const sheetName = 'CellStatusLarge';
  const headers = makeHeaderRows(['id', 'value', 'owner', 'note']);
  const baseBody: SheetRow[] = Array.from({ length: 1200 }, (_, index) => {
    const id = index + 1;
    return [`ID-${pad(id)}`, `value-${pad(id)}`, `owner-${id % 29}`, `note-${pad(id)}`];
  });
  const oursBody = cloneRows(baseBody);
  const theirsBody = cloneRows(baseBody);

  const oursOnlyIds = Array.from({ length: 220 }, (_, index) => index + 1);
  const theirsOnlyIds = Array.from({ length: 210 }, (_, index) => index + 221);
  const bothSameIds = Array.from({ length: 180 }, (_, index) => index + 431);
  const conflictIds = Array.from({ length: 150 }, (_, index) => index + 611);

  oursOnlyIds.forEach((id) => {
    oursBody[id - 1][3] = `ours-only-${pad(id)}`;
  });
  theirsOnlyIds.forEach((id) => {
    theirsBody[id - 1][3] = `theirs-only-${pad(id)}`;
  });
  bothSameIds.forEach((id) => {
    const value = `both-same-${pad(id)}`;
    oursBody[id - 1][3] = value;
    theirsBody[id - 1][3] = value;
  });
  conflictIds.forEach((id) => {
    oursBody[id - 1][3] = `ours-conflict-${pad(id)}`;
    theirsBody[id - 1][3] = `theirs-conflict-${pad(id)}`;
  });

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'cell-status-large',
    [{ name: sheetName, rows: [...headers, ...baseBody] }],
    [{ name: sheetName, rows: [...headers, ...oursBody] }],
    [{ name: sheetName, rows: [...headers, ...theirsBody] }],
  );

  const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );
  const sheet = getSheetOrThrow(mergeSheets, sheetName);
  const bodyNoteCells = sheet.cells.filter((cell) => cell.row > HEADER_ROW_COUNT && cell.col === 4);
  const cellStats = countBy(bodyNoteCells.map((cell) => cell.status as 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict'));

  assert.equal(bodyNoteCells.length, 760, 'CellStatusLarge should expose all generated body diffs');
  assert.equal(getCount(cellStats, 'ours-changed'), 220);
  assert.equal(getCount(cellStats, 'theirs-changed'), 210);
  assert.equal(getCount(cellStats, 'both-changed-same'), 180);
  assert.equal(getCount(cellStats, 'conflict'), 150);
  assert.equal(sheet.hasExactDiff, true);

  const bodyRowsMeta = (sheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);
  const oursStats = countBy(bodyRowsMeta.map((row) => row.oursStatus));
  const theirsStats = countBy(bodyRowsMeta.map((row) => row.theirsStatus));
  assert.equal(getCount(oursStats, 'modified'), 550);
  assert.equal(getCount(theirsStats, 'modified'), 540);
};

const runModeDifferenceLargeScenario = async (rootDir: string) => {
  const sheetName = 'ModeDifferenceLarge';
  const headers = makeHeaderRows(['id', 'value', 'status', 'note']);
  const baseBody: SheetRow[] = Array.from({ length: 900 }, (_, index) => {
    const id = index + 1;
    return [`MD-${pad(id)}`, `base-${pad(id)}`, `status-${id % 11}`, `note-${pad(id)}`];
  });
  const oursBody = cloneRows(baseBody);
  const theirsBody = cloneRows(baseBody);

  Array.from({ length: 540 }, (_, index) => index + 1).forEach((id) => {
    const value = `shared-change-${pad(id)}`;
    oursBody[id - 1][1] = value;
    theirsBody[id - 1][1] = value;
  });

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'mode-difference-large',
    [{ name: sheetName, rows: [...headers, ...baseBody] }],
    [{ name: sheetName, rows: [...headers, ...oursBody] }],
    [{ name: sheetName, rows: [...headers, ...theirsBody] }],
  );

  const mergeResult = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );
  const diffResult = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'diff',
  );

  const mergeSheet = getSheetOrThrow(mergeResult.mergeSheets, sheetName);
  const diffSheet = getSheetOrThrow(diffResult.mergeSheets, sheetName);

  const mergeCells = mergeSheet.cells.filter((cell) => cell.row > HEADER_ROW_COUNT && cell.col === 2);
  const mergeStats = countBy(mergeCells.map((cell) => cell.status as 'both-changed-same'));
  assert.equal(mergeCells.length, 540);
  assert.equal(getCount(mergeStats, 'both-changed-same'), 540);

  const mergeBodyRows = (mergeSheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);
  const mergeOursStats = countBy(mergeBodyRows.map((row) => row.oursStatus));
  const mergeTheirsStats = countBy(mergeBodyRows.map((row) => row.theirsStatus));
  assert.equal(getCount(mergeOursStats, 'modified'), 540);
  assert.equal(getCount(mergeTheirsStats, 'modified'), 540);

  assert.equal(diffSheet.cells.length, 0, 'diff mode should short-circuit when ours and theirs are coordinate-equal');
  assert.equal(diffSheet.hasExactDiff, false);
  assert.equal((diffSheet.rowsMeta ?? []).length, 0);
};

const runStructuredHeaderDetectionScenario = async (rootDir: string) => {
  const sheetName = 'StructuredHeaderRows';
  const headerRows: SheetRow[] = [
    ['##comment', 'id', '#成就组备注', '成就'],
    ['##var', 'id', '##mark', 'achievement'],
    ['##type', 'int', 'string', '(list#sep=|),int'],
    ['##group', null, null, null],
  ];
  const baseBody: SheetRow[] = [
    [null, 1, null, '10001|10002|10003'],
    [null, 2, null, '11001|11002|11003|41001'],
  ];
  const oursBody: SheetRow[] = [
    [null, 1, null, '10001|1000aasdasd2|10003'],
    [null, 2, null, '11001|11002|11003|41001'],
  ];
  const theirsBody: SheetRow[] = [
    [null, 1, null, '10001|100xx|10003'],
    [null, 2, null, '11001|11002|11003|41001'],
  ];

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'structured-header-rows',
    [{ name: sheetName, rows: [...headerRows, ...baseBody] }],
    [{ name: sheetName, rows: [...headerRows, ...oursBody] }],
    [{ name: sheetName, rows: [...headerRows, ...theirsBody] }],
  );

  const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    -1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );
  const sheet = getSheetOrThrow(mergeSheets, sheetName);
  const rowsMeta = sheet.rowsMeta ?? [];
  assert.ok(rowsMeta.length >= 6, 'Structured header scenario should preserve all header and body rows');
  assert.deepStrictEqual(
    rowsMeta.slice(0, 6).map((row) => ({
      visualRowNumber: row.visualRowNumber,
      baseRowNumber: row.baseRowNumber,
      oursRowNumber: row.oursRowNumber,
      theirsRowNumber: row.theirsRowNumber,
    })),
    [
      { visualRowNumber: 1, baseRowNumber: 1, oursRowNumber: 1, theirsRowNumber: 1 },
      { visualRowNumber: 2, baseRowNumber: 2, oursRowNumber: 2, theirsRowNumber: 2 },
      { visualRowNumber: 3, baseRowNumber: 3, oursRowNumber: 3, theirsRowNumber: 3 },
      { visualRowNumber: 4, baseRowNumber: 4, oursRowNumber: 4, theirsRowNumber: 4 },
      { visualRowNumber: 5, baseRowNumber: 5, oursRowNumber: 5, theirsRowNumber: 5 },
      { visualRowNumber: 6, baseRowNumber: 6, oursRowNumber: 6, theirsRowNumber: 6 },
    ],
    'Structured Luban headers should not shift body row mapping',
  );
  const diffCells = (sheet.cells ?? []).map((cell) => ({
    row: cell.row,
    col: cell.col,
    status: cell.status,
    baseValue: cell.baseValue,
    oursValue: cell.oursValue,
    theirsValue: cell.theirsValue,
  }));
  assert.deepStrictEqual(diffCells, [
    {
      row: 1,
      col: 4,
      status: 'unchanged',
      baseValue: '成就',
      oursValue: '成就',
      theirsValue: '成就',
    },
    {
      row: 2,
      col: 4,
      status: 'unchanged',
      baseValue: 'achievement',
      oursValue: 'achievement',
      theirsValue: 'achievement',
    },
    {
      row: 3,
      col: 4,
      status: 'unchanged',
      baseValue: '(list#sep=|),int',
      oursValue: '(list#sep=|),int',
      theirsValue: '(list#sep=|),int',
    },
    {
      row: 4,
      col: 4,
      status: 'unchanged',
      baseValue: null,
      oursValue: null,
      theirsValue: null,
    },
    {
      row: 5,
      col: 4,
      status: 'conflict',
      baseValue: '10001|10002|10003',
      oursValue: '10001|1000aasdasd2|10003',
      theirsValue: '10001|100xx|10003',
    },
  ]);
};

const runStructuralRowsLargeScenario = async (rootDir: string) => {
  const sheetName = 'StructuralRowsLarge';
  const headers = makeHeaderRows(['id', 'value', 'qty', 'owner']);
  const baseBody: SheetRow[] = Array.from({ length: 900 }, (_, index) => {
    const id = index + 1;
    return [`ST-${pad(id)}`, `value-${pad(id)}`, id * 10, `owner-${id % 17}`];
  });

  const oursById = new Map(baseBody.map((row) => [String(row[0]), [...row] as SheetRow]));
  const theirsById = new Map(baseBody.map((row) => [String(row[0]), [...row] as SheetRow]));

  const oursDeleteIds = Array.from({ length: 60 }, (_, index) => `ST-${pad(index + 1)}`);
  const theirsDeleteIds = Array.from({ length: 50 }, (_, index) => `ST-${pad(index + 61)}`);
  const bothDeleteIds = Array.from({ length: 40 }, (_, index) => `ST-${pad(index + 111)}`);
  const oursModifyIds = Array.from({ length: 120 }, (_, index) => `ST-${pad(index + 151)}`);
  const theirsModifyIds = Array.from({ length: 90 }, (_, index) => `ST-${pad(index + 271)}`);
  const bothSameIds = Array.from({ length: 70 }, (_, index) => `ST-${pad(index + 361)}`);
  const conflictIds = Array.from({ length: 60 }, (_, index) => `ST-${pad(index + 431)}`);

  [...oursDeleteIds, ...bothDeleteIds].forEach((id) => oursById.delete(id));
  [...theirsDeleteIds, ...bothDeleteIds].forEach((id) => theirsById.delete(id));

  oursModifyIds.forEach((id, index) => {
    const row = oursById.get(id);
    assert.ok(row, `Expected ours row ${id} to exist`);
    row[2] = 100000 + index;
  });
  theirsModifyIds.forEach((id, index) => {
    const row = theirsById.get(id);
    assert.ok(row, `Expected theirs row ${id} to exist`);
    row[3] = `theirs-owner-${index}`;
  });
  bothSameIds.forEach((id) => {
    const oursRow = oursById.get(id);
    const theirsRow = theirsById.get(id);
    assert.ok(oursRow && theirsRow, `Expected shared row ${id} to exist`);
    const value = `shared-${id}`;
    oursRow[1] = value;
    theirsRow[1] = value;
  });
  conflictIds.forEach((id, index) => {
    const oursRow = oursById.get(id);
    const theirsRow = theirsById.get(id);
    assert.ok(oursRow && theirsRow, `Expected conflict row ${id} to exist`);
    oursRow[2] = 200000 + index;
    theirsRow[2] = 300000 + index;
  });

  const oursInserts: SheetRow[] = Array.from({ length: 45 }, (_, index) => [
    `ST-OURS-ADD-${pad(index + 1, 3)}`,
    `ours-added-${pad(index + 1, 3)}`,
    500000 + index,
    `ours-owner-${index % 9}`,
  ]);
  const theirsInserts: SheetRow[] = Array.from({ length: 55 }, (_, index) => [
    `ST-THEIRS-ADD-${pad(index + 1, 3)}`,
    `theirs-added-${pad(index + 1, 3)}`,
    600000 + index,
    `theirs-owner-${index % 11}`,
  ]);

  const oursBody = insertEvery(Array.from(oursById.values()), oursInserts, 17);
  const theirsBody = insertEvery(Array.from(theirsById.values()), theirsInserts, 13);

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'structural-rows-large',
    [{ name: sheetName, rows: [...headers, ...baseBody] }],
    [{ name: sheetName, rows: [...headers, ...oursBody] }],
    [{ name: sheetName, rows: [...headers, ...theirsBody] }],
  );

  const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );
  const sheet = getSheetOrThrow(mergeSheets, sheetName);
  const bodyRows = (sheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);
  assert.equal(bodyRows.length, 1000);

  const oursStats = countBy(bodyRows.map((row) => row.oursStatus));
  const theirsStats = countBy(bodyRows.map((row) => row.theirsStatus));
  assert.equal(getCount(oursStats, 'added'), 45);
  assert.equal(getCount(oursStats, 'deleted'), 100);
  assert.equal(getCount(oursStats, 'modified'), 250);
  assert.equal(getCount(oursStats, 'unchanged'), 605);
  assert.equal(getCount(oursStats, 'ambiguous'), 0);

  assert.equal(getCount(theirsStats, 'added'), 55);
  assert.equal(getCount(theirsStats, 'deleted'), 90);
  assert.equal(getCount(theirsStats, 'modified'), 220);
  assert.equal(getCount(theirsStats, 'unchanged'), 635);
  assert.equal(getCount(theirsStats, 'ambiguous'), 0);
};
const runStructuralColumnsLargeScenario = async (rootDir: string) => {
  const sheetName = 'StructuralColumnsLarge';
  const baseHeaders = makeHeaderRows(['id', 'name', 'shared_value', 'status', 'tail']);
  const oursHeaders = makeHeaderRows(['id', 'name', 'ours_only_metric', 'shared_value', 'status', 'tail']);
  const theirsHeaders = makeHeaderRows(['id', 'name', 'shared_value', 'status', 'tail', 'theirs_only_metric']);
  const baseBody: SheetRow[] = Array.from({ length: 750 }, (_, index) => {
    const id = index + 1;
    return [`SC-${pad(id)}`, `name-${pad(id)}`, `shared-${pad(id)}`, `status-${id % 13}`, `tail-${id % 9}`];
  });
  const oursBody = baseBody.map((row) => [row[0], row[1], `ours-only-${row[0]}`, row[2], row[3], row[4]]);
  const theirsBody = baseBody.map((row) => [row[0], row[1], row[2], row[3], row[4], `theirs-only-${row[0]}`]);

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'structural-columns-large',
    [{ name: sheetName, rows: [...baseHeaders, ...baseBody] }],
    [{ name: sheetName, rows: [...oursHeaders, ...oursBody] }],
    [{ name: sheetName, rows: [...theirsHeaders, ...theirsBody] }],
  );

  const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );
  const sheet = getSheetOrThrow(mergeSheets, sheetName);
  assert.ok(sheet.columnsMeta?.some((col) => col.oursCol === 3 && col.theirsCol == null));
  assert.ok(sheet.columnsMeta?.some((col) => col.oursCol == null && col.theirsCol === 6));

  const oursOnlyCells = sheet.cells.filter((cell) => cell.row > HEADER_ROW_COUNT && cell.oursCol === 3);
  const theirsOnlyCells = sheet.cells.filter((cell) => cell.row > HEADER_ROW_COUNT && cell.theirsCol === 6);
  assert.equal(oursOnlyCells.length, 750);
  assert.equal(theirsOnlyCells.length, 750);
  assert.ok(oursOnlyCells.every((cell) => cell.status === 'ours-changed'));
  assert.ok(theirsOnlyCells.every((cell) => cell.status === 'theirs-changed'));

  const bodyRows = (sheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);
  const oursStats = countBy(bodyRows.map((row) => row.oursStatus));
  const theirsStats = countBy(bodyRows.map((row) => row.theirsStatus));
  assert.equal(getCount(oursStats, 'modified'), 750);
  assert.equal(getCount(theirsStats, 'modified'), 750);
  assert.equal(getCount(oursStats, 'ambiguous'), 0);
  assert.equal(getCount(theirsStats, 'ambiguous'), 0);
};

const runNoKeyLargeScenario = async (rootDir: string) => {
  const sheetName = 'NoKeyLarge';
  const headers = makeHeaderRows(['group', 'bucket', 'payload', 'zone', 'marker']);
  const baseBody: SheetRow[] = Array.from({ length: 600 }, (_, index) => {
    const rowIndex = index + 1;
    return [
      `G-${rowIndex % 20}`,
      `B-${Math.floor(index / 20)}`,
      `P-${(rowIndex * 7) % 15}`,
      `Z-${(rowIndex * 11) % 12}`,
      `M-${(rowIndex * 13) % 10}`,
    ];
  });

  const oursBody = cloneRows(baseBody);
  const theirsBody = cloneRows(baseBody);

  const oursDeleteIndexes = new Set(Array.from({ length: 25 }, (_, index) => index));
  const theirsDeleteIndexes = new Set(Array.from({ length: 20 }, (_, index) => index + 25));
  const bothDeleteIndexes = new Set(Array.from({ length: 15 }, (_, index) => index + 45));
  const oursModifyIndexes = Array.from({ length: 40 }, (_, index) => index + 60);
  const theirsModifyIndexes = Array.from({ length: 35 }, (_, index) => index + 100);
  const bothSameIndexes = Array.from({ length: 30 }, (_, index) => index + 135);
  const conflictIndexes = Array.from({ length: 25 }, (_, index) => index + 165);

  oursModifyIndexes.forEach((index) => {
    oursBody[index][2] = `ours-payload-${pad(index + 1, 3)}`;
  });
  theirsModifyIndexes.forEach((index) => {
    theirsBody[index][3] = `theirs-zone-${pad(index + 1, 3)}`;
  });
  bothSameIndexes.forEach((index) => {
    const value = `both-marker-${pad(index + 1, 3)}`;
    oursBody[index][4] = value;
    theirsBody[index][4] = value;
  });
  conflictIndexes.forEach((index) => {
    oursBody[index][2] = `ours-conflict-${pad(index + 1, 3)}`;
    theirsBody[index][2] = `theirs-conflict-${pad(index + 1, 3)}`;
  });

  const filteredOurs = oursBody.filter((_, index) => !oursDeleteIndexes.has(index) && !bothDeleteIndexes.has(index));
  const filteredTheirs = theirsBody.filter((_, index) => !theirsDeleteIndexes.has(index) && !bothDeleteIndexes.has(index));

  const oursInserts: SheetRow[] = Array.from({ length: 20 }, (_, index) => [
    `G-INS-${index % 5}`,
    `B-INS-${Math.floor(index / 5)}`,
    `P-INS-${index % 7}`,
    `Z-INS-${index % 6}`,
    `M-INS-${index % 4}`,
  ]);
  const theirsInserts: SheetRow[] = Array.from({ length: 22 }, (_, index) => [
    `G-TINS-${index % 6}`,
    `B-TINS-${Math.floor(index / 6)}`,
    `P-TINS-${index % 8}`,
    `Z-TINS-${index % 5}`,
    `M-TINS-${index % 3}`,
  ]);

  const oursRows = insertEvery(filteredOurs, oursInserts, 19);
  const theirsRows = insertEvery(filteredTheirs, theirsInserts, 23);

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'no-key-large',
    [{ name: sheetName, rows: [...headers, ...baseBody] }],
    [{ name: sheetName, rows: [...headers, ...oursRows] }],
    [{ name: sheetName, rows: [...headers, ...theirsRows] }],
  );

  const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    -1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );
  const sheet = getSheetOrThrow(mergeSheets, sheetName);
  const bodyRows = (sheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);

  const oursStats = countBy(bodyRows.map((row) => row.oursStatus));
  const theirsStats = countBy(bodyRows.map((row) => row.theirsStatus));
  assert.equal(bodyRows.length, 642);
  assert.equal(getCount(oursStats, 'added'), 20);
  assert.equal(getCount(oursStats, 'deleted'), 40);
  assert.equal(getCount(oursStats, 'modified'), 95);
  assert.equal(getCount(oursStats, 'unchanged'), 487);
  assert.equal(getCount(oursStats, 'ambiguous'), 0);

  assert.equal(getCount(theirsStats, 'added'), 22);
  assert.equal(getCount(theirsStats, 'deleted'), 35);
  assert.equal(getCount(theirsStats, 'modified'), 90);
  assert.equal(getCount(theirsStats, 'unchanged'), 495);
  assert.equal(getCount(theirsStats, 'ambiguous'), 0);
};

const runSeededRandomBatchScenario = async (rootDir: string) => {
  const cases = [
    buildSeededKeyedCase(101, 260, 'SeededRandom101'),
    buildSeededKeyedCase(203, 320, 'SeededRandom203'),
    buildSeededKeyedCase(307, 380, 'SeededRandom307'),
    buildSeededKeyedCase(401, 440, 'SeededRandom401'),
    buildSeededKeyedCase(509, 300, 'SeededRandom509'),
    buildSeededKeyedCase(601, 360, 'SeededRandom601'),
  ];

  for (const seededCase of cases) {
    const headerRows = makeHeaderRows(seededCase.headers);
    const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
      rootDir,
      seededCase.sheetName,
      [{ name: seededCase.sheetName, rows: [...headerRows, ...seededCase.baseRows] }],
      [{ name: seededCase.sheetName, rows: [...headerRows, ...seededCase.oursRows] }],
      [{ name: seededCase.sheetName, rows: [...headerRows, ...seededCase.theirsRows] }],
    );
    const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
      basePath,
      oursPath,
      theirsPath,
      1,
      HEADER_ROW_COUNT,
      LARGE_ROW_SIMILARITY_THRESHOLD,
      'merge',
    );
    const sheet = getSheetOrThrow(mergeSheets, seededCase.sheetName);
    const bodyRows = (sheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);
    assert.equal(bodyRows.length, seededCase.expected.alignedRowCount, `${seededCase.sheetName} aligned row count mismatch`);

    const oursStats = countBy(bodyRows.map((row) => row.oursStatus));
    const theirsStats = countBy(bodyRows.map((row) => row.theirsStatus));
    (Object.keys(seededCase.expected.oursStats) as StatusKey[]).forEach((key) => {
      assert.equal(getCount(oursStats, key), getCount(seededCase.expected.oursStats, key), `${seededCase.sheetName} ours ${key} mismatch`);
    });
    (Object.keys(seededCase.expected.theirsStats) as StatusKey[]).forEach((key) => {
      assert.equal(getCount(theirsStats, key), getCount(seededCase.expected.theirsStats, key), `${seededCase.sheetName} theirs ${key} mismatch`);
    });

    const completeRows = new Set(
      bodyRows
        .filter((row) => row.baseRowNumber && row.oursRowNumber && row.theirsRowNumber)
        .map((row) => row.visualRowNumber),
    );
    const cellStats = countBy(
      sheet.cells
        .filter((cell) => cell.col === 4 && completeRows.has(cell.row))
        .map((cell) => cell.status as MergeStatusKey),
    );
    (Object.keys(seededCase.expected.cellStats) as MergeStatusKey[]).forEach((key) => {
      assert.equal(getCount(cellStats, key), getCount(seededCase.expected.cellStats, key), `${seededCase.sheetName} cell ${key} mismatch`);
    });
  }
};

const runMultiSheetLargeScenario = async (rootDir: string) => {
  const namedAlpha = 'NamedAlpha';
  const indexFallbackBase = 'IndexFallbackBase';
  const namedOmega = 'NamedOmega';
  const alphaHeaders = makeHeaderRows(['id', 'value', 'note']);
  const alphaBaseBody: SheetRow[] = Array.from({ length: 320 }, (_, index) => {
    const id = index + 1;
    return [`MSA-${pad(id)}`, `value-${pad(id)}`, `note-${pad(id)}`];
  });
  const alphaOursBody = cloneRows(alphaBaseBody);
  const alphaTheirsBody = cloneRows(alphaBaseBody);
  Array.from({ length: 80 }, (_, index) => index + 1).forEach((id) => {
    alphaOursBody[id - 1][2] = `ours-alpha-${pad(id)}`;
  });
  Array.from({ length: 70 }, (_, index) => index + 81).forEach((id) => {
    alphaTheirsBody[id - 1][2] = `theirs-alpha-${pad(id)}`;
  });

  const fallbackHeaders = makeHeaderRows(['id', 'qty', 'owner']);
  const fallbackBaseBody: SheetRow[] = Array.from({ length: 260 }, (_, index) => {
    const id = index + 1;
    return [`MSF-${pad(id)}`, id * 3, `owner-${id % 19}`];
  });
  const fallbackOursBody = cloneRows(fallbackBaseBody);
  const fallbackTheirsBody = cloneRows(fallbackBaseBody);
  Array.from({ length: 60 }, (_, index) => index + 1).forEach((id, index) => {
    fallbackOursBody[id - 1][1] = 50000 + index;
  });
  Array.from({ length: 55 }, (_, index) => index + 61).forEach((id, index) => {
    fallbackTheirsBody[id - 1][2] = `fallback-theirs-${index}`;
  });

  const omegaBaseHeaders = makeHeaderRows(['id', 'name', 'shared_value', 'status', 'tail']);
  const omegaOursHeaders = makeHeaderRows(['id', 'name', 'ours_only_metric', 'shared_value', 'status', 'tail']);
  const omegaTheirsHeaders = makeHeaderRows(['id', 'name', 'shared_value', 'status', 'tail', 'theirs_only_metric']);
  const omegaBaseBody: SheetRow[] = Array.from({ length: 240 }, (_, index) => {
    const id = index + 1;
    return [`MSO-${pad(id)}`, `name-${pad(id)}`, `shared-${pad(id)}`, `status-${id % 7}`, `tail-${id % 5}`];
  });
  const omegaOursBody = omegaBaseBody.map((row) => [row[0], row[1], `omega-ours-${row[0]}`, row[2], row[3], row[4]]);
  const omegaTheirsBody = omegaBaseBody.map((row) => [row[0], row[1], row[2], row[3], row[4], `omega-theirs-${row[0]}`]);

  const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
    rootDir,
    'multi-sheet-large',
    [
      { name: namedAlpha, rows: [...alphaHeaders, ...alphaBaseBody] },
      { name: indexFallbackBase, rows: [...fallbackHeaders, ...fallbackBaseBody] },
      { name: namedOmega, rows: [...omegaBaseHeaders, ...omegaBaseBody] },
    ],
    [
      { name: namedAlpha, rows: [...alphaHeaders, ...alphaOursBody] },
      { name: 'IndexFallbackOurs', rows: [...fallbackHeaders, ...fallbackOursBody] },
      { name: namedOmega, rows: [...omegaOursHeaders, ...omegaOursBody] },
    ],
    [
      { name: namedAlpha, rows: [...alphaHeaders, ...alphaTheirsBody] },
      { name: 'IndexFallbackTheirs', rows: [...fallbackHeaders, ...fallbackTheirsBody] },
      { name: namedOmega, rows: [...omegaTheirsHeaders, ...omegaTheirsBody] },
    ],
  );

  const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
    basePath,
    oursPath,
    theirsPath,
    1,
    HEADER_ROW_COUNT,
    LARGE_ROW_SIMILARITY_THRESHOLD,
    'merge',
  );

  assert.equal(mergeSheets.length, 3);
  assert.deepStrictEqual(
    mergeSheets.map((sheet) => sheet.sheetName),
    [namedAlpha, namedOmega, indexFallbackBase],
  );

  const alphaSheet = getSheetOrThrow(mergeSheets, namedAlpha);
  const alphaBodyNoteCells = alphaSheet.cells.filter((cell) => cell.row > HEADER_ROW_COUNT && cell.col === 3);
  assert.equal(alphaBodyNoteCells.length, 150);

  const omegaSheet = getSheetOrThrow(mergeSheets, namedOmega);
  assert.ok(omegaSheet.columnsMeta?.some((col) => col.oursCol === 3 && col.theirsCol == null));
  assert.ok(omegaSheet.columnsMeta?.some((col) => col.oursCol == null && col.theirsCol === 6));

  const fallbackSheet = getSheetOrThrow(mergeSheets, indexFallbackBase);
  const fallbackBodyRows = (fallbackSheet.rowsMeta ?? []).filter((row) => row.visualRowNumber > HEADER_ROW_COUNT);
  const fallbackOursStats = countBy(fallbackBodyRows.map((row) => row.oursStatus));
  const fallbackTheirsStats = countBy(fallbackBodyRows.map((row) => row.theirsStatus));
  assert.equal(getCount(fallbackOursStats, 'modified'), 60);
  assert.equal(getCount(fallbackTheirsStats, 'modified'), 55);
};

const runSaveMergeWriteBackScenario = async (rootDir: string) => {
  const mainSheet = 'WriteBackMain';
  const sideSheet = 'WriteBackSide';
  const mainHeaders = makeHeaderRows(['id', 'label', 'ours_only', 'amount', 'status']);
  const mainBody: SheetRow[] = Array.from({ length: 120 }, (_, index) => {
    const id = index + 1;
    return [`WB-${pad(id)}`, `label-${pad(id)}`, `ours-only-${pad(id)}`, id * 100, `status-${id % 9}`];
  });
  const sideHeaders = makeHeaderRows(['id', 'phase', 'score']);
  const sideBody: SheetRow[] = Array.from({ length: 80 }, (_, index) => {
    const id = index + 1;
    return [`WS-${pad(id)}`, `phase-${id % 5}`, id * 10];
  });

  const templatePath = path.join(rootDir, 'writeback-template.xlsx');
  const outputPath = path.join(rootDir, 'writeback-output.xlsx');
  await writeWorkbook(templatePath, [
    { name: mainSheet, rows: [...mainHeaders, ...mainBody] },
    { name: sideSheet, rows: [...sideHeaders, ...sideBody] },
  ]);

  const deleteMainIds = [5, 15, 25, 35, 45, 55, 65, 75];
  const insertMainRows = [
    {
      afterId: 'WB-0010',
      values: ['WB-ADD-001', 'inserted-001', 900001, 'inserted-a', 'theirs-inserted-001'] as SheetRow,
    },
    {
      afterId: 'WB-0020',
      values: ['WB-ADD-002', 'inserted-002', 900002, 'inserted-b', 'theirs-inserted-002'] as SheetRow,
    },
    {
      afterId: 'WB-0050',
      values: ['WB-ADD-003', 'inserted-003', 900003, 'inserted-c', 'theirs-inserted-003'] as SheetRow,
    },
    {
      afterId: 'WB-0080',
      values: ['WB-ADD-004', 'inserted-004', 900004, 'inserted-d', 'theirs-inserted-004'] as SheetRow,
    },
    {
      afterId: 'WB-0100',
      values: ['WB-ADD-005', 'inserted-005', 900005, 'inserted-e', 'theirs-inserted-005'] as SheetRow,
    },
    {
      afterId: 'WB-0120',
      values: ['WB-ADD-006', 'inserted-006', 900006, 'inserted-f', 'theirs-inserted-006'] as SheetRow,
    },
  ];
  const amountEditIds = [12, 24, 36, 48, 60, 72, 84, 96, 108, 120];
  const statusEditIds = [18, 30, 42, 54, 66, 78, 90, 102, 114];
  const mainInsertedColumnValues: (string | number | null)[] = [
    'catalog:theirs_only',
    'group-6',
    'theirs_only',
    ...mainBody.map((row) => `theirs-${row[0]}`),
  ];
  const saveReq: Parameters<typeof __testOnly.saveMergeResultDirect>[0] = {
    templatePath,
    colOps: [
      {
        sheetName: mainSheet,
        action: 'delete',
        targetColNumber: 3,
        alignedColNumber: 3,
        source: 'theirs',
      },
      {
        sheetName: mainSheet,
        action: 'insert',
        targetColNumber: 6,
        alignedColNumber: 6,
        source: 'theirs',
        values: mainInsertedColumnValues,
      },
    ],
    rowOps: [
      ...deleteMainIds.map((id) => ({
        sheetName: mainSheet,
        action: 'delete' as const,
        targetRowNumber: HEADER_ROW_COUNT + id,
        visualRowNumber: HEADER_ROW_COUNT + id,
      })),
      ...insertMainRows.map((insertRow, index) => {
        const afterId = Number(String(insertRow.afterId).slice(-4));
        return {
          sheetName: mainSheet,
          action: 'insert' as const,
          targetRowNumber: HEADER_ROW_COUNT + afterId + 1,
          visualRowNumber: HEADER_ROW_COUNT + afterId + 1,
          values: insertRow.values,
        };
      }),
    ],
    cells: [
      ...amountEditIds.map((id, index) => ({
        sheetName: mainSheet,
        address: makeAddress(4, HEADER_ROW_COUNT + id),
        value: 700000 + index,
      })),
      ...statusEditIds.map((id, index) => ({
        sheetName: mainSheet,
        address: makeAddress(5, HEADER_ROW_COUNT + id),
        value: `merged-status-${pad(index + 1, 3)}`,
      })),
      ...sideBody
        .map((row, index) => ({ row, index: index + 1 }))
        .flatMap(({ row, index }) => {
          const edits: Array<{ sheetName: string; address: string; value: string | number | null }> = [];
          if (index % 7 === 0) {
            edits.push({
              sheetName: sideSheet,
              address: makeAddress(2, HEADER_ROW_COUNT + index),
              value: `phase-merged-${pad(index, 3)}`,
            });
          }
          if (index % 5 === 0) {
            edits.push({
              sheetName: sideSheet,
              address: makeAddress(3, HEADER_ROW_COUNT + index),
              value: 800000 + index,
            });
          }
          return edits;
        }),
    ],
  };

  const saveResult = await __testOnly.saveMergeResultDirect(saveReq, outputPath);
  assert.equal(saveResult.success, true);
  assert.equal(saveResult.filePath, outputPath);

  const expectedMainHeaders: SheetRow[] = [
    ['catalog:id', 'catalog:label', 'catalog:amount', 'catalog:status', 'catalog:theirs_only'],
    ['group-1', 'group-2', 'group-4', 'group-5', 'group-6'],
    ['id', 'label', 'amount', 'status', 'theirs_only'],
  ];
  const expectedMainBody = mainBody
    .map((row) => [row[0], row[1], row[3], row[4], `theirs-${row[0]}`] as SheetRow)
    .filter((row) => !deleteMainIds.includes(Number(String(row[0]).slice(-4))));
  insertMainRows.forEach((insertRow) => {
    const idx = expectedMainBody.findIndex((row) => row[0] === insertRow.afterId);
    if (idx >= 0) expectedMainBody.splice(idx + 1, 0, [...insertRow.values]);
    else expectedMainBody.push([...insertRow.values]);
  });
  amountEditIds.forEach((id, index) => {
    const row = expectedMainBody.find((item) => item[0] === `WB-${pad(id)}`);
    assert.ok(row, `Expected amount edit row WB-${pad(id)} to exist`);
    row[2] = 700000 + index;
  });
  statusEditIds.forEach((id, index) => {
    const row = expectedMainBody.find((item) => item[0] === `WB-${pad(id)}`);
    assert.ok(row, `Expected status edit row WB-${pad(id)} to exist`);
    row[3] = `merged-status-${pad(index + 1, 3)}`;
  });

  const expectedSideRows: SheetRow[] = [...sideHeaders, ...cloneRows(sideBody)];
  for (let index = 1; index <= sideBody.length; index += 1) {
    const actualRow = expectedSideRows[HEADER_ROW_COUNT + index - 1];
    if (index % 7 === 0) actualRow[1] = `phase-merged-${pad(index, 3)}`;
    if (index % 5 === 0) actualRow[2] = 800000 + index;
  }

  const actualMainRows = await readWorkbookRows(outputPath, mainSheet);
  const actualSideRows = await readWorkbookRows(outputPath, sideSheet);
  assert.deepStrictEqual(actualMainRows, [...expectedMainHeaders, ...expectedMainBody]);
  assert.deepStrictEqual(actualSideRows, expectedSideRows);
};

const runSeededSaveRoundTripScenario = async (rootDir: string) => {
  const cases = [
    buildSeededKeyedCase(1101, 220, 'SeededRoundTrip1101'),
    buildSeededKeyedCase(2203, 260, 'SeededRoundTrip2203'),
    buildSeededKeyedCase(3307, 300, 'SeededRoundTrip3307'),
  ];

  for (const seededCase of cases) {
    const headerRows = makeHeaderRows(seededCase.headers);
    const { basePath, oursPath, theirsPath } = await writeWorkbookTriplet(
      rootDir,
      seededCase.sheetName,
      [{ name: seededCase.sheetName, rows: [...headerRows, ...seededCase.baseRows] }],
      [{ name: seededCase.sheetName, rows: [...headerRows, ...seededCase.oursRows] }],
      [{ name: seededCase.sheetName, rows: [...headerRows, ...seededCase.theirsRows] }],
    );
    const { mergeSheets } = await __testOnly.buildMergeSheetsForWorkbooks(
      basePath,
      oursPath,
      theirsPath,
      1,
      HEADER_ROW_COUNT,
      LARGE_ROW_SIMILARITY_THRESHOLD,
      'merge',
    );
    const sheet = getSheetOrThrow(mergeSheets, seededCase.sheetName);
    const bodyRows = (sheet.rowsMeta ?? [])
      .filter((row) => row.visualRowNumber > HEADER_ROW_COUNT)
      .sort((a, b) => a.visualRowNumber - b.visualRowNumber);
    const metaByVisual = new Map(bodyRows.map((row) => [row.visualRowNumber, row]));

    const rowOps: NonNullable<Parameters<typeof __testOnly.saveMergeResultDirect>[0]['rowOps']> = [];
    bodyRows.forEach((row) => {
      if (!row.oursRowNumber && row.theirsRowNumber) {
        const key = row.key ?? '';
        const values = seededCase.theirsById.get(String(key));
        assert.ok(values, `Expected inserted theirs row ${String(key)} in ${seededCase.sheetName}`);
        rowOps.push({
          sheetName: seededCase.sheetName,
          action: 'insert',
          targetRowNumber: computeInsertTargetRowNumber(bodyRows, row.visualRowNumber),
          visualRowNumber: row.visualRowNumber,
          values: [...values],
        });
        return;
      }
      if (row.oursRowNumber && !row.theirsRowNumber) {
        rowOps.push({
          sheetName: seededCase.sheetName,
          action: 'delete',
          targetRowNumber: row.oursRowNumber,
          visualRowNumber: row.visualRowNumber,
        });
      }
    });

    const cells = sheet.cells
      .filter((cell) => {
        if (cell.row <= HEADER_ROW_COUNT) return false;
        const meta = metaByVisual.get(cell.row);
        return !!meta?.oursRowNumber && !!meta?.theirsRowNumber;
      })
      .map((cell) => ({
        sheetName: seededCase.sheetName,
        address: cell.address,
        value: cell.theirsValue,
      }));

    const outputPath = path.join(rootDir, `${seededCase.sheetName}.roundtrip.xlsx`);
    const saveResult = await __testOnly.saveMergeResultDirect(
      {
        templatePath: oursPath,
        cells,
        rowOps,
      },
      outputPath,
    );
    assert.equal(saveResult.success, true, `${seededCase.sheetName} save should succeed`);

    const actualRows = await readWorkbookRows(outputPath, seededCase.sheetName);
    assert.deepStrictEqual(actualRows, [...headerRows, ...seededCase.theirsRows], `${seededCase.sheetName} output should match theirs exactly`);
  }
};

const main = async () => {
  await app.whenReady();
  const rootDir = await fs.mkdtemp(path.join(os.tmpdir(), 'electron-excel-merge-threeway-regression-'));
  try {
    await runCellStatusLargeScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runModeDifferenceLargeScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runStructuredHeaderDetectionScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runStructuralRowsLargeScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runStructuralColumnsLargeScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runNoKeyLargeScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runSeededRandomBatchScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runMultiSheetLargeScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runSaveMergeWriteBackScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await runSeededSaveRoundTripScenario(rootDir);
    __testOnly.clearWorkbookCache();
    await fs.rm(rootDir, { recursive: true, force: true });
    console.log('threeWayDiff regression test passed');
    app.exit(0);
  } catch (error) {
    __testOnly.clearWorkbookCache();
    await fs.rm(rootDir, { recursive: true, force: true });
    console.error(error);
    app.exit(1);
  }
};

void main();
