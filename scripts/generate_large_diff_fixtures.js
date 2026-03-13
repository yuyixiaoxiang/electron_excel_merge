const fs = require('node:fs/promises');
const path = require('node:path');
const ExcelJS = require('exceljs');

const outputDir = path.resolve(__dirname, '..', 'generated', 'manual-diff-fixtures');
const leftPath = path.join(outputDir, 'large_diff_left.xlsx');
const rightPath = path.join(outputDir, 'large_diff_right.xlsx');

const pad = (value, width = 5) => String(value).padStart(width, '0');

const createPrng = (seed) => {
  let state = seed >>> 0;
  return () => {
    state += 0x6d2b79f5;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
};

const shuffle = (items, rng) => {
  const arr = items.slice();
  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(rng() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
};

const pickMany = (allIds, count, rng) => {
  const shuffled = shuffle(allIds, rng);
  return new Set(shuffled.slice(0, Math.min(count, shuffled.length)));
};

const insertEvery = (rows, inserts, every) => {
  const result = [];
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

const cloneRows = (rows) => rows.map((row) => [...row]);

const applyWorksheetStyle = (ws, colCount) => {
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 3 }];
  for (let col = 1; col <= colCount; col += 1) {
    ws.getColumn(col).width = col === 1 ? 18 : col <= 4 ? 16 : 14;
  }
};

const addSheet = (workbook, name, rows) => {
  const ws = workbook.addWorksheet(name);
  rows.forEach((row) => ws.addRow(row));
  applyWorksheetStyle(ws, Math.max(...rows.map((row) => row.length)));
};

const makeWorkbook = () => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Oz';
  wb.created = new Date('2026-03-11T00:00:00Z');
  wb.modified = new Date('2026-03-11T00:00:00Z');
  return wb;
};

const ordersScenario = () => {
  const rng = createPrng(101);
  const headerRows = [
    ['domain:orders', 'domain:orders', 'domain:orders', 'domain:orders', 'domain:orders', 'domain:orders', 'domain:orders', 'domain:orders', 'domain:orders', 'domain:orders'],
    ['identity', 'timeline', 'geo', 'customer', 'line', 'metrics', 'metrics', 'metrics', 'workflow', 'workflow'],
    ['ORDER_ID', 'ORDER_DATE', 'REGION', 'CUSTOMER', 'CATEGORY', 'QTY', 'UNIT_PRICE', 'AMOUNT', 'STATUS', 'NOTE'],
  ];
  const baseRows = Array.from({ length: 2400 }, (_, index) => {
    const id = index + 1;
    const qty = 5 + (id % 19);
    const unitPrice = 50 + ((id * 7) % 130);
    return [
      `ORD-${pad(id)}`,
      `2026-02-${pad((id % 28) + 1, 2)}`,
      ['North', 'South', 'West', 'East'][id % 4],
      `Customer-${pad((id * 3) % 700, 4)}`,
      ['Hardware', 'Software', 'Service', 'Bundle'][id % 4],
      qty,
      unitPrice,
      qty * unitPrice,
      ['Open', 'Packed', 'Ready', 'Closed'][id % 4],
      `note-${pad(id)}`,
    ];
  });

  const rightRows = cloneRows(baseRows);
  const ids = baseRows.map((row) => row[0]);
  const deleteIds = pickMany(ids, 140, rng);
  const remainingAfterDelete = ids.filter((id) => !deleteIds.has(id));
  const modifyIds = pickMany(remainingAfterDelete, 420, rng);
  const keepRows = rightRows.filter((row) => !deleteIds.has(row[0]));

  keepRows.forEach((row, index) => {
    if (!modifyIds.has(row[0])) return;
    const nextQty = row[5] + 2 + (index % 3);
    row[5] = nextQty;
    row[7] = nextQty * row[6];
    row[8] = ['Repriced', 'Delayed', 'Review', 'Escalated'][index % 4];
    row[9] = `changed-${row[0]}`;
  });

  const inserts = Array.from({ length: 180 }, (_, index) => {
    const seq = index + 1;
    const qty = 7 + (seq % 11);
    const unitPrice = 90 + ((seq * 13) % 170);
    return [
      `ORD-NEW-${pad(seq, 4)}`,
      `2026-03-${pad((seq % 28) + 1, 2)}`,
      ['North', 'South', 'West', 'East'][seq % 4],
      `Customer-New-${pad(seq, 4)}`,
      ['Hardware', 'Software', 'Service', 'Bundle'][seq % 4],
      qty,
      unitPrice,
      qty * unitPrice,
      'Inserted',
      `inserted-${pad(seq, 4)}`,
    ];
  });

  return {
    name: 'Orders_Q1',
    left: [...headerRows, ...baseRows],
    right: [...headerRows, ...insertEvery(keepRows, inserts, 12)],
  };
};

const inventoryScenario = () => {
  const rng = createPrng(202);
  const leftHeaders = [
    ['domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory'],
    ['identity', 'geo', 'metrics', 'metrics', 'policy', 'owner', 'audit', 'audit'],
    ['SKU', 'WAREHOUSE', 'ON_HAND', 'RESERVED', 'SAFETY_STOCK', 'BUYER', 'LAST_AUDIT', 'COMMENT'],
  ];
  const rightHeaders = [
    ['domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory', 'domain:inventory'],
    ['identity', 'geo', 'metrics', 'metrics', 'policy', 'policy', 'owner', 'audit', 'audit'],
    ['SKU', 'WAREHOUSE', 'ON_HAND', 'RESERVED', 'SAFETY_STOCK', 'CYCLE_TAG', 'BUYER', 'LAST_AUDIT', 'COMMENT'],
  ];

  const baseRows = Array.from({ length: 1850 }, (_, index) => {
    const id = index + 1;
    return [
      `SKU-${pad(id)}`,
      `WH-${pad((id % 24) + 1, 2)}`,
      100 + ((id * 17) % 900),
      10 + ((id * 7) % 120),
      40 + ((id * 5) % 160),
      `Buyer-${pad((id * 3) % 120, 3)}`,
      `2026-01-${pad((id % 28) + 1, 2)}`,
      `inventory-note-${pad(id)}`,
    ];
  });

  const rightRows = baseRows.map((row) => [row[0], row[1], row[2], row[3], row[4], `C-${row[0].slice(-2)}`, row[5], row[6], row[7]]);
  const ids = baseRows.map((row) => row[0]);
  const deleteIds = pickMany(ids, 95, rng);
  const remainingAfterDelete = ids.filter((id) => !deleteIds.has(id));
  const modifyIds = pickMany(remainingAfterDelete, 360, rng);
  const filtered = rightRows.filter((row) => !deleteIds.has(row[0]));

  filtered.forEach((row, index) => {
    if (!modifyIds.has(row[0])) return;
    row[2] = row[2] + 25 + (index % 17);
    row[3] = Math.max(0, row[3] - (index % 9));
    row[5] = `R${pad((index % 40) + 1, 2)}`;
    row[8] = `inventory-updated-${row[0]}`;
  });

  const inserts = Array.from({ length: 130 }, (_, index) => {
    const seq = index + 1;
    return [
      `SKU-NEW-${pad(seq, 4)}`,
      `WH-${pad((seq % 24) + 1, 2)}`,
      300 + ((seq * 19) % 700),
      12 + ((seq * 5) % 80),
      50 + ((seq * 7) % 130),
      `N${pad((seq % 55) + 1, 2)}`,
      `Buyer-New-${pad(seq, 3)}`,
      `2026-03-${pad((seq % 28) + 1, 2)}`,
      `inventory-insert-${pad(seq, 4)}`,
    ];
  });

  return {
    name: 'Inventory',
    left: [...leftHeaders, ...baseRows],
    right: [...rightHeaders, ...insertEvery(filtered, inserts, 14)],
  };
};

const noKeyScenario = () => {
  const rng = createPrng(303);
  const headers = [
    ['domain:log', 'domain:log', 'domain:log', 'domain:log', 'domain:log', 'domain:log', 'domain:log', 'domain:log'],
    ['timeline', 'geo', 'service', 'bucket', 'metrics', 'metrics', 'metrics', 'remark'],
    ['DAY', 'REGION', 'SERVICE', 'BUCKET', 'METRIC_A', 'METRIC_B', 'METRIC_C', 'REMARK'],
  ];
  const leftRows = Array.from({ length: 1600 }, (_, index) => {
    const i = index + 1;
    return [
      `2026-02-${pad((i % 28) + 1, 2)}`,
      ['CN', 'US', 'EU', 'APAC'][i % 4],
      ['auth', 'billing', 'risk', 'search'][i % 4],
      `B-${Math.floor(index / 20)}`,
      100 + ((i * 13) % 90),
      200 + ((i * 7) % 110),
      300 + ((i * 11) % 95),
      `remark-${i % 12}`,
    ];
  });
  const rightRows = cloneRows(leftRows);
  const indexes = Array.from({ length: leftRows.length }, (_, index) => index);
  const deleteIndexes = pickMany(indexes, 90, rng);
  const remaining = indexes.filter((idx) => !deleteIndexes.has(idx));
  const modifyIndexes = pickMany(remaining, 260, rng);
  const filtered = rightRows.filter((_, index) => !deleteIndexes.has(index));

  filtered.forEach((row, index) => {
    if (!modifyIndexes.has(index)) return;
    row[4] = row[4] + 11 + (index % 5);
    row[5] = row[5] - (index % 7);
    row[7] = `recomputed-${index % 15}`;
  });

  const inserts = Array.from({ length: 95 }, (_, index) => {
    const seq = index + 1;
    return [
      `2026-03-${pad((seq % 28) + 1, 2)}`,
      ['CN', 'US', 'EU', 'APAC'][seq % 4],
      ['auth', 'billing', 'risk', 'search'][seq % 4],
      `NB-${Math.floor(seq / 5)}`,
      180 + ((seq * 17) % 70),
      220 + ((seq * 9) % 60),
      330 + ((seq * 7) % 55),
      `inserted-nokey-${seq % 9}`,
    ];
  });

  return {
    name: 'PriceHistory_NoKey',
    left: [...headers, ...leftRows],
    right: [...headers, ...insertEvery(filtered, inserts, 17)],
  };
};

const formulaScenario = () => {
  const headers = [
    ['domain:formula', 'domain:formula', 'domain:formula', 'domain:formula', 'domain:formula', 'domain:formula'],
    ['identity', 'input', 'input', 'calc', 'calc', 'calc'],
    ['ROW_ID', 'QTY', 'UNIT_PRICE', 'SUBTOTAL', 'TAX', 'TOTAL'],
  ];

  const buildRows = (variant) =>
    Array.from({ length: 920 }, (_, index) => {
      const id = index + 1;
      let qty = 5 + (id % 15);
      let unitPrice = 80 + ((id * 9) % 140);
      if (variant === 'right' && id % 11 === 0) qty += 3;
      if (variant === 'right' && id % 17 === 0) unitPrice += 7;
      const subtotal = qty * unitPrice;
      const tax = Math.round(subtotal * 0.13);
      const rowNumber = index + 4;
      return [
        `F-${pad(id)}`,
        qty,
        unitPrice,
        { formula: `B${rowNumber}*C${rowNumber}`, result: subtotal },
        { formula: `ROUND(D${rowNumber}*0.13,0)`, result: tax },
        { formula: `D${rowNumber}+E${rowNumber}`, result: subtotal + tax },
      ];
    });

  return {
    name: 'Summary_Formulas',
    left: [...headers, ...buildRows('left')],
    right: [...headers, ...buildRows('right')],
  };
};

const buildWorkbooks = async () => {
  await fs.mkdir(outputDir, { recursive: true });
  const scenarios = [ordersScenario(), inventoryScenario(), noKeyScenario(), formulaScenario()];

  const leftWorkbook = makeWorkbook();
  const rightWorkbook = makeWorkbook();

  scenarios.forEach((scenario) => {
    addSheet(leftWorkbook, scenario.name, scenario.left);
    addSheet(rightWorkbook, scenario.name, scenario.right);
  });

  await leftWorkbook.xlsx.writeFile(leftPath);
  await rightWorkbook.xlsx.writeFile(rightPath);
};

const main = async () => {
  await buildWorkbooks();
  console.log(`Generated: ${leftPath}`);
  console.log(`Generated: ${rightPath}`);
};

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
