# 代码审查报告

审查日期：2026-02-05  
审查范围：整个代码库

## 🔴 严重问题 (Critical)

### 1. 列插入操作的索引计算问题

**位置**: `App.tsx` 第 464-476 行

**问题描述**:
```typescript
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
    effectiveColMap.splice(insertIdx, 0, { alignedCol: ac, oursCol: null });
  }
}
```

**潜在问题**:
- 如果有多个插入列，每次 `splice` 会改变后续索引
- 可能导致插入位置不正确

**修复建议**:
```typescript
// 先收集所有插入位置和列信息，最后一次性插入
const insertions: Array<{ idx: number; col: number }> = [];
for (const ac of insertedAlignedCols) {
  const meta = mergeColumnsMeta.find((m) => m.col === ac);
  if (meta && !meta.oursCol && meta.theirsCol) {
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
// 从后往前插入，避免索引变化
insertions.sort((a, b) => b.idx - a.idx);
for (const ins of insertions) {
  effectiveColMap.splice(ins.idx, 0, { alignedCol: ins.col, oursCol: null });
}
```

**影响**: 中 - 可能导致列顺序错乱

---

### 2. 保存时列操作与单元格修改的顺序问题

**位置**: `main.ts` 第 2257-2366 行

**问题描述**:
保存逻辑的执行顺序是：
1. 修改单元格值 (2257-2261)
2. 执行列操作 (2262-2328)
3. 执行行操作 (2330-2365)

**潜在问题**:
- 单元格修改使用的 address（如 "B5"）是基于原始文件的
- 列操作会改变列索引，导致后续的地址可能不正确
- 例如：删除 A 列后，原来的 B5 应该变成 A5，但代码中还是用 B5

**修复建议**:
```typescript
// 应该先执行列/行操作，再修改单元格
// 1. 列操作
// 2. 行操作
// 3. 单元格修改（需要根据操作调整 address）
```

**影响**: 高 - 可能导致数据写入错误的单元格

---

### 3. `buildMergedColumnValues` 中的行数不匹配

**位置**: `App.tsx` 第 843-862 行

**问题描述**:
```typescript
result.rows.forEach((rowRes: any) => {
  // ...
  columnValues.push(value);
});
return columnValues;
```

**潜在问题**:
- 这里收集的是"有效行"的列值（跳过了 deleted 行）
- 但保存时 `op.values` 应该包含**所有行**的值（包括将被删除的行）
- 因为保存是先处理列操作，此时行还没被删除

**修复建议**:
```typescript
// 不应该跳过任何行，应该收集所有行的列值
result.rows.forEach((rowRes: any, idx: number) => {
  const visualRowNumber = rowRes.rowNumber ?? 0;
  // 不要根据 rowOp 跳过行
  // if (oursMissing && rowOp?.action !== 'insert') return;  // 删除这些逻辑
  // if (!oursMissing && rowOp?.action === 'delete') return;
  
  // 直接收集值
  const value = ...;
  columnValues.push(value);
});
```

**影响**: 高 - 导致插入列的数据行数不对

---

## 🟡 中等问题 (Medium)

### 4. 合并预览中对插入行的处理

**位置**: `App.tsx` 第 499-501 行

```typescript
if (op?.action === 'insert' && op.values) {
  mergedRow.push(op.values[alignedCol - 1] ?? null);
}
```

**潜在问题**:
- `op.values` 是按对齐后的列顺序存储的
- 但如果有列被删除，`alignedCol` 可能不等于 `effectiveColMap` 的索引
- 应该使用 `effectiveColMap` 的索引而不是 `alignedCol`

**修复建议**:
```typescript
// 使用索引而不是 alignedCol
for (let i = 0; i < effectiveColMap.length; i += 1) {
  const colInfo = effectiveColMap[i];
  const alignedCol = colInfo.alignedCol;
  // ...
  if (op?.action === 'insert' && op.values) {
    mergedRow.push(op.values[i] ?? null);  // 使用 i 而不是 alignedCol - 1
  }
}
```

**影响**: 中 - 插入行的预览可能不正确

---

### 5. 切换工作表时 colOps/rowOps 未清空预览

**位置**: `App.tsx` 第 205-222 行

```typescript
setSelectedMergeSheetIndex(nextIndex);
setMergeCells(allMergeSheets[nextIndex]?.cells ?? []);
setMergeRowsMeta(allMergeSheets[nextIndex]?.rowsMeta ?? []);
setMergeColumnsMeta(allMergeSheets[nextIndex]?.columnsMeta ?? []);
```

**潜在问题**:
- 切换工作表时，`currentRowOps` 和 `currentColOps` 会自动更新
- 但 `mergedPreviewRows` 的 useEffect 依赖这些值
- 如果依赖没触发，预览可能显示旧数据

**修复建议**:
添加立即清空预览：
```typescript
setMergedPreviewRows([]);
setMergedPreviewRowVisuals([]);
```

**影响**: 低 - useEffect 通常会正确触发，但边界情况可能有问题

---

### 6. 列对齐时对空列的处理

**位置**: `main.ts` 第 298-300 行

```typescript
const isFullyEmpty = !headerText && !hasDataSample;
if (isFullyEmpty) continue;
```

**潜在问题**:
- 空列会被跳过，不生成 ColumnRecord
- 如果 base 有空列，ours 在同位置有数据列，可能无法正确对齐
- 因为 base 的列被跳过了，列号会错位

**修复建议**:
考虑保留空列，或者在对齐算法中处理列号偏移

**影响**: 低 - 实际场景中很少有完全空的列

---

## 🟢 轻微问题 (Minor)

### 7. 内存泄漏风险 - workbookCache 无上限

**位置**: `main.ts` 工作簿缓存

```typescript
const workbookCache = new Map<string, Workbook>();
const loadWorkbookCached = async (filePath: string) => {
  if (workbookCache.has(filePath)) {
    return workbookCache.get(filePath)!;
  }
  // ...
  workbookCache.set(filePath, wb);
  return wb;
};
```

**潜在问题**:
- 缓存无上限，长时间运行可能内存溢出
- 文件被修改后，缓存的工作簿可能过期

**修复建议**:
```typescript
const MAX_CACHE_SIZE = 10;
const cacheAccessOrder: string[] = [];

const loadWorkbookCached = async (filePath: string) => {
  if (workbookCache.has(filePath)) {
    // Update access order (LRU)
    const idx = cacheAccessOrder.indexOf(filePath);
    if (idx >= 0) cacheAccessOrder.splice(idx, 1);
    cacheAccessOrder.push(filePath);
    return workbookCache.get(filePath)!;
  }
  
  // Evict oldest if cache full
  if (workbookCache.size >= MAX_CACHE_SIZE) {
    const oldest = cacheAccessOrder.shift();
    if (oldest) workbookCache.delete(oldest);
  }
  
  const wb = new Workbook();
  await wb.xlsx.readFile(filePath);
  workbookCache.set(filePath, wb);
  cacheAccessOrder.push(filePath);
  return wb;
};
```

**影响**: 低 - 通常不会缓存太多文件

---

### 8. 错误处理不完善

**位置**: 多处

**问题描述**:
- 很多 async 函数没有 try-catch
- 错误信息对用户不友好

**例子**: `App.tsx` handleOpenThreeWay
```typescript
const handleOpenThreeWay = useCallback(async () => {
  const result = await window.excelAPI.openThreeWay();  // 无错误处理
  if (!result) return;
  // ...
}, []);
```

**修复建议**:
```typescript
const handleOpenThreeWay = useCallback(async () => {
  try {
    const result = await window.excelAPI.openThreeWay();
    if (!result) return;
    // ...
  } catch (error) {
    console.error('Failed to open three-way merge:', error);
    alert('打开文件失败：' + (error as Error).message);
  }
}, []);
```

**影响**: 低 - 但会影响用户体验

---

### 9. 主键列的映射问题

**位置**: `App.tsx` 第 96-100 行

```typescript
const displayPrimaryKeyCol = useMemo(() => {
  if (typeof primaryKeyCol !== 'number' || primaryKeyCol < 1) return primaryKeyCol;
  const hit = mergeColumnsMeta.find((c) => c.oursCol === primaryKeyCol);
  return hit ? hit.col : primaryKeyCol;
}, [primaryKeyCol, mergeColumnsMeta]);
```

**潜在问题**:
- 用户设置的 `primaryKeyCol` 是 ours 的物理列号
- 但需要转换为对齐后的逻辑列号才能正确显示
- 如果找不到映射，直接返回 `primaryKeyCol` 可能不正确

**修复建议**:
```typescript
const displayPrimaryKeyCol = useMemo(() => {
  if (typeof primaryKeyCol !== 'number' || primaryKeyCol < 1) return -1;
  const hit = mergeColumnsMeta.find((c) => c.oursCol === primaryKeyCol);
  if (!hit) {
    console.warn('Primary key column not found in aligned columns');
    return -1;  // 明确返回无效值
  }
  return hit.col;
}, [primaryKeyCol, mergeColumnsMeta]);
```

**影响**: 低 - 主键列通常不会被删除

---

## 📋 代码质量问题

### 10. 魔法数字

**位置**: 多处

```typescript
// 相似度阈值
const threshold = 0.55;
const headerThreshold = 0.8;

// 权重
const wHeader = hasHeader ? 0.6 : 0.2;
const wType = 0.2;
const wVal = 0.2;
```

**修复建议**:
定义常量：
```typescript
const COLUMN_SIMILARITY_THRESHOLD = 0.55;
const HEADER_SIMILARITY_THRESHOLD = 0.8;
const HEADER_WEIGHT = 0.6;
const TYPE_WEIGHT = 0.2;
const VALUE_WEIGHT = 0.2;
```

**影响**: 无 - 仅代码质量问题

---

### 11. 类型断言过多

**位置**: 多处使用 `as any`

```typescript
cell.value = cellInfo.value as any;
```

**修复建议**:
使用更精确的类型：
```typescript
cell.value = cellInfo.value as CellValue;
```

**影响**: 无 - 仅代码质量问题

---

## 🔍 需要测试的边界情况

1. **空文件**: 所有工作表都是空的
2. **单列文件**: 只有一列数据
3. **超大文件**: 10000+ 行
4. **重复主键**: 多行有相同的主键值
5. **列完全重排**: ours 和 theirs 的列顺序完全不同
6. **多工作表**: 每个工作表的列数不同
7. **混合操作**: 同时有行插入、行删除、列插入、列删除和单元格修改

---

## ✅ 优先修复建议

1. **立即修复** (Critical):
   - 问题 #2: 保存时的操作顺序
   - 问题 #3: buildMergedColumnValues 的行过滤问题

2. **尽快修复** (High):
   - 问题 #1: 列插入索引计算
   - 问题 #4: 插入行预览的列索引

3. **计划修复** (Medium):
   - 问题 #7: 工作簿缓存 LRU
   - 问题 #8: 错误处理

4. **优化** (Low):
   - 其他代码质量问题

---

## 📝 测试建议

创建单元测试覆盖：
1. 列对齐算法
2. 行对齐算法
3. 保存逻辑（特别是列/行操作）
4. 边界情况

创建集成测试：
1. 完整的合并流程
2. 多工作表场景
3. 复杂的列/行操作组合
