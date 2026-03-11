# 代码库详细文档

本文档详细解释了 Excel 三方合并工具的代码结构和核心逻辑。

## 📁 文件结构

```
src/
├── main/                    # 主进程（Electron 后端）
│   ├── main.ts             # 主进程入口，Excel 读写和三方 diff/merge 核心逻辑
│   └── preload.ts          # 预加载脚本，定义 IPC 通信接口
└── renderer/               # 渲染进程（React 前端）
    ├── App.tsx             # 主应用组件，状态管理和业务逻辑
    ├── MergeSideBySide.tsx # 左右对比视图组件
    ├── VirtualGrid.tsx     # 虚拟滚动表格组件（性能优化）
    ├── ExcelTable.tsx      # 单文件编辑表格组件
    └── index.tsx           # 渲染进程入口
```

## 🏗️ 核心架构

### 1. 主进程 (main.ts)

主进程负责：
- 创建 Electron 窗口
- 解析命令行参数（git/Fork 传入的文件路径）
- 读写 Excel 文件（使用 ExcelJS 库）
- 执行三方 diff 和 merge 算法
- 通过 IPC 与渲染进程通信

#### 关键概念

**三方合并模式**：
- `base`: 共同祖先版本
- `ours`: 当前分支版本（本地修改）
- `theirs`: 对方分支版本（远程修改）
- `merged`: 最终合并结果

**两种启动模式**：
1. **diff 模式** (2个参数): `app.exe OURS THEIRS`
   - 仅用于查看差异，base = ours
2. **merge 模式** (3-4个参数): `app.exe BASE OURS THEIRS [MERGED]`
   - 完整三方合并，可指定输出文件

### 2. 渲染进程 (App.tsx)

渲染进程负责：
- UI 展示和用户交互
- 管理合并状态（已解决的冲突、用户选择等）
- 构建合并预览
- 触发保存操作

## 🔧 核心算法详解

### 一、列对齐算法 (Column Alignment)

**目的**：当 base/ours/theirs 的列不一致时（插入、删除、重排），智能匹配相同语义的列。

#### 1.1 列特征提取 (`buildColumnRecords`)

为每一列提取特征：

```typescript
interface ColumnRecord {
  colNumber: number;        // 列号（1-based，如A=1, B=2）
  headerText: string;       // 表头文本（前N行拼接，如 "icon|名称|string"）
  headerKey: string;        // 严格标准化的表头键（去空格、特殊字符，如 "icon名称string"）
  typeSig: {                // 类型签名（数据类型分布）
    num: number;            // 数字类型的单元格数量
    str: number;            // 字符串类型数量
    empty: number;          // 空单元格数量
    other: number;          // 其他类型数量
  };
  sampleValues: string[];   // 样本值（用于相似度计算）
}
```

**为什么需要这些特征？**
- `headerText`: 人类可读的表头
- `headerKey`: 用于精确匹配（忽略空格、标点等差异）
- `typeSig`: 判断列的数据类型是否一致
- `sampleValues`: 判断列的数据内容是否相似

#### 1.2 列对齐过程 (`buildAlignedColumns`)

**步骤**：

1. **提取列特征**
   ```
   base:   [A: "id|主键", B: "name|名称", C: "desc|描述"]
   ours:   [A: "id|主键", B: "name|名称", C: "新列X", D: "desc|描述"]
   theirs: [A: "id|主键", B: "name|名称", C: "desc|描述"]
   ```

2. **LCS 锚点匹配** (`lcsMatchPairs`)
   - 找到最长公共子序列作为对齐"锚点"
   - 例如：base 的 A,B,C 与 ours 的 A,B,D 匹配
   ```
   锚点: (base.A ↔ ours.A), (base.B ↔ ours.B), (base.C ↔ ours.D)
   ```

3. **相似度匹配** (`alignColumnsBySimilarity`)
   - 对锚点之间的"缝隙"进行相似度匹配
   - 计算列相似度 = 0.6 * 表头相似度 + 0.2 * 类型相似度 + 0.2 * 样本相似度
   - 阈值：相似度 >= 0.55 才认为是同一列

4. **生成对齐结果** (`AlignedColumn[]`)
   ```typescript
   [
     { baseCol: 1, oursCol: 1, theirsCol: 1 },  // A列：三方都有
     { baseCol: 2, oursCol: 2, theirsCol: 2 },  // B列：三方都有
     { baseCol: null, oursCol: 3, theirsCol: null },  // 新列X：只有ours
     { baseCol: 3, oursCol: 4, theirsCol: 3 }   // desc列：三方都有但位置不同
   ]
   ```

**关键点**：
- 对齐后的列号是"逻辑列号"，用于统一表示
- 每一行的 `oursCol/theirsCol` 指向实际文件中的物理列号
- 这样即使列顺序不同，也能正确比较对应的单元格

### 二、行对齐算法 (Row Alignment)

**目的**：当 base/ours/theirs 的行不一致时（插入、删除、移动），智能匹配相同的行。

#### 2.1 基于主键的对齐 (`alignRowsByKey`)

如果指定了主键列（例如第1列是ID），直接用主键值匹配：

```typescript
base:   { key: "101", row: 1 }
ours:   { key: "101", row: 2 }  // 行号变了，但主键相同
theirs: { key: "101", row: 1 }

→ 匹配: (base.row1 ↔ ours.row2 ↔ theirs.row1)
```

**优点**：精确、快速
**缺点**：需要稳定的主键列

#### 2.2 基于相似度的对齐 (`alignRowsBySimilarity`)

没有主键时，使用行内容相似度匹配：

1. **提取行特征**
   - 非空单元格的列号列表
   - 所有单元格值的拼接字符串

2. **LCS 锚点** + **相似度匹配**
   - 类似列对齐的方式
   - 行相似度计算：Levenshtein 距离 + Jaccard 相似度

3. **歧义检测**
   - 如果一行与多行相似度都很高 → 标记为 `ambiguous`
   - UI 会用特殊颜色提示用户

**相似度阈值**：默认 0.9（可调整）

#### 2.3 合并对齐结果 (`mergeAlignedRows`)

```typescript
interface AlignedRow {
  base?: RowRecord | null;      // base 的行记录
  ours?: RowRecord | null;      // ours 的行记录
  theirs?: RowRecord | null;    // theirs 的行记录
  key?: string | null;          // 主键值（如果有）
  ambiguousOurs?: boolean;      // ours 匹配有歧义
  ambiguousTheirs?: boolean;    // theirs 匹配有歧义
}
```

**行状态判断**：
- `unchanged`: base = ours = theirs
- `ours-changed`: base ≠ ours, base = theirs
- `theirs-changed`: base = ours, base ≠ theirs
- `both-changed-same`: base ≠ ours = theirs（双方改成相同值）
- `conflict`: base ≠ ours ≠ theirs（真正的冲突）

### 三、单元格级别 Diff

对于每个对齐后的行，逐列比较单元格：

```typescript
interface MergeCell {
  row: number;              // 逻辑行号
  col: number;              // 逻辑列号
  oursCol: number | null;   // ours 的物理列号
  theirsCol: number | null; // theirs 的物理列号
  baseValue: string | number | null;
  oursValue: string | number | null;
  theirsValue: string | number | null;
  status: 'unchanged' | 'ours-changed' | 'theirs-changed' | 'both-changed-same' | 'conflict';
  mergedValue: string | number | null;  // 用户选择或自动合并的值
}
```

**自动合并规则**：
- `unchanged` / `ours-changed` / `theirs-changed` / `both-changed-same` → 自动设置 `mergedValue`
- `conflict` → 需要用户手动选择

## 🎨 UI 交互流程

### 1. 加载文件

```
用户点击"打开三方 Merge/Diff" 
  → 渲染进程调用 window.excelAPI.openThreeWay()
  → 主进程读取 cliThreeWayArgs 或弹出文件选择对话框
  → 主进程执行 buildMergeSheetsForWorkbooks()
  → 返回 MergeSheetData[] 给渲染进程
  → 渲染进程显示左右对比视图
```

### 2. 解决冲突

#### 2.1 单元格级别

```
用户在差异单元格上点击"使用 ours/theirs"
  → handleApplyMergeCellChoice(row, col, source)
  → 更新 mergedValue
  → 标记为 resolved
  → 单元格背景变为灰色（已解决）
```

#### 2.2 整行级别

```
用户在行头右键 → "使用整行数据"
  → handleApplyMergeRowChoice(row, source)
  → 该行所有单元格的 mergedValue 都设置为选择的 source
  → 如果是插入/删除行 → 创建 SaveMergeRowOp
```

#### 2.3 整列级别（新增功能）

```
用户在列头右键 → "使用本列数据"
  → handleApplyMergeColumnChoice(col, source)
  → 该列所有单元格的 mergedValue 都设置为选择的 source
  → 如果是列插入/删除 → 创建 SaveMergeColOp
```

**列操作场景**：
- **ours-only 列 + theirs 侧选择** → 删除该列
  - 用户在 theirs 侧点击"使用本列数据"，但 theirs 没有这列
  - 意味着"不要 ours 的这一列" → 创建删除操作
  
- **theirs-only 列 + theirs 侧选择** → 插入该列
  - 用户在 theirs 侧点击"使用本列数据"，theirs 有这列但 ours 没有
  - 意味着"要 theirs 的这一列" → 创建插入操作

### 3. 合并预览 (Merged Preview)

**实时构建**：
```typescript
useEffect(() => {
  // 获取所有行数据
  const result = await window.excelAPI.getThreeWayRows({...});
  
  // 应用列操作
  const deletedCols = new Set<number>();
  const insertedCols: number[] = [];
  currentColOps.forEach((op, col) => {
    if (op.action === 'delete') deletedCols.add(col);
    else if (op.action === 'insert') insertedCols.push(col);
  });
  
  // 过滤列：排除 deleted，加入 inserted
  const effectiveCols = [...].filter(c => !deletedCols.has(c));
  
  // 应用行操作
  const mergedRows = result.rows.filter(row => {
    // 排除 deleted 行
    if (rowOp?.action === 'delete') return false;
    // 包含 inserted 行
    if (rowOp?.action === 'insert') return true;
    return true;
  });
  
  // 填充每个单元格的值
  for (const col of effectiveCols) {
    if (diffCell) {
      // 优先使用用户选择的值
      row.push(diffCell.mergedValue);
    } else if (colInserted) {
      // 插入列：从 theirs 取值
      row.push(rowRes.theirs[col - 1]);
    } else {
      // 普通列：从 ours 取值
      row.push(rowRes.ours[col - 1]);
    }
  }
  
  setMergedPreviewRows(mergedRows);
}, [currentColOps, currentRowOps, mergeCells]);
```

**关键点**：
- 预览会实时反映用户的所有操作（列插入/删除、行插入/删除、单元格选择）
- 用户可以在保存前预览最终结果

### 4. 保存合并结果

```
用户点击"保存合并结果"
  → 收集所有 mergedValue 不同于原值的单元格
  → 收集所有行操作 (rowOps) 和列操作 (colOps)
  → 调用 window.excelAPI.saveMergeResult({
      templatePath: ours,  // 以 ours 为模板（保留格式）
      cells: [{ sheetName, address, value }],
      rowOps: [{ action: 'insert'|'delete', targetRowNumber, values }],
      colOps: [{ action: 'insert'|'delete', targetColNumber, values }]
    })
  → 主进程执行保存：
      1. 加载 ours 文件
      2. 应用列操作（先删除后插入）
      3. 应用行操作（先删除后插入）
      4. 修改单元格值
      5. 写入目标文件 (MERGED 或 ours)
  → Git merge 模式下自动执行 git add
```

**保存逻辑细节**：

**列操作顺序**：
```typescript
// 1. 先处理删除（从右向左，避免索引变化）
const deletes = colOps.filter(op => op.action === 'delete')
  .sort((a, b) => b.targetColNumber - a.targetColNumber);
for (const op of deletes) {
  ws.spliceColumns(op.targetColNumber, 1);
}

// 2. 再处理插入（从左向右，维护offset）
const inserts = colOps.filter(op => op.action === 'insert')
  .sort((a, b) => a.targetColNumber - b.targetColNumber);
let offset = 0;
for (const op of inserts) {
  ws.spliceColumns(op.targetColNumber + offset, 0, op.values);
  offset += 1;
}
```

**行操作顺序**：
```typescript
// 删除和插入混合处理，按 visualRowNumber 排序，维护offset
let offset = 0;
for (const op of sorted) {
  if (op.action === 'insert') {
    ws.spliceRows(op.targetRowNumber + offset, 0, op.values);
    offset += 1;
  } else if (op.action === 'delete') {
    ws.spliceRows(op.targetRowNumber + offset, 1);
    offset -= 1;
  }
}
```

## 🐛 常见问题和调试

### 1. 列对齐不正确

**症状**：两个文件明明有相同的列，但没有匹配上

**可能原因**：
- 表头文本格式不同（大小写、空格、标点）
- 数据类型不匹配（一个是数字，一个是字符串）
- 相似度阈值太严格

**调试方法**：
```typescript
// 在 buildAlignedColumns 中添加日志
console.log('Base columns:', baseCols.map(c => c.headerKey));
console.log('Side columns:', sideCols.map(c => c.headerKey));
console.log('Matches:', matched);
```

**解决方案**：
- 调整 `headerKey` 的标准化逻辑
- 调整相似度阈值（当前是 0.55）
- 使用更严格的 `headerKey` 匹配

### 2. 行对齐有歧义

**症状**：某些行被标记为 `ambiguous`

**原因**：
- 没有主键列，相似度匹配找到多个候选行
- 数据重复度高（例如很多空行）

**解决方案**：
- 指定主键列（如果数据有唯一标识）
- 调整行相似度阈值
- 手动选择正确的匹配

### 3. 合并预览不更新

**症状**：用户选择了 ours/theirs，但预览没有反映

**可能原因**：
- React 依赖项缺失
- 状态更新异步问题

**检查点**：
```typescript
// 确保 useEffect 依赖完整
useEffect(() => {
  // 构建预览...
}, [
  currentColOps,    // ✓ 列操作
  currentRowOps,    // ✓ 行操作
  mergeCells,       // ✓ 单元格选择
  mergeColumnsMeta, // ✓ 列元信息
]);
```

### 4. 保存后 Git 仍提示冲突

**症状**：保存成功但 `git status` 仍显示冲突

**原因**：
- `git add` 失败（可能 git 不在 PATH）
- 保存到了错误的文件

**检查点**：
```bash
# 手动执行
git add <merged-file>
git status
```

## 📊 性能优化

### 1. 虚拟滚动 (VirtualGrid.tsx)

**问题**：大型 Excel 文件（上万行）渲染卡顿

**解决方案**：
- 只渲染可视区域的行（± overscan）
- 用户滚动时动态更新渲染范围
- 避免一次性渲染所有 DOM 节点

```typescript
const visibleRowStart = Math.floor(scrollTop / rowHeight);
const visibleRowEnd = Math.ceil((scrollTop + viewportHeight) / rowHeight);
const renderStart = Math.max(0, visibleRowStart - overscanRows);
const renderEnd = Math.min(totalRows, visibleRowEnd + overscanRows);
```

### 2. 工作簿缓存 (`workbookCache`)

**问题**：频繁读取同一文件性能差

**解决方案**：
- 内存中缓存已加载的工作簿
- 避免重复读取磁盘
- LRU 淘汰策略（最多缓存10个）

```typescript
const workbookCache = new Map<string, Workbook>();
const loadWorkbookCached = async (filePath: string) => {
  if (workbookCache.has(filePath)) {
    return workbookCache.get(filePath)!;
  }
  const wb = new Workbook();
  await wb.xlsx.readFile(filePath);
  workbookCache.set(filePath, wb);
  return wb;
};
```

### 3. 稀疏单元格存储

**问题**：存储整个表格的矩阵占用内存大

**解决方案**：
- 只存储有差异的单元格（`MergeCell[]`）
- 而不是存储 `cells[row][col]` 的二维数组
- 大幅减少内存占用（差异通常 < 10%）

## 🔐 安全性考虑

### 1. 路径验证

```typescript
// 确保所有文件路径都是绝对路径
const normalizeCliPath = (p: string) => {
  const raw = stripOuterQuotes(p);
  return path.isAbsolute(raw) ? raw : path.resolve(process.cwd(), raw);
};
```

### 2. IPC 安全

```typescript
// 使用 contextIsolation 隔离
webPreferences: {
  contextIsolation: true,
  nodeIntegration: false,
}

// 只暴露必要的 API
contextBridge.exposeInMainWorld('excelAPI', {
  openFile: () => ipcRenderer.invoke('excel:open'),
  // ...
});
```

### 3. 文件写入确认

```typescript
// 保存前弹出确认对话框（交互模式）
const result = await dialog.showSaveDialog({
  title: '保存合并后的 Excel',
  defaultPath: templatePath,
});
if (result.canceled) return;
```

## 🚀 未来优化方向

1. **并行处理**：多个工作表并行 diff
2. **增量更新**：只重新计算变化的部分
3. **撤销/重做**：支持 undo/redo 操作
4. **格式保留**：更好地保留单元格格式（颜色、字体等）
5. **冲突标记**：在文件中插入 Git 风格的冲突标记
6. **测试覆盖**：添加单元测试和集成测试

## 📖 参考资料

- [ExcelJS 文档](https://github.com/exceljs/exceljs)
- [Electron 文档](https://www.electronjs.org/docs)
- [React 文档](https://react.dev/)
- [Git Mergetool](https://git-scm.com/docs/git-mergetool)
