# 三方 Merge/Diff 算法说明

## 目标与输入
- 目标：对 `base / ours / theirs` 三个 Excel 表进行三方对比，仅比较**单元格值**，忽略格式。
- 输入：三份 Excel 文件路径、主键列（可选）、冻结行数、行相似度阈值。

## 值的规范化与读取
- 单元格值统一转为可比较的“简单值”：
  - 富文本取 `text` 合并；
  - 公式取 `result`；
  - 日期转为可读字符串；
  - 其他对象尽量转成字符串。
- 比较时会做字符串化与去空白（trim）。

## 工作表对齐策略
1. **同名工作表优先**：按 base 的工作表顺序，匹配三边同名表。
2. **索引兜底**：剩余未匹配的工作表按索引对齐（第 1 张对第 1 张…）。

## 列对齐策略
1. 构建 `ColumnRecord`：基于表头行（`frozenRowCount`）+ 少量正文采样，提取列名文本、类型分布、取值指纹。
2. 先用 **LCS/Myers** 对齐“表头完全相同”的列，作为锚点。
3. 对未匹配列，在锚点之间用“列相似度矩阵（列名/类型/内容）”进行匹配。
4. 生成对齐后的列序列，并记录每个对齐列的 `baseCol / oursCol / theirsCol` 映射。

## 行对齐总流程
1. 先构造 `RowRecord`（行号、非空列、行值、主键）。
2. **冻结行**（header）：
   - 前 `frozenRowCount` 行按**固定行号**逐行比较（不做对齐）。
3. **正文行**：
   - 如果可用主键：走 **按主键对齐**；
   - 否则：走 **序列对齐（Myers diff）**。

---

## 主键对齐（alignRowsByKey）
### 主键来源
- 若用户提供且合法（1-based列号）则使用；
- 否则自动检测隐式主键列：
  - 覆盖率 ≥ 0.8、唯一性 ≥ 0.9；
  - 选取 coverage * uniqueness 最高的列。

### 对齐规则
- 先按 key 分组；
- 对重复 key：
  - 若 base/side 数量一致 → 按出现次序对齐；
  - 数量不一致 → 从候选中用**行相似度**挑最佳；
    - 低于阈值或与第二名差距过小 → 判为**ambiguous**。

### “主键变更”兜底
- 若某 side 完全缺失该 key，尝试在**未匹配行**中用“忽略主键列的相似度”找最佳匹配。

---

## 序列对齐（alignRowsBySequence）
1. 将每行值拼接成 token（标准化后的值用 `||` 连接）。
2. 用 **Myers diff** 找 equal / insert / delete。
3. 对 delete/insert：
   - **优先匹配 token 完全相同**的行（避免重复行错配）；
   - 再用“相似度 + 窗口搜索”匹配；
   - 若多个候选得分接近 → 标记 **ambiguous**。
4. 未匹配的插入行按“gap”位置插入到结果里。

---

## 行相似度与行状态
- 行相似度 = 相同单元格数 / 参与比较的列数。
- 行状态：
  - `added` / `deleted` / `modified` / `unchanged` / `ambiguous`
- 结果会记录每个视觉行对应的 base/ours/theirs 原始行号及相似度。

---

## 单元格差异判定
对齐后的每行、每列做三方比较（只比较值）：

- `unchanged`：base=ours=theirs  
- `ours-changed`：ours≠base，theirs=base  
- `theirs-changed`：theirs≠base，ours=base  
- `both-changed-same`：ours=theirs≠base  
- `conflict`：三者不一致  

合并默认值：
- `unchanged` → base
- `ours-changed` → ours
- `theirs-changed` → theirs
- `both-changed-same` → ours
- `conflict` → ours（默认）

---

## 冻结行上下文补齐
- 如果存在差异列，会将这些列在冻结行中**补齐显示**（即使未变化），用于保持表头/冻结行上下文。

---

## hasExactDiff（避免对齐误判）
- 额外提供**坐标级**全表扫描：
  - 只要任一单元格三方值不一致，即判为 `hasExactDiff = true`。
- 用于避免因对齐错位造成的“误红点”。

---

## 输出结构
- `cells[]`：仅包含有差异的单元格（含状态与 mergedValue）。
- `rowsMeta[]`：视觉行信息（行号映射、相似度、状态）。
- `hasExactDiff`：工作表是否存在坐标级真实差异。

---

## merged 数据的存储与写入流程

### merged 值存储位置（渲染进程）
- 每个差异单元格 `MergeCell` 都带有 `mergedValue`，作为“最终合并值”。
- 该值会同时维护在两份状态里：
  - `mergeSheets[].cells[].mergedValue`（全量，跨所有工作表）
  - `mergeCells[].mergedValue`（当前工作表展示用副本）
- 当用户选择“用 base / ours / theirs”时，会同步更新这两份状态，确保 UI 与保存数据一致。

### 写入 merged 文件（主进程）
- 保存时由渲染进程汇总所有 `mergedValue`，生成 `{ sheetName, address, value }[]`。
- 通过 IPC 调用 `excel:saveMergeResult` 发送到主进程。
- 主进程处理流程：
  1. 选择目标路径：
     - CLI `mode=merge`：优先 `mergedPath`，否则回退写 `oursPath`
     - CLI `mode=diff`：写 `oursPath`
     - 交互模式：弹出保存对话框
  2. 用 `exceljs` 打开模板文件（默认 `oursPath` 作为格式模板）
  3. 逐个写入单元格 `value`（仅改值，不改样式/公式）
  4. `writeFile(targetPath)` 保存
  5. 若是 `merge` 模式并有目标文件，会尝试执行一次 `git add`

### 新增列的合并与写回
- 当某列仅存在于 `theirs` 时，可在列头右键选择“使用本新增列”。
- 该操作会：
  - 将该列所有差异单元格的 `mergedValue` 切换为 `theirs`；
  - 记录列插入操作，保存时在目标文件中插入该列并写入对齐后的列值。
