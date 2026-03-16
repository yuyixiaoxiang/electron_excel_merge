# electron_excel

## 欢迎界面
![冲欢迎界面1](docs/assets/welcome.png)
说明：开始界面，提示用户比较文件夹或者文件。

## 文件夹对比
![文件夹对比](docs/assets/folder.png)
说明：文件夹对比界面。

## Diff
![Diff 占位图1](docs/assets/diff.png)
说明：Diff 流程截图 1（在文件列表中触发 External Diff）。

## 冲突处理
![冲突处理占位图1](docs/assets/merge.png)
说明：冲突处理流程截图 1（进入冲突文件后的操作入口）。



## 开发环境
- Windows
- Node.js + npm
- WebStorm（2024.2.x）

## 安装依赖
```bash path=null start=null
npm install
```

## 打包发布（EXE：压缩 / 非压缩）
本项目使用 `electron-builder`，输出目录为 `dist_electron/`。

### 压缩版（发布用）
压缩版通常体积更小，适合发给最终用户安装或直接运行。

1. `NSIS 安装包（.exe）`
```bash path=null start=null
npm run pack:nsis
```

2. `Portable 单文件（.exe）`
```bash path=null start=null
npm run pack:win
```

3. 一次构建两种发布包（NSIS + Portable）
```bash path=null start=null
npm run dist
```

### 非压缩版（调试/排查用）
非压缩目录版不打安装包，便于查看和定位打包产物内容。

```bash path=null start=null
npm run pack:unpacked
```

### 产物位置说明
- 安装包（压缩）：`dist_electron/*.exe`
- Portable（压缩）：`dist_electron/*.exe`（执行 `npm run pack:win` 生成）
- Unpacked（非压缩）：`dist_electron/win-unpacked/`

### Git 提交约定（打包产物）
- 需要提交：`dist_electron/*.exe`（发布给他人的安装包/便携包）。
- 不提交：`dist_electron/win-unpacked/` 以及临时打包目录（用于本地调试）。
- 建议流程：每次执行打包命令后，只挑选新的 `.exe` 进入提交。

### 体积对比（记录模板）
可在每次发布前后记录一次，便于持续优化包体积。

| 构建方式 | 命令 | 典型产物 | 体积（MB） | 适用场景 |
|---|---|---|---:|---|
| NSIS 安装包（压缩） | `npm run pack:nsis` | `dist_electron/*.exe` | 待填写 | 正式安装发布 |
| Portable 单文件（压缩） | `npm run pack:win` | `dist_electron/*.exe` | 待填写 | 免安装分发、U盘携带 |
| Unpacked 目录（非压缩） | `npm run pack:unpacked` | `dist_electron/win-unpacked/` | 待填写 | 调试排查、文件核查 |

可用 PowerShell 快速查看体积：
```powershell path=null start=null
Get-ChildItem dist_electron -File *.exe | Select-Object Name,@{N='SizeMB';E={[math]::Round($_.Length/1MB,2)}}
```

### 发布策略（建议）
1. 对外正式发布：优先 `NSIS`，同时提供 `Portable` 作为免安装备选。
2. 内部测试联调：优先 `Portable`，安装负担小、回归快。
3. 定位打包问题：使用 `Unpacked`，可直接检查 `win-unpacked` 内文件是否完整。
4. 每次发版至少保留一份 `Unpacked` 构建结果，便于问题复现和比对。

### 打包前检查清单
1. 执行 `npm install`，依赖完整且无中断。
2. 执行 `npm run build` 能成功通过（主进程 + 渲染进程）。
3. 图标文件存在：`build/icon.ico`。
4. 额外资源存在：`resources/portable-git`（打包后应进入 `resources/git`）。
5. 清理历史产物：`dist_electron/` 内旧文件不用于本次发版。

### 打包后验证清单
1. 启动应用无白屏，主流程可用（打开文件、diff/merge、导出等关键路径）。
2. 校验版本与名称（窗口标题、可执行文件名）符合发布预期。
3. 校验资源目录：`resources/git` 已随包携带。
4. 在未安装开发环境的干净机器上进行一次冒烟测试（推荐）。

## 开发启动（不调试）
```bash path=null start=null
npm run dev
```
- `dev:renderer` 会启动 webpack-dev-server，默认端口 `http://localhost:3000`
- `dev:main` 会以 watch 模式编译主进程 TS 到 `build/`

## 调试与常见问题
调试流程与常见问题已迁移到单独文档：
- `docs/DEBUG_SETUP.md`

## Fork 配置（Merge / Diff Tools）
用于在 Fork 中配置 `eMerge` 作为外部合并/对比工具。

### 前置条件
1. 已打包或可运行 `eMerge.exe`。
2. 可执行文件路径可访问，例如：`E:\electron_excel\dist_electron\win-unpacked\eMerge.exe`。

### Merge Tools 配置
在 Fork 中打开 `Settings -> Integration -> External Merge Tools`，新增或编辑一项：

- Name: `eMerge`
- Path: `E:\electron_excel\dist_electron\win-unpacked\eMerge.exe`
- Arguments: `"$BASE" "$LOCAL" "$REMOTE" "$MERGED"`

参数说明：
- `"$BASE"`：共同基线文件
- `"$LOCAL"`：当前分支文件
- `"$REMOTE"`：目标分支文件
- `"$MERGED"`：输出合并结果文件

### Diff Tools 配置
在 Fork 中打开 `Settings -> Integration -> External Diff Tools`，新增或编辑一项：

- Name: `eMerge`
- Path: `E:\electron_excel\dist_electron\win-unpacked\eMerge.exe`
- Arguments: `"$LOCAL" "$REMOTE"`

参数说明：
- `"$LOCAL"`：当前版本文件
- `"$REMOTE"`：对比版本文件

### 在 Fork 中使用
1. 在文件列表中右键目标文件。
2. 点击 `External Diff`。
3. 选择 `eMerge`。

### 配置截图
![Fork 集成设置](docs/assets/fork-integration-settings.png)





完整独立文档：`docs/FORK_SETUP.md`

## Excel Diff 原理（jsdiff/Myers + 行相似度 + 三方分类）
Excel diff / merge 采用“先粗对齐，再细匹配，最后做三方分类”的思路：
1. **jsdiff 行级 diff（粗对齐）**：使用 `diff`（jsdiff）里的 Myers 序列算法，把每一行当作一个整体，先找出确定相同的行，以及新增/删除的变动块。
2. **行相似度匹配（细对齐）**：在变动块内对“删除行”和“新增行”计算相似度（按单元格内容/相同列比例等），相似度高的行会被识别为“修改行”，而不是“删了又加”。
3. **三方单元格分类**：在行已经对齐之后，再基于 `base / ours / theirs` 把单元格区分成 `ours-changed`、`theirs-changed`、`both-changed-same`、`conflict`，用于驱动 UI 颜色和默认 merged 值。

这样既保留了 Myers 的速度与全局最短编辑优势，又能把改动行对齐到具体单元格，并让三方 merge 的默认行为更接近真实语义。
