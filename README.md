# electron_excel

## 开发环境
- Windows
- Node.js + npm
- WebStorm（2024.2.x）

## 安装依赖
```bash path=null start=null
npm install
```

## 开发启动（不调试）
```bash path=null start=null
npm run dev
```
- `dev:renderer` 会启动 webpack-dev-server，默认端口 `http://localhost:3000`
- `dev:main` 会以 watch 模式编译主进程 TS 到 `build/`

## WebStorm 调试 Electron 主进程（main process）
本项目主进程入口为 TypeScript（`src/main/main.ts`），开发模式下主进程会加载 `http://localhost:3000`。

由于在 Windows 上直接用 IDE 的 Debug 按钮“启动 Electron”可能不稳定，推荐使用：
- **Run 启动 Electron（带 --inspect）**
- **再使用 Attach 连接调试端口**

### 1. 前置：确保 WebStorm 插件可用
在 WebStorm 中打开：`File | Settings | Plugins`，确保启用：
- Node.js
- JavaScript Debugger（或与 JS 调试相关的插件）

启用后重启 WebStorm。

### 2. 创建/检查 Run 配置：npm dev（启动 3000 + tsc watch）
`Run | Edit Configurations...` → `+` → **npm**
- **package.json**：`E:\electron_excel\package.json`
- **Command**：`run`
- **Scripts**：`dev`

启动：右上角选择 `npm: dev` → **Run**。

### 3. 创建/检查 Run 配置：Electron Main (inspect run)
`Run | Edit Configurations...` → `+` → **Node.js**
- **Node interpreter**：`E:\electron_excel\node_modules\electron\dist\electron.exe`
- **Working directory**：`E:\electron_excel`
- **Application parameters**：`.`
- **Node parameters**：`--inspect=9229`
- **Environment variables**：`NODE_ENV=development`

启动：选择 `Electron Main (inspect run)` → **Run**。

> 注意：
> - `.` 必须放在 **Application parameters**，不要放到 Node parameters。
> - `NODE_ENV=development` 会让主进程加载 `http://localhost:3000`。

### 4. 创建/检查 Debug 配置：Attach 9229
`Run | Edit Configurations...` → `+` → **Attach to Node.js/Chrome**（名称可能略有差异）
- **Host**：`127.0.0.1`
- **Port**：`9229`

启动：选择 `Attach 9229` → **Debug**。

### 5. 推荐启动顺序（稳定）
1. **Run**：`npm: dev`
2. **Run**：`Electron Main (inspect run)`
3. **Debug**：`Attach 9229`

此时在 `src/main/main.ts` 中下断点即可命中。

## 调试渲染进程（renderer process）
渲染进程是 React + Webpack 的页面，开发模式下由 webpack-dev-server 提供：`http://localhost:3000`。

### 方式 A：用 Electron 内置 DevTools 调试（推荐）
1. 按上面的流程启动 `npm: dev` + `Electron Main (inspect run)`。
2. 在 Electron 窗口打开 DevTools：`Ctrl + Shift + I`。
3. 打开 DevTools → `Sources`：
   - 在左侧 `webpack://`（或类似项）中找到 `src/renderer/*.tsx`。
   - 直接在 `.tsx` 源码行号处打断点即可。
4. 也可以在代码里临时加入：
```js path=null start=null
debugger;
```
然后触发对应操作（点击/滚动等），会自动断住。

> 如果只能看到 `bundle.js`，看不到 `src/renderer/*.tsx`：通常是 source map 没开。
> 本项目开发模式建议保持 webpack 的 `devtool` 为 `source-map` / `eval-source-map` / `cheap-module-source-map` 之一。

### 方式 B：用 IDE 断点调试渲染进程（可选）
大多数情况下直接用 DevTools 就够了；如果你希望在 IDE 里调试 TSX：
- 优先建议在 DevTools 里断点（方式 A）。
- 如需 IDE 调试，通常要启用 Electron 的远程调试端口（例如启动参数加 `--remote-debugging-port=9222`），
  再用 IDE 的 Chrome/JS 调试器 attach 到该端口。

## 同时调试 main + renderer（推荐组合）
1. `npm: dev`（启动 webpack-dev-server + tsc watch）
2. `Electron Main (inspect run)`（启动 Electron 主进程）
3. `Attach 9229`（IDE 调试 main）
4. Electron 窗口 `Ctrl+Shift+I`（DevTools 调试 renderer）

## 常见问题
### 1) 3000 端口被占用
webpack-dev-server 默认使用 3000。可用以下命令查看占用进程：
```powershell path=null start=null
netstat -ano | findstr ":3000"
```
然后用 PID 查进程：
```powershell path=null start=null
tasklist /FI "PID eq <PID>"
```

### 2) Electron 白屏
开发模式下白屏通常是 `http://localhost:3000` 未启动或渲染进程报错。
- 确认 `npm: dev` 正在运行且 `http://localhost:3000` 可访问
- 在 Electron 窗口打开 DevTools（Ctrl+Shift+I）查看 Console 报错

## Excel Diff 原理（Myers + 行相似度）
Excel diff 通常采用“先粗对齐，再细匹配”的思路：
1. **Myers 行级 diff（粗对齐）**：把每一行当作一个整体，快速找出完全相同的行，以及新增/删除的变动块。
2. **行相似度匹配（细对齐）**：在变动块内对“删除行”和“新增行”计算相似度（按单元格内容/相同列比例等），相似度高的行会被识别为“修改行”，而不是“删了又加”。

这样既保留了 Myers 的速度与全局最短编辑优势，又能把改动行对齐到具体单元格，提升 diff 可读性。
