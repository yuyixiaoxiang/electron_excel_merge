# WebStorm 断点调试配置说明

## 已完成的配置

1. ✅ **TypeScript Source Maps**: 已在 `tsconfig.json` 中启用
2. ✅ **Webpack Source Maps**: 已在 `webpack.config.js` 中启用 `devtool: 'source-map'`
3. ✅ **WebStorm 调试配置**: 已创建主进程和渲染进程的调试配置

## 使用方法

### 方法一：使用 WebStorm 调试配置（推荐）

#### 调试主进程（Main Process）

1. 首先确保已编译主进程代码：
   ```bash
   npm run build:main
   ```

2. 在 WebStorm 中：
   - 打开 **Run/Debug Configurations**（运行/调试配置）
   - 选择 **Electron Main Debug**
   - 在 `src/main/main.ts` 中设置断点
   - 点击调试按钮（🐛）启动

#### 调试渲染进程（Renderer Process）

1. 首先启动 webpack dev server（在终端运行）：
   ```bash
   npm run dev:renderer
   ```

2. 在 WebStorm 中：
   - 打开 **Run/Debug Configurations**
   - 选择 **Electron Renderer Debug**
   - 在 `src/renderer/` 中的任何文件设置断点
   - 点击调试按钮启动
   - 然后手动启动 Electron（在另一个终端运行）：
     ```bash
     npm run build:main
     cross-env NODE_ENV=development electron .
     ```

### 方法二：使用 Chrome DevTools（渲染进程）

1. 启动开发模式：
   ```bash
   npm run dev:renderer
   ```

2. 在 `src/main/main.ts` 中添加以下代码以打开 DevTools：
   ```typescript
   if (isDev) {
     mainWindow.webContents.openDevTools();
   }
   ```

3. 在 Chrome DevTools 中设置断点并调试

### 方法三：同时调试主进程和渲染进程

1. **步骤 1**: 启动 webpack dev server
   ```bash
   npm run dev:renderer
   ```

2. **步骤 2**: 在 WebStorm 中启动 **Electron Main Debug** 配置
   - 这会启动 Electron 并附加调试器到主进程

3. **步骤 3**: 在渲染进程代码中设置断点后，使用 Chrome DevTools
   - 在 Electron 窗口中按 `Ctrl+Shift+I` 打开 DevTools
   - 在 Sources 标签页中找到你的源文件并设置断点

## 注意事项

1. **Source Maps**: 确保在开发模式下 source maps 已生成
2. **端口冲突**: 如果 9229 端口被占用，修改调试配置中的端口号
3. **文件路径**: 确保断点设置在源文件（`src/`）中，而不是编译后的文件（`build/`）

## 故障排除

如果断点无法工作：

1. 检查 source maps 是否生成：
   - `build/main/main.js.map`（主进程）
   - `dist/bundle.js.map`（渲染进程）

2. 确保在开发模式下运行（`NODE_ENV=development`）

3. 清除缓存并重新构建：
   ```bash
   npm run clean
   npm run build:main
   ```

4. 检查 WebStorm 的调试器设置：
   - File → Settings → Build, Execution, Deployment → Debugger
   - 确保 "JavaScript" 和 "Node.js" 调试器已启用

## WebStorm 调试 Electron 主进程（main process）
本项目主进程入口为 TypeScript（`src/main/main.ts`），开发模式下主进程会加载 `http://localhost:3000`。

由于在 Windows 上直接用 IDE 的 Debug 按钮“启动 Electron”可能不稳定，推荐使用：
- Run 启动 Electron（带 `--inspect`）
- 再使用 Attach 连接调试端口

### 1. 前置：确保 WebStorm 插件可用
在 WebStorm 中打开：`File | Settings | Plugins`，确保启用：
- Node.js
- JavaScript Debugger（或与 JS 调试相关的插件）

启用后重启 WebStorm。

### 2. 创建/检查 Run 配置：npm dev（启动 3000 + tsc watch）
`Run | Edit Configurations...` -> `+` -> npm
- package.json：`E:\electron_excel\package.json`
- Command：`run`
- Scripts：`dev`

启动：右上角选择 `npm: dev` -> Run。

### 3. 创建/检查 Run 配置：Electron Main (inspect run)
`Run | Edit Configurations...` -> `+` -> Node.js
- Node interpreter：`E:\electron_excel\node_modules\electron\dist\electron.exe`
- Working directory：`E:\electron_excel`
- Application parameters：`.`
- Node parameters：`--inspect=9229`
- Environment variables：`NODE_ENV=development`

启动：选择 `Electron Main (inspect run)` -> Run。

注意：
- `.` 必须放在 Application parameters，不要放到 Node parameters。
- `NODE_ENV=development` 会让主进程加载 `http://localhost:3000`。

### 4. 创建/检查 Debug 配置：Attach 9229
`Run | Edit Configurations...` -> `+` -> Attach to Node.js/Chrome（名称可能略有差异）
- Host：`127.0.0.1`
- Port：`9229`

启动：选择 `Attach 9229` -> Debug。

### 5. 推荐启动顺序（稳定）
1. Run：`npm: dev`
2. Run：`Electron Main (inspect run)`
3. Debug：`Attach 9229`

此时在 `src/main/main.ts` 中下断点即可命中。

## 调试渲染进程（renderer process）
渲染进程是 React + Webpack 的页面，开发模式下由 webpack-dev-server 提供：`http://localhost:3000`。

### 方式 A：用 Electron 内置 DevTools 调试（推荐）
1. 按上面的流程启动 `npm: dev` + `Electron Main (inspect run)`。
2. 在 Electron 窗口打开 DevTools：`Ctrl + Shift + I`。
3. 打开 DevTools -> Sources：
   - 在左侧 `webpack://`（或类似项）中找到 `src/renderer/*.tsx`。
   - 直接在 `.tsx` 源码行号处打断点即可。
4. 也可以在代码里临时加入：
   ```js
   debugger;
   ```
   然后触发对应操作（点击/滚动等），会自动断住。

如果只能看到 `bundle.js`，看不到 `src/renderer/*.tsx`：通常是 source map 没开。
本项目开发模式建议保持 webpack 的 `devtool` 为 `source-map` / `eval-source-map` / `cheap-module-source-map` 之一。

### 方式 B：用 IDE 断点调试渲染进程（可选）
大多数情况下直接用 DevTools 就够了；如果你希望在 IDE 里调试 TSX：
- 优先建议在 DevTools 里断点（方式 A）。
- 如需 IDE 调试，通常要启用 Electron 的远程调试端口（例如启动参数加 `--remote-debugging-port=9222`），再用 IDE 的 Chrome/JS 调试器 attach 到该端口。

## 同时调试 main + renderer（推荐组合）
1. `npm: dev`（启动 webpack-dev-server + tsc watch）
2. `Electron Main (inspect run)`（启动 Electron 主进程）
3. `Attach 9229`（IDE 调试 main）
4. Electron 窗口 `Ctrl+Shift+I`（DevTools 调试 renderer）

## 常见问题
### 1) 3000 端口被占用
webpack-dev-server 默认使用 3000。可用以下命令查看占用进程：
```powershell
netstat -ano | findstr ":3000"
```
然后用 PID 查进程：
```powershell
tasklist /FI "PID eq <PID>"
```

### 2) Electron 白屏
开发模式下白屏通常是 `http://localhost:3000` 未启动或渲染进程报错。
- 确认 `npm: dev` 正在运行且 `http://localhost:3000` 可访问
- 在 Electron 窗口打开 DevTools（Ctrl+Shift+I）查看 Console 报错
