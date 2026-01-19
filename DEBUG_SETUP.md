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
