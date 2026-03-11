# WebStorm 调试配置问题修复指南

## 问题诊断

如果看到配置旁边有红色 ❌ 或问号 ❓，可能是以下原因：

1. **路径无法验证** - WebStorm 无法找到 electron 可执行文件
2. **配置格式问题** - XML 配置格式不正确
3. **需要重新加载** - WebStorm 需要重新识别配置

## 解决方案

### 方法一：手动在 WebStorm 中创建配置（推荐）

这是最可靠的方法：

#### 1. 创建主进程调试配置

1. 点击运行配置下拉菜单 → **Edit Configurations...**
2. 点击左上角的 **+** 号 → 选择 **Node.js**
3. 配置如下：
   - **Name**: `Electron Main Debug`
   - **Node interpreter**: 选择项目中的 Node.js（通常是 `project`）
   - **Node parameters**: 留空
   - **Working directory**: `$PROJECT_DIR$`
   - **JavaScript file**: `$PROJECT_DIR$/node_modules/electron/cli.js`
   - **Application parameters**: `.`
   - **Environment variables**: 点击 **...** 添加：
     - `NODE_ENV` = `development`
4. 在 **Debugger** 标签页：
   - 确保 **Port** 是 `9229`
   - **Mode** 选择 `Attach to Node.js/Chrome`
5. 点击 **OK** 保存

#### 2. 创建渲染进程调试配置

1. 点击 **+** 号 → 选择 **JavaScript Debug**
2. 配置如下：
   - **Name**: `Electron Renderer Debug`
   - **URL**: `http://localhost:3000`
3. 点击 **OK** 保存

### 方法二：使用 npm 脚本配置（更简单）

#### 创建主进程调试配置

1. 点击 **+** 号 → 选择 **npm**
2. 配置如下：
   - **Name**: `Electron Debug`
   - **Command**: `run`
   - **Scripts**: `dev`（如果已创建）或手动输入：
     - 先运行：`build:main`
     - 然后运行：`cross-env NODE_ENV=development electron .`
3. 在 **Before launch** 部分，添加：
   - **Run npm script** → `build:main`
4. 点击 **OK** 保存

### 方法三：修复现有配置

如果配置已存在但显示错误：

1. **检查路径**：
   - 打开配置编辑对话框
   - 确认 `node_modules/electron/cli.js` 路径正确
   - 如果路径显示为红色，点击路径旁边的文件夹图标重新选择

2. **重新加载项目**：
   - File → Invalidate Caches / Restart...
   - 选择 **Invalidate and Restart**

3. **检查 Node.js 解释器**：
   - File → Settings → Languages & Frameworks → Node.js
   - 确保 Node.js 解释器已正确配置

## 使用步骤

### 调试主进程

1. **先编译代码**（重要！）：
   ```bash
   npm run build:main
   ```

2. 在 `src/main/main.ts` 中设置断点

3. 在 WebStorm 中选择 **Electron Main Debug** 配置

4. 点击调试按钮 🐛 启动

### 调试渲染进程

1. **启动 webpack dev server**（在终端）：
   ```bash
   npm run dev:renderer
   ```

2. **编译主进程**（在另一个终端）：
   ```bash
   npm run build:main
   ```

3. **启动 Electron**（在第三个终端）：
   ```bash
   cross-env NODE_ENV=development electron .
   ```

4. 在 WebStorm 中选择 **Electron Renderer Debug** 配置

5. 点击调试按钮 🐛 启动

6. 或者直接在 Chrome DevTools 中调试（Electron 窗口会自动打开 DevTools）

## 常见问题

### Q: 配置显示红色 ❌

**A**: 
- 检查 `node_modules/electron/cli.js` 是否存在
- 运行 `npm install` 确保依赖已安装
- 在 WebStorm 中：File → Invalidate Caches / Restart

### Q: 断点无法命中

**A**:
- 确保已启用 source maps（已完成）
- 确保在开发模式下运行（`NODE_ENV=development`）
- 检查 source map 文件是否存在：
  - `build/main/main.js.map`
  - `dist/bundle.js.map`
- 确保断点设置在源文件（`src/`）中，不是编译后的文件

### Q: 端口 9229 被占用

**A**:
- 修改调试配置中的端口号（例如改为 9230）
- 或者关闭占用该端口的其他进程

## 验证配置

运行以下命令验证环境：

```bash
# 检查 electron 是否存在
Test-Path node_modules\electron\cli.js

# 检查编译后的文件
Test-Path build\main\main.js

# 检查 source map
Test-Path build\main\main.js.map
```

如果所有文件都存在，配置应该可以正常工作。
