# WebStorm Electron 调试配置详细指南

## 配置界面参数说明

根据你看到的配置界面，需要填写以下参数：

### 1. **File:** 字段（最重要！）

**填写内容：**
```
$PROJECT_DIR$/node_modules/electron/cli.js
```

**或者手动选择：**
- 点击 `File:` 字段旁边的文件夹图标 📁
- 导航到：`node_modules/electron/cli.js`
- 选择该文件

**说明：** 这是 Electron 的启动入口文件，必须填写！

---

### 2. **Application parameters:** 字段

**填写内容：**
```
.
```

**说明：** 这个点号 `.` 表示当前目录，告诉 Electron 从当前目录启动（会读取 package.json 中的 main 字段）

---

### 3. **Working directory:** 字段

**应该显示：**
```
E:\electron_excel
```

**如果为空或错误：**
- 点击文件夹图标 📁
- 选择项目根目录：`E:\electron_excel`

---

### 4. **Node interpreter:** 字段

**应该显示：**
```
node (C:\Program Files\nodejs\node.exe) 22.17.1
```

**如果显示错误：**
- 点击下拉箭头或浏览按钮
- 选择正确的 Node.js 解释器路径

---

### 5. **Environment variables:** 字段

**应该显示：**
```
NODE_ENV=development
```

**如果为空：**
- 点击编辑图标（铅笔图标）
- 点击 **+** 添加新变量
- Name: `NODE_ENV`
- Value: `development`
- 点击 **OK**

---

### 6. **Before launch** 部分的问题修复

如果看到 `Unknown Task MODE` 和 `Unknown Task DEBUG_PORT`：

**解决方法：**
1. 选中这些 Unknown Task
2. 点击 **-** 号删除它们
3. 这些是调试器内部配置，不需要手动添加

**正确的 Before launch 应该包含：**
- 点击 **+** 号
- 选择 **Run npm script**
- Script: `build:main`
- 这样在调试前会自动编译主进程代码

---

## 完整配置步骤

### 步骤 1：填写基本参数

1. **File:** `$PROJECT_DIR$/node_modules/electron/cli.js`
2. **Application parameters:** `.`
3. **Working directory:** `E:\electron_excel`（或 `$PROJECT_DIR$`）
4. **Node interpreter:** 选择你的 Node.js（22.17.1）
5. **Environment variables:** 添加 `NODE_ENV=development`

### 步骤 2：配置 Before launch

1. 在 **Before launch** 部分，点击 **+** 号
2. 选择 **Run npm script**
3. Script: `build:main`
4. 点击 **OK**

这样配置后，每次调试前会自动编译主进程代码。

### 步骤 3：配置调试器

1. 点击 **Debugger** 选项卡（在 Configuration 旁边）
2. 确保：
   - **Port:** `9229`
   - **Mode:** `Attach to Node.js/Chrome` 或 `Listen for incoming connections`

### 步骤 4：保存并测试

1. 点击 **OK** 保存配置
2. 在 `src/main/main.ts` 中设置一个断点
3. 点击调试按钮 🐛 启动
4. 如果一切正常，断点应该会被命中

---

## 配置后的界面应该显示

✅ **File:** `node_modules/electron/cli.js`  
✅ **Application parameters:** `.`  
✅ **Working directory:** `E:\electron_excel`  
✅ **Environment variables:** `NODE_ENV=development`  
✅ **Before launch:** `Run npm script 'build:main'`  

---

## 如果配置后仍然无法调试

### 检查清单：

1. ✅ **确保已编译代码：**
   ```bash
   npm run build:main
   ```

2. ✅ **检查文件是否存在：**
   - `node_modules/electron/cli.js` ✓
   - `build/main/main.js` ✓
   - `build/main/main.js.map` ✓（source map）

3. ✅ **验证 Electron 路径：**
   在终端运行：
   ```bash
   node node_modules/electron/cli.js .
   ```
   如果 Electron 窗口能打开，说明路径正确。

4. ✅ **检查端口是否被占用：**
   ```bash
   netstat -ano | findstr :9229
   ```
   如果被占用，修改调试配置中的端口号。

5. ✅ **重新加载 WebStorm：**
   - File → Invalidate Caches / Restart...
   - 选择 **Invalidate and Restart**

---

## 快速配置模板

如果手动配置太麻烦，可以直接复制以下配置到 WebStorm：

**File:** `$PROJECT_DIR$/node_modules/electron/cli.js`  
**Application parameters:** `.`  
**Working directory:** `$PROJECT_DIR$`  
**Node interpreter:** `project`（使用项目配置的 Node.js）  
**Environment variables:** `NODE_ENV=development`  
**Before launch:** `Run npm script 'build:main'`  

保存后应该就可以正常调试了！
