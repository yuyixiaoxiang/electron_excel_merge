# Fork 安装与配置（eMerge）

本文档用于在 Fork 中配置 `eMerge` 作为外部合并/对比工具。

## 前置条件
1. 已打包或可运行 `eMerge.exe`。
2. 可执行文件路径可访问，例如：`E:\electron_excel\dist_electron\win-unpacked\eMerge.exe`。

## Merge Tools 配置
在 Fork 中打开 `Settings -> Integration -> External Merge Tools`，新增或编辑一项：

- Name: `eMerge`
- Path: `E:\electron_excel\dist_electron\win-unpacked\eMerge.exe`
- Arguments: `"$BASE" "$LOCAL" "$REMOTE" "$MERGED"`

参数说明：
- `"$BASE"`：共同基线文件
- `"$LOCAL"`：当前分支文件
- `"$REMOTE"`：目标分支文件
- `"$MERGED"`：输出合并结果文件

## Diff Tools 配置
在 Fork 中打开 `Settings -> Integration -> External Diff Tools`，新增或编辑一项：

- Name: `eMerge`
- Path: `E:\electron_excel\dist_electron\win-unpacked\eMerge.exe`
- Arguments: `"$LOCAL" "$REMOTE"`

参数说明：
- `"$LOCAL"`：当前版本文件
- `"$REMOTE"`：对比版本文件

## 在 Fork 中使用
1. 在文件列表中右键目标文件。
2. 点击 `External Diff`。
3. 选择 `eMerge`。

## 截图
把你提供的截图保存到以下路径后，文档会自动显示：

1. `docs/assets/fork/fork-external-diff-menu.png`
2. `docs/assets/fork/fork-integration-settings.png`

![Fork 外部对比菜单](./assets/fork/fork-external-diff-menu.png)
![Fork 集成设置](./assets/fork/fork-integration-settings.png)
