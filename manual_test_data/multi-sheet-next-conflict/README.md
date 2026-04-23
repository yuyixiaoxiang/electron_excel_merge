# Multi-Sheet Merge Test Data

用于手工验证 merge 模式下“下一个冲突”是否会跨 sheet 跳转。

文件：

- `base.xlsx`
- `ours.xlsx`
- `theirs.xlsx`
- `merged-output.xlsx`

工作表与冲突点：

- `achievementGroup`
  - `D5`
    - `base`: `10001|10002|10003`
    - `ours`: `10001|10002|90001`
    - `theirs`: `10001|10002|80001`
- `achievement`
  - `D6`
    - `base`: `完成10次战斗`
    - `ours`: `完成10次胜利战斗`
    - `theirs`: `完成10次竞技战斗`
  - `E8`
    - `base`: `gem:10`
    - `ours`: `gem:12`
    - `theirs`: `gem:15`

建议验证：

1. 先在 `achievementGroup` 处理完 `D5`。
2. 再点“下一个冲突”。
3. 应自动切到 `achievement`，并定位到 `D6`。
4. 再点一次，应跳到 `E8`。
