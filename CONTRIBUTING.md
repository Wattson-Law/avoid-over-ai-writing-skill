# Contributing

感谢你愿意改进 `avoid-over-ai-writing`。

## 提交规则

提交新规则、模型适配或格式修复时，请尽量同时提供：

- 原始输入和处理目标；
- 不应改变的事实、数字和结论；
- 修改前后的最小示例；
- 对应的验证命令或复现步骤。

## 文风规则

新增文风规则应说明适用文种和触发条件，避免把一次模型输出概括为永久模型人格。规则只处理表达、结构和格式，不凭空补充业务事实。

## 提交前检查

```powershell
$env:PYTHONUTF8 = '1'
python C:\Users\asus\.codex\skills\.system\skill-creator\scripts\quick_validate.py .
python scripts\test_markdown_input.py
```

Pull Request 可以从一个小而可复现的改动开始。谢谢你的贡献。
