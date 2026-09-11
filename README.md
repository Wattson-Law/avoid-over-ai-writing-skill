# 避免过度 AI 写作

一个可复用的 Codex skill，用来把模型生成的论文、技术方案、专利披露和正式 Word 文稿整理成更确定、可读、可交付的版本。

它解决两类经常同时出现的问题：

- 模型喜欢叠加“可能、通常、需要进一步”等防御性表达，或用“全面、显著、行业领先”等没有证据的宣传词；
- 正文内容基本可用，但标题层级、编号、字体、表格、分页和目录不适合直接交付。

skill 会先锁定用户的立场、受支持的事实和输出要求，再做保留式改写；随后可把 Word、Markdown 或粘贴文本转换为统一格式的 DOCX。它强调呈现方式，不能代替事实核验、同行评审、专利检索或技术验收。

## 使用

将本目录安装为 Codex skill 后，在请求中使用 `$avoid-over-ai-writing`，并说明文种、读者、已确定的结论以及期望格式。需要 Word 输出时，提供 `.docx`、`.md` 或直接粘贴 Markdown；skill 会另存新文件并运行内容守恒、格式和分页检查。

标题编号默认使用 `一、`、`二、`、`六、` 等单一形式。不会把同一编号包装成 `第六章`，也不会生成 `第六、` 或 `六章、` 这样的混合形式。

Markdown 转换器默认会把 `第六章 标题` 规范为 `六、标题`；确需保留原称谓时，使用 `--preserve-chapter-wrapper`。

## 项目结构

- `SKILL.md`：路由、证据边界、文风适配和 Word 工作流。
- `references/`：论文、方案、专利、模型适配器、格式与验收规则。
- `scripts/`：Markdown→DOCX、目录、内容守恒和格式校验工具。
- `vendor/`：转换脚本运行所需的精简 Python 依赖。

公开包不包含任何公司模板、业务样例或内部文档。

## 本地校验

```powershell
$env:PYTHONUTF8='1'
python C:\Users\asus\.codex\skills\.system\skill-creator\scripts\quick_validate.py .
python scripts\test_markdown_input.py
```

MIT License.
