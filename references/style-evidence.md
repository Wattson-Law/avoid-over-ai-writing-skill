# 文风规则的来源与使用边界

本文件记录 `avoid-over-ai-writing` 的规则来源。来源用于建立写作结构、证据呈现和格式检查的参照，不作为业务事实或模型答案的替代品。

## 来源分层

| 来源层 | 代表来源 | 用途 | 使用边界 |
| --- | --- | --- | --- |
| 中文规范 | GB/T 7713.2-2022《学术论文编写规则》、GB/T 7714-2015《信息与文献 参考文献著录规则》 | 论文结构、章节组织、参考文献和编号意识 | 仅采用规范要求的结构原则；不复制标准示例中的业务内容 |
| 英文学术写作规范 | APA Style、MLA Style Center、The Chicago Manual of Style | 清晰、简洁、引文和读者导向的表达 | 不把某一种学科的格式要求强加给所有文种 |
| 有证据的学术文本 | `allenai/qasper`（论文与证据支持的问答） | 训练评估维度：主张是否能回到证据，回答是否越过材料边界 | 数据集用于评估设计；不把单个答案当作写作模板 |
| 专利文本 | `NortheasternUniversity/big_patent` | 观察技术描述、发明摘要和权利范围的表达关系 | 不替用户新增权利要求、参数或法律结论 |
| 工程文本 | `princeton-nlp/SWE-bench_Verified` | 观察 issue、约束、修改动作和交付结果之间的对应关系 | 不把补丁或测试结果当作用户项目事实 |
| 模型评测 | Codex、Gemini、Claude 的相同提示词、多轮样本和盲评记录 | 识别叠甲、无支持的扩展、过度强调局限和直接交付失败 | 结论只适用于记录的模型版本和入口 |

## 规则如何落地

从上述来源提取的是可复核的写作动作：

- 把主张、依据、方法和结果分开写，减少一个长句承载多个逻辑层次；
- 保留数字、主体、限定词和结论，不用修辞替代证据；
- 用标题层级和编号呈现结构，避免重复包装和跳级；
- 把表格、引文、目录和页面样式作为交付结构单独验收；
- 对不同模型只添加经过重复评测的窄适配器，未达到稳定性门槛的观察保留为未决记录。

## 复现要求

每条模型适配记录应至少包含：模型标签、产品入口、日期、完整提示词、完整响应、样本编号、盲评结果、被拒绝的规则和最终补丁。评测协议要求每种文种使用三个来源案例、每个案例两次新运行，并由两名独立评审对同一输出投票。详细门槛见 [evaluation.md](evaluation.md)。

## 公开参考入口

- APA Style：<https://apastyle.apa.org/style-grammar-guidelines>
- MLA Style Center：<https://style.mla.org/>
- The Chicago Manual of Style：<https://www.chicagomanualofstyle.org/home.html>
- 国家标准全文公开系统：<https://openstd.samr.gov.cn/>
- Qasper：<https://huggingface.co/datasets/allenai/qasper>
- BIGPATENT：<https://huggingface.co/datasets/NortheasternUniversity/big_patent>
- SWE-bench Verified：<https://huggingface.co/datasets/princeton-nlp/SWE-bench_Verified>
