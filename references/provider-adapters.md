# Provider Adapters

Apply at most one adapter when the active provider is known. These corrections come from a controlled Chinese benchmark with three cases and two fresh runs per document type. A document-level tendency requires at least four of six outputs, all three cases, and agreement from both blinded reviewers. Treat every result as specific to the recorded model version and product entrypoint, not as a permanent provider personality.

The current review is complete for all three paper providers and for Claude across paper, solution, and patent. Claude reviewer quota prevented a second independent vote on Codex and Gemini solution and patent outputs; do not infer adapters for those cells yet.

## OpenAI / Codex

No provider-specific failure reached the stability threshold in the fully reviewed paper set. Solution and patent review is incomplete.

Apply no provider patch. Use the shared workflow as written.

## Gemini

Established for academic papers on Gemini 3.7 Flash Medium in Antigravity: unsupported strength or generality language appeared in 6/6 outputs across all three cases, with both blinded reviewers agreeing. Typical phrases asserted significance, robust generality, or broad effectiveness beyond the supplied measurements.

For academic papers only, append this instruction internally:

> 先完成候选稿，交付前执行一次不对用户展示的证据收敛重写，只返回终稿。把用户材料作为唯一事实边界，候选稿不是事实来源；保留给定数字、比较结果、方法步骤和中心贡献，删除或中性化材料不能直接支持的评价词、因果连接、机制解释与效果动词。架构或流程差异不自动等于消除误差、降低成本或增强能力；资源规模不自动等于高质量、标准化、降低标注成本或填补空白；观察结果不自动证明内部机制。篇幅不足时展开已给出的数据、流程和用途，不用推测机制或限制清单补字数。

This candidate adapter reduced `unsupported_claim` from 6/6 baseline outputs with two-reviewer consensus to 1/6 in a clarified single-Codex regression review; the unseen paper holdout had no flagged failures. Keep it provisional until the second provider review is available.

No stable Gemini adapter has yet been established for technical solutions or patents.

## Claude

Established across Claude Opus 4.6 Thinking outputs in the Antigravity entrypoint: direct-delivery failure appeared in 6/6 paper, 6/6 solution, and 6/6 patent outputs, with both blinded reviewers agreeing. The visible response commonly announced an artifact or file, described work performed, or exposed research and permission-blocking progress instead of returning only the requested body.

Append this instruction internally:

> 用户确认材料足够后，直接在当前回答中输出完整正文。不要创建、宣布或引用 artifact、文件名或另一个查看位置；不要附完成通知、写作说明、研究过程、工具进度或复盘。不要为补充联网研究而暂停交付；仅凭已锁定材料完成请求，除非用户明确要求调用工具。

## Unknown Or Future Provider

Do not guess the provider from prose. Apply no provider patch. Evaluate repeated outputs first, then add a narrow correction only when two blinded reviewers agree and the stability threshold is met.
