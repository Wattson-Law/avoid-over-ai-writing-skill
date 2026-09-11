# Provider Adapters

Apply at most one adapter when the active provider is known. These corrections come from a controlled Chinese benchmark with three cases and two fresh runs per document type. A document-level tendency requires at least four of six outputs, all three cases, and agreement from both blinded reviewers. Treat every result as specific to the recorded model version and product entrypoint, not as a permanent provider personality.

The current review records model labels, entrypoints, prompts, complete responses, and blinded comparisons. Adapter text is kept narrow and tied to the recorded model version and entrypoint.

## OpenAI / Codex

Use the shared workflow as the baseline for Codex. The current records do not require an additional Codex-specific patch.

Apply no provider patch. Use the shared workflow as written.

## Gemini

Gemini 3.7 Flash Medium in the Antigravity entrypoint participated in six academic-paper runs across three source cases. The review used these samples to tune evidence-convergence wording and reduce unsupported strength or generality language.

For academic papers only, append this instruction internally:

> 先完成候选稿，交付前执行一次不对用户展示的证据收敛重写，只返回终稿。把用户材料作为唯一事实边界，候选稿不是事实来源；保留给定数字、比较结果、方法步骤和中心贡献，删除或中性化材料不能直接支持的评价词、因果连接、机制解释与效果动词。架构或流程差异不自动等于消除误差、降低成本或增强能力；资源规模不自动等于高质量、标准化、降低标注成本或填补空白；观察结果不自动证明内部机制。篇幅不足时展开已给出的数据、流程和用途，不用推测机制或限制清单补字数。

Keep this adapter tied to academic papers and the recorded Gemini entrypoint. Re-run the same source cases and an unseen holdout when the model version changes.

No stable Gemini adapter has yet been established for technical solutions or patents.

## Claude

For Claude Opus 4.6 Thinking in the Antigravity entrypoint, use a direct-delivery instruction so the final response stays focused on the requested body.

Append this instruction internally:

> 用户确认材料足够后，直接在当前回答中输出完整正文。不要创建、宣布或引用 artifact、文件名或另一个查看位置；不要附完成通知、写作说明、研究过程、工具进度或复盘。不要为补充联网研究而暂停交付；仅凭已锁定材料完成请求，除非用户明确要求调用工具。

## Unknown Or Future Provider

Do not guess the provider from prose. Apply no provider patch. Evaluate repeated outputs first, then add a narrow correction only when two blinded reviewers agree and the stability threshold is met.
