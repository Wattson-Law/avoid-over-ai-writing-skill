---
name: avoid-over-ai-writing
language: en
description: Reduce formulaic AI prose in papers, technical proposals, patent disclosures, and formal Word documents. Normalize expression, heading numbering, fonts, paragraphs, monochrome tables, pagination, and tables of contents while preserving user facts, evidence, metrics, conclusions, and responsibility boundaries. Does not verify facts or perform technical review.
---

# Avoid Over AI Writing (English)

This skill has two coordinated paths: rewrite model output into precise, reader-focused language, then format the locked content as a deliverable Word document. It supports `.docx`, `.md`/`.markdown`, and pasted Markdown.

## Non-negotiable content boundaries

- Improve wording, paragraph organization, logic presentation, and formatting only.
- Do not verify, correct, or flag factual conflicts, logical contradictions, or suspicious metrics.
- Preserve every viewpoint in the source, including background judgments, problems, causes, goals, measures, benefits, limitations, recommendations, and implementation intent.
- Preserve all numbers, units, dates, prices, budgets, ratios, thresholds, direction words, named entities, products, conclusions, priorities, and responsibility assignments exactly.
- Do not add policies, examples, parameters, metrics, conclusions, or placeholders that are absent from the input.
- Reference assets provide format and structure only; never copy their project names, people, amounts, parameters, or conclusions.
- Use white-background, black-text tables with consistent 0.5 pt black grid lines. Bold a genuine business header row.
- Add standalone table labels in order: `Table 1`, `Table 2`, `Table 3`, and so on. Do not invent descriptive titles.
- Deliver only the requested normalized document. Keep internal audits and mapping evidence out of the body.

Read [content-boundaries.md](references/content-boundaries.md) before making content changes.

## Remove formulaic AI style

Rewrite hedging, boilerplate, stacked qualifiers, and unsupported promotional claims into sentences with a clear subject, action, condition, and result. Keep the user’s position and evidence. Do not use words such as “comprehensive,” “complete,” or “industry-leading” as substitutes for evidence.

Read [writing-style.md](references/writing-style.md) and [logic-structures.md](references/logic-structures.md) for `editorial` and `restructure` processing.

## Processing modes and strength

Choose one mode from `requirements-list`, `decision-proposal`, `rd-application`, and `rd-implementation-outline` based on the document’s structure and purpose. Use `format-only` only when explicitly requested; otherwise use conservative `editorial` rewriting. Use `restructure` only when the user explicitly requests deep restructuring.

Read the matching mode reference and [execution-and-acceptance.md](references/execution-and-acceptance.md) before conversion. For decision proposals or documents with an existing TOC, also read [toc-generation.md](references/toc-generation.md).

## Workflow

1. Inspect the available DOCX editing and rendering tools.
2. Build an immutable source baseline and inspect headings, paragraphs, tables, images, fields, sections, page breaks, revisions, and comments.
3. Lock the mode and processing strength; read the required references.
4. Build an internal content-preservation map, then process a new copy without overwriting the input.
5. Normalize heading levels, table labels, styles, pagination, and TOC fields.
6. Run content-preservation and standardized-format validation. Fix failures and rerun.
7. Update Word fields, render every page, and inspect for empty pages, broken tables, clipping, overlap, incorrect TOC entries, or orphaned headings.
8. Re-run audits after every layout-affecting fix. Deliver only after all checks pass.

The detailed scripts and acceptance commands are documented in the Chinese source skill and the repository references. This English file is the human-readable English entry point; the executable skill contract remains `SKILL.md`.
