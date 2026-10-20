# Slide Markdown contract

The compiler supports a small, deliberate Markdown subset. UTF-8 and CRLF/LF are supported. Every heading of the form `## Clear title {layout}` creates exactly one slide. There is no automatic cover or closing; include them in the source while respecting the requested total count.

Markdown compiles to self-contained, editable presentation HTML before PPTX. For review-first, output `.html` and stop; for direct delivery, output `.pptx` and the generator retains the intermediate HTML. See [HTML conversion](html-conversion.md) for resuming from reviewed HTML.

Optional single-line frontmatter keys: `language` (`en` or `id`, default `en`), `title`, `subtitle`, `client`, `date`, `author`. Title/subtitle metadata describe the source; visible slide copy comes from slide headings and bodies. Unknown keys, multiline YAML, arrays, and nested YAML are rejected.

Do not use manual slide numbering, rich Markdown, code blocks, arbitrary HTML, nested lists, or Markdown links in visible copy. The compiler uses plain text; it does not execute source content. Put factual sources and URLs in speaker notes.

## Component syntax

| Layout | Body syntax |
|---|---|
| cover | Plain supporting text lines |
| section | One plain supporting line |
| ledger | 1–5 `- Plain item` rows, or 1–3 `- Label \| Description \| Status` rows |
| kpi | 3–4 `- Value \| Label \| Context` rows |
| proof | Six `- Value \| Label \| Context` rows |
| cards | 2–3 `- Title \| Description \| Footer label` rows |
| flow / process | 2–4 `- Label \| Description` rows |
| timeline | 3–4 `- Date \| Milestone \| Description` rows |
| columns | Two `- Label \| Description` rows |
| table | Markdown pipe table; 2–4 columns, 2–6 rows including header |
| image | One `![Meaningful alt text](relative.png)` line, then one plain caption |
| quote | One plain statement line, then one attribution line |
| closing | Secondary headline line, then consultation subtitle line |

The backslashes in this reference table escape the separator for Markdown display. Do not write backslashes around the actual ` | ` field separator in a slide source. The compiler rejects the wrong number of fields instead of discarding them. Escaped pipes and multiline table cells are unsupported. Only local PNG artwork is supported; paths resolve relative to the source Markdown. Images are embedded proportionally.

## Runnable fragment

```markdown
---
language: en
title: Pilot delivery plan
client: Illustrative client
author: Aptaworks Team
date: 20 October 2026
---

## A clear delivery flow {flow}
- Diagnose | Confirm the scope and sources.
- Build | Implement the pilot workflow.
- Review | Validate the result with its owner.
<!-- notes: This is an illustrative delivery example, not an actual client commitment. -->

## Milestones across the pilot {timeline}
- WEEK 01 | Discover | Agree the scope and acceptance criteria.
- WEEK 02 | Build | Create and review the pilot.
- WEEK 03 | Handover | Transfer the runbook and ownership.

## Maximize your operational efficiency and business growth {closing}
with Aptaworks' cutting-edge data solutions.
Get a consultation with our experts at
```

Use [examples/example.md](../examples/example.md) for all supported components. Optional notes use `<!-- notes: Sources, credits, or presenter guidance. -->`; multiple blocks are combined. Do not put QA logs in presenter notes.
