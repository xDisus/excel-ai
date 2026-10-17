---
title: "feat: Number-format hints and cell comments - Plan"
type: feat
date: 2026-10-17
artifact_contract: ce-unified-plan/v1
artifact_readiness: requirements-only
product_contract_source: ce-brainstorm
---

# feat: Number-format hints and cell comments - Plan

## Goal Capsule

- **Objective:** the AI misreads calculation logic because the text drops two meaning clues the workbook already holds: number format (rate vs money vs date) and author comments. Put both in the output.
- **Product authority:** user, in the brainstorm of 2026-10-17.
- **Open blockers:** none.

## Product Contract

### Problem

Today `excel_ai.py` ignores number formats and cell comments (verified: no `number_format` or comment handling in the script). A rate shown as `0,00%` and a value in `R$` look the same in the text. Authors often explain assumptions in comments. The concern is the AI explaining the calculation logic wrongly. Not yet confirmed that format/comments are the cause (assumption: validate on one real workbook before building).

### Requirements

- R1. Formula and data lines carry a short type hint taken from the cell's number format: percent, currency, date. General or unknown formats add nothing.
- R2. Cell comments/notes appear in the output next to their cell, text verbatim (long ones truncated with a marker).
- R3. Hints are facts read from the file, never inferred. No per-sheet "what it does" summary.
- R4. Output stays lossless for formulas (self-check unchanged) and does not grow noticeably on sheets without these features.
- R5. Existing modes (`--celula`, `--diff`, folder) keep working; `--diff` does not report a format-only change as a formula change.
- R6. The legend the AI reads (embedded `FORMAT_MD` and `SKILL_MD`) documents the new hints and comment lines, otherwise the AI cannot interpret them.

### Out of scope

- Per-sheet purpose summaries (guessing; revisit only if still needed).
- Reading styles beyond number format (colors, borders, fonts).

### Acceptance examples

- AE1. A column formatted `0.00%` shows a `%` hint on its formula block; a `R$` column shows currency.
- AE2. A cell with the comment "taxa anual" shows that text next to the cell in the output.
- AE3. A workbook with no custom formats and no comments produces the same output as today.

### Key Decisions

- Chosen over a per-sheet summary: only facts from the workbook, no guessing (user-approved).
- Single ASCII file, openpyxl only, no network (session-settled: carried from earlier plans).

### Outstanding Questions

Resolved by the author's judgment (user: "siga com o seu melhor julgamento"):

- Mixed formats inside a formula block: formats are listed in their own `## Formats` section by range, independent of formula blocks. No block splitting, no parser change.
- R4 growth: only sheets that have percent/date/currency cells or comments gain lines; ranges keep it compact (cap 50 ranges per type, 100 notes, 200 chars each). Measure on a real workbook before relying on it.
- `--diff` ignores comment and format changes.

Still open: whether this actually helps the AI (validate on one real workbook).
