---
title: "feat: Diff between two workbook versions (--diff)"
type: feat
date: 2026-10-16
artifact_contract: ce-unified-plan/v1
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: standard
---

# feat: Diff between two workbook versions (`--diff`)

## Summary

`python excel_ai.py antiga.xlsx nova.xlsx --diff` writes a text with only what changed between two versions: sheets added or removed, formulas new or changed (shown in the compact block format, with the old formula beside), formulas removed, and constants changed. The AI answers "what changed in v7.4 vs v7.3" without reading both full dumps.

## Problem Frame

The user keeps versioned calculators (v7.3, v7.4...). Two full dumps cost about 2 x 33k tokens and the AI must diff them by eye. A focused diff is tiny and exact.

## Requirements

- R1. CLI `file_old file_new --diff`. Anything else keeps working (single file, `--celula`). A missing file or bad args prints an error and exits 2. `--diff` with `--celula` is not supported (usage error).
- R2. Sheets match by exact title. A sheet only in new is listed as added (with formula count); only in old as removed. Unchanged sheets are left out and counted in the header.
- R3. For a sheet in both, formulas are compared by the normalized text `esperado()` gives (so a pure Excel fill-down shift is not a change). Cells are:
  - new or changed: shown with `secao_formulas()` blocks (labels and `(v X)` values from the new file), under `## Formulas new or changed`;
  - changed or removed: the old formula shown under `## Formulas before (changed or removed cells)`, same block format.
- R4. A cell that is a formula in one version and a constant (or empty) in the other counts as changed or removed on the formula side and shows the constant under Constants.
- R5. Constants (non-formula cells) compared by value (`fmt()`): one line per cell `A1 old -> new`, `A1 added X`, `A1 removed (was X)`. Capped at 200 lines per sheet; the rest is summarized as `... N more`.
- R6. Self-check: each formula section's lines must expand (`confere`) to exactly the formulas they stand for; on failure the section falls back to one formula per cell (`por_celula`) and the count is reported, like the normal dump.
- R7. Defined names added or removed are listed (a changed destination shows as removed + added).
- R8. Output file `<new stem>.vs_<old stem>.diff.ai.txt` (ASCII-sanitized). Console prints token count of the diff, changed counts and the self-check result.
- R9. Real calculator (local only): diff the file against itself gives "no changes" with 100% check; against a locally edited copy (change a few cells) shows exactly those. Under 1 minute.

## Key Technical Decisions

- KTD1. Single ASCII file, openpyxl only, Python 3.8 (session-settled: user-approved — chosen over multi-module package / extra deps: user copies one .py by hand at work).
- KTD2. Public repo, no real workbook data or names in tests (session-settled: user-directed — chosen over committing the real calculator as fixture: work files confidential).
- KTD3. No network (session-settled: user-approved — chosen over any remote call: financial data stays local).
- KTD4. Lossless self-check stays for the formulas shown (session-settled: user-approved — chosen over dropping verification: AI must trust the text).
- KTD5. Extract the per-sheet reading loop of `converter()` into `ler_abas()` and reuse it for both files; add `converter_diff()` beside `converter()`. Rejected: a separate reader, which would duplicate normalization.
- KTD6. Compare `esperado()` strings, not raw formulas: they already strip internal prefixes and carry array ranges, so what the AI sees is what is compared.
- KTD7. CLI stays hand-parsed `sys.argv` (third shape). Rejected: `argparse`, same reason as `--celula`.

## Implementation Units

### U1. Reader refactor and diff engine

**Goal:** `ler_abas()` shared; `converter_diff(antiga, nova)` returns (text, numbers).

**Requirements:** R2, R3, R4, R5, R6, R7, KTD5, KTD6

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
1. Move the sheet-reading loop (cache, formulas, valores, `bruto_chars`, `sem_cache`) into `ler_abas(wb, wb_v)`; `converter()` calls it with unchanged output.
2. `converter_diff()` opens both files, reads sheets, and per sheet builds the new/changed, before and constants sections.
3. Self-check each formula section; fallback per R6.

**Patterns to follow:** `diferenca()` and the `confere`/`por_celula` fallback in `converter()`.

**Test scenarios:** (synthetic fixtures built in `autoteste()`)
- Identical workbooks: header says no changes, no sheet sections.
- A changed formula in a column run shows the new block and the old block; an untouched run is absent.
- A pure fill shift (same relative formula) is not a change.
- Added sheet and removed sheet are listed.
- Formula replaced by a constant: shows in "before" and as a constant line.
- Constant changed, added, removed lines; the 200-line cap with `... N more`.
- A defined name added/removed/changed destination.
- Array formula change is detected (range or text).
- A forced `confere` failure falls back to one line per cell and counts a failure.
- Normal single-file conversion output is unchanged (existing asserts pass).

**Verification:** self-test passes.

### U2. CLI, naming, docs

**Goal:** `--diff` entry point, console report, README.

**Requirements:** R1, R8, R9

**Dependencies:** U1

**Files:** `excel_ai.py`, `README.md`

**Approach:** `main_diff(antiga, nova)` mirrors `main()`; `__main__` accepts `a b --diff`; usage text and docstring updated; README section.

**Test scenarios:**
- Subprocess: `a b --diff` exits 0 and writes the expected file name.
- Missing file exits 2; `a b` without `--diff` and `a b --diff --celula x` give usage exit 2.

**Verification:** self-test passes; real-file check per R9.

## Scope Boundaries

### Deferred to Follow-Up Work

- Matching renamed sheets, moved rows or inserted columns (a row insert shows as many changes).
- Diffing more than two files or whole folders (improvement 5).
- Diffing formatting, charts, styles.

## Assumptions

- Sheets keep their names between versions. Row/column insertions produce noisy diffs; accepted for this version.
