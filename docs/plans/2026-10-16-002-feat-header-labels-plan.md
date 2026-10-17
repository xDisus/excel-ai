---
title: "feat: Header labels on formula blocks"
type: feat
date: 2026-10-16
artifact_contract: ce-unified-plan/v1
status: completed
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: lightweight
---

# feat: Header labels on formula blocks

## Summary

Each formula block line gains an optional `[label]` taken from the nearest text header, e.g. `E2:E361 [Juros] =C2*G2`, so the AI reads meaning, not just cell math. The self-check parser accepts and ignores the label, so round-trip stays lossless.

## Problem Frame

Today `E2:E361 =C2*G2` tells the AI how a column is computed but not what it is. The meaning lives in a text header cell (often in the Data section far away). Putting the header beside the block lets the AI map formulas to business concepts (Juros, Saldo, PDD) without cross-referencing.

## Requirements

- R1. Block lines from `blocos()` show `[label]` between range and formula when a text header is found.
- R2. Label source: vertical run (same column, many rows) looks above the first row; horizontal run looks left of the first column. A single cell looks left; it looks above only when the left walk finds no non-empty cell at all. The walk skips at most 2 empty cells; farther away means no label.
- R3. The first non-empty cell hit ends the walk. It is a label only if it is constant text (not formula, number or date) containing at least 2 letters (so `R$`, `%`, `-` are not labels).
- R4. Labels are sanitized: whitespace collapsed, `[`/`]` removed, cut to 30 chars. Empty after sanitizing means no label.
- R5. `expandir()` parses lines with or without the label; the self-check still proves 100% round-trip.
- R6. File header explains the label in one line.
- R7. On the real calculator: 100% round-trip, token increase under 10%.

## Key Technical Decisions

- KTD1. Single ASCII file, openpyxl only, Python 3.8 (session-settled: user-approved — chosen over multi-module package / extra deps: user copies one .py by hand at work).
- KTD2. Public repo, no real workbook data or names in tests (session-settled: user-directed — chosen over committing the real calculator as fixture: work files confidential).
- KTD3. No network (session-settled: user-approved — chosen over any remote call: financial data stays local).
- KTD4. Lossless self-check stays (session-settled: user-approved — chosen over dropping verification: AI must trust the text). Labels are hints outside the formula, so the check compares formulas only.
- KTD5. Syntax `RANGE [label] =formula`: formula always starts with `=` and sanitized labels never contain `]`, so `RE_BLOCO` gets one optional non-capturing group. Rejected: trailing `# label` comment (could be confused with formula text).
- KTD6. Labels only on block lines, not on `except`, array, or `same formulas as` diff lines. Diff lines call `secao_formulas` with empty values, so they naturally get no labels.

## Implementation Units

### U1. Label lookup and rendering

**Goal:** `blocos()` adds `[label]` to block lines.

**Requirements:** R1, R2, R3, R4, KTD5, KTD6

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
1. Add a small helper that, given `formulas`, `valores`, a start cell and a direction, walks per R2/R3 and returns a sanitized label or empty.
2. A position in `formulas` (any type) counts as non-empty and stops the walk with no label; `blocos()` already receives both dicts, no signature change.
3. In `blocos()` pick the direction from the block shape (R2) and insert the label in the block line.

**Patterns to follow:** existing `fmt()` / `ref()` helpers; terse Portuguese names and comments.

**Test scenarios:** (in `autoteste()`, synthetic fixture only)
- Column header "Juros" above `E2:E5` formulas gives `E2:E5 [Juros] =...`.
- Row label "Total" left of a horizontal run gives `[Total]`.
- Numeric cell above a column block gives no label.
- Formula cell directly above gives no label (walk stops).
- Header 2 empty cells above still labels; 3 empty cells above gives no label.
- Unit cell `%` left of a horizontal run gives no label.
- Single cell with a number to its left gives no label (no fall-through); single cell with nothing to its left but text above gets the above label.
- Header with `]` and a long text is sanitized and cut to 30 chars.

**Verification:** self-test passes; labels visible in fixture output.

### U2. Parser, header text, real-file check

**Goal:** round-trip accepts labels; docs updated.

**Requirements:** R5, R6, R7, KTD4

**Dependencies:** U1

**Files:** `excel_ai.py`, `README.md`

**Approach:**
1. Extend `RE_BLOCO` with an optional `(?:\[[^\]]*\] )?` before the formula.
2. Add one header line in `converter()` explaining the label.
3. One README bullet.

**Test scenarios:**
- `expandir()` on a labeled line returns the same cells as the unlabeled line.
- Full fixture conversion still reports 0 self-check failures.

**Verification:** real calculator (local only, never committed): falhas 0, tokens grow less than 10% vs the 31,743 baseline. Spot-check about 20 labeled blocks by eye; most labels should name the column/row correctly. If growth is 10% or more, restrict labels to blocks spanning 2+ cells and re-measure.

## Scope Boundaries

### Deferred to Follow-Up Work

- Naming referenced columns inside the formula (`C=Saldo`): bigger output; revisit after measuring labels.
- Labels from Excel table column names (`ws.tables`) — the map already lists them.
- Multi-row headers (two header lines stacked).

## Assumptions

- Nearest text above/left is the header in most financial calculators; wrong guesses are harmless hints because the formula itself is unchanged.
