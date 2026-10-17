---
title: "feat: Cached value beside formula blocks"
type: feat
date: 2026-10-16
artifact_contract: ce-unified-plan/v1
status: completed
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: lightweight
---

# feat: Cached value beside formula blocks

## Summary

Each formula block line gains an optional `(v X)` hint with the value Excel last calculated for the block's first cell, e.g. `E2:E361 [Juros] (v 12.5) =C2*G2`. The AI sees results without recalculating. The self-check parser accepts and ignores the hint.

## Problem Frame

Formulas alone do not say what they produced. Today the cached value is used only for error warnings and the `--celula` header. Showing it per block lets the AI sanity-check logic (e.g. "this IRR cell gives 0.012") with a few extra tokens.

## Requirements

- R1. A block line from `blocos()` shows `(v X)` between the optional `[label]` and the formula when the first cell has a cached value.
- R2. Only the first cell of a block spanning 2+ cells gets a value (single cells measured +10.2% tokens on the real file, so they are left out; final +3.5%). `except` lines and array/diff lines get none (same scope rule as labels).
- R3. X comes from `fmt()` of the cached value, with `)` removed, whitespace collapsed, cut to 20 chars. Missing cache, `None` or empty text means no hint.
- R4. `expandir()` parses lines with or without the hint (`RE_BLOCO` gets one optional non-capturing group), so the round-trip stays 100% lossless.
- R5. Works with `--celula`: the cache passed to `blocos()` is the unfiltered sheet cache.
- R6. File header explains the hint in one line.
- R7. On the real calculator (local only): round-trip 100%, token growth vs the post-labels baseline (32,601) under 10%. If it is 10% or more, restrict hints to blocks spanning 2+ cells and re-measure.

## Key Technical Decisions

- KTD1. Single ASCII file, openpyxl only, Python 3.8 (session-settled: user-approved — chosen over multi-module package / extra deps: user copies one .py by hand at work).
- KTD2. Public repo, no real workbook data or names in tests (session-settled: user-directed — chosen over committing the real calculator as fixture: work files confidential).
- KTD3. No network (session-settled: user-approved — chosen over any remote call: financial data stays local).
- KTD4. Lossless self-check stays (session-settled: user-approved — chosen over dropping verification: AI must trust the text). The hint sits before the `=`, so formulas are untouched and the check compares formulas only.
- KTD5. Syntax `RANGE [label] (v X) =formula`. Sanitized X never contains `)`, and the formula always starts with `=`. Rejected: trailing `-> X` (would be read as part of the formula text).
- KTD6. First cell only, not every cell: keeps growth small; the Data section and per-cell dump are not touched.

## Implementation Units

### U1. Hint rendering and parsing

**Goal:** `blocos()` adds `(v X)`; `RE_BLOCO` accepts it.

**Requirements:** R1, R2, R3, R4, R5, KTD5, KTD6

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
1. Helper `valor_curto(v)` returns the sanitized text or "".
2. `blocos()` and `secao_formulas()` take an optional `cache` dict; `converter()` passes the sheet cache (unfiltered in chain mode).
3. Insert the hint after the label in the block line only.
4. Extend `RE_BLOCO` with `(?:\(v [^)]*\) )?` after the label group.

**Patterns to follow:** `rotulo()` and the `[label]` insertion in `blocos()`.

**Test scenarios:** (synthetic fixture in `autoteste()`)
- Block with cached number gives `E2:E5 [Juros] (v 12.5) =...`.
- Cached string with `)` and newline is sanitized; 30-char text is cut to 20.
- No cache, or value `None`, gives no hint.
- `expandir()` on a line with label+hint equals the unlabeled line; a tampered hint-line that breaks the formula fails `confere`.
- `except` and `array` lines have no hint.
- `--celula` output shows hints for chain blocks.
- Error value (`#N/A`) is shown as is.

**Verification:** self-test passes; hints visible in fixture output.

### U2. Header, README, real-file check

**Goal:** docs and measurement.

**Requirements:** R6, R7

**Dependencies:** U1

**Files:** `excel_ai.py`, `README.md`

**Approach:** one header line in `converter()`, one README bullet, then run the local real file (never committed) for tokens and round-trip.

**Test scenarios:** Test expectation: none -- docs plus local measurement.

**Verification:** real file: falhas 0, growth under 10%.

## Scope Boundaries

### Deferred to Follow-Up Work

- Values for every cell of a block, or last-cell value.
- Hints in the Data section.

## Assumptions

- First-cell value is the most useful single sample; it is a hint, wrong guesses are harmless.
