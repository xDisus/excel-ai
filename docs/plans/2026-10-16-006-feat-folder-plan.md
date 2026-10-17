---
title: "feat: Convert a whole folder"
type: feat
date: 2026-10-16
artifact_contract: ce-unified-plan/v1
status: completed
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: lightweight
---

# feat: Convert a whole folder

## Summary

`python excel_ai.py pasta/` converts every `.xlsx`/`.xlsm` in the folder, one `.ai.txt` per workbook next to it, and keeps going when one file fails. Same output as running the tool on each file.

## Requirements

- R1. If the single argument is a directory, convert each `.xlsx`/`.xlsm` directly inside it (not subfolders), sorted by name. Skip Excel lock files (`~$*`).
- R2. Per file, reuse `main(caminho)` unchanged (same output file, same console report), preceded by a `== name ==` line.
- R3. A file that fails does not stop the others. Exit code: 0 all ok, 1 if any self-check failed, 2 if any file could not be read or the folder has no workbooks.
- R4. Last line summarizes: `Pasta: N planilha(s), K com problema`.
- R5. Single-file, `--celula` and `--diff` usage unchanged.

## Key Technical Decisions

- KTD1. Single ASCII file, openpyxl only, Python 3.8 (session-settled: user-approved — chosen over multi-module package / extra deps: user copies one .py by hand at work).
- KTD2. Public repo, no real workbook data or names in tests (session-settled: user-directed — chosen over committing the real calculator as fixture: work files confidential).
- KTD3. No network (session-settled: user-approved — chosen over any remote call: financial data stays local).
- KTD4. Lossless self-check stays (session-settled: user-approved — chosen over dropping verification: AI must trust the text). Each file keeps its own check via `main()`.
- KTD5. Loop over `main()`, no new converter path, no combined index file. Rejected: a merged folder dump (bigger output, no asked need).

## Implementation Units

### U1. `main_pasta` and CLI

**Goal:** folder argument works.

**Requirements:** R1-R5

**Dependencies:** none

**Files:** `excel_ai.py`, `README.md`

**Approach:** `main_pasta(pasta)` lists, loops `main()`, collects codes; `__main__` routes a directory argument to it; usage text and README line.

**Test scenarios:**
- Folder with two valid workbooks, a `~$` lock file, a `.txt` and a subfolder with a workbook: two `.ai.txt` created, others ignored.
- A corrupt `.xlsx` among valid ones: the others are still converted; exit 2; summary says 1 with problem.
- Empty folder: exit 2.
- A file path argument still behaves as before.

**Verification:** self-test passes.

## Scope Boundaries

### Deferred to Follow-Up Work

- Recursing into subfolders; a combined index; parallel conversion.
