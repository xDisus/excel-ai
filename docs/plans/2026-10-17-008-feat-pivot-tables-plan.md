---
title: "feat: Read pivot tables"
type: feat
date: 2026-10-17
artifact_contract: ce-unified-plan/v1
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: standard
---

# feat: Read pivot tables

## Summary

Add a `## Pivots` section to the `.ai.txt`: for each pivot table, where it sits, what range or table feeds it, which fields are rows, columns, filters and values, and what each value computes. Today pivots are ignored, so the AI sees only their cached numbers with no hint they are a summary. First of five plans (pivots, connections, Power Query, external links, macros); this one has no dependency on the others.

## Problem Frame

A pivot's cells are plain values in the sheet. The AI cannot tell that `Resumo!A3:D20` is "sum of Valor by Mes" fed from `Dados!A1:F900`. It may explain those numbers as hardcoded inputs. openpyxl keeps pivot definitions on load (`ws._pivots`, each with `.cache`), so the facts are readable without a new dependency.

## Requirements

- R1. Each pivot yields one block: name, sheet, location range, source (worksheet range, or table name, or "external/other" when the cache source is not a worksheet).
- R2. List fields by role: rows, columns, filters (pages), and values with their function (sum, count, average...) and caption.
- R3. Facts only, read from the file. No inferred meaning. Unknown or unreadable parts say so in one word, never crash.
- R4. Warn when the pivot refreshes on open or its cache is missing, because its numbers may differ from what the file shows.
- R5. Workbooks without pivots produce byte-identical output to today.
- R6. The legend the AI reads (`FORMAT_MD` and `SKILL_MD`) documents the `## Pivots` section.
- R7. `--celula`, `--diff` and folder modes keep working. `--diff` ignores pivot changes. Under `--celula` (chain mode, the `continue` near `excel_ai.py:706`) pivots and all new workbook-level sections are omitted.
- R8. Self-check (lossless formulas) unchanged. Single file, openpyxl only, Python 3.8, no network.

## Scope Boundaries

- Not rendering pivot results (the cell values already appear in the data sections).
- Not calculated fields/items formulas beyond showing their names and formula text if openpyxl exposes it. Slicers and timelines: ignored.

### Deferred to Follow-Up Work

- Connections, Power Query, external links, macros: plans 009 to 012.
- `--diff` awareness of pivots.

## Key Technical Decisions

- KTD1. Single ASCII file, openpyxl only, Python 3.8 (session-settled: user-directed — chosen over a multi-module package: user copies one .py by hand at work).
- KTD2. Public repo, no real workbook data in tests (session-settled: user-directed — chosen over a real fixture: work files confidential).
- KTD3. No network (session-settled: user-approved).
- KTD4. Read pivots through openpyxl objects (`ws._pivots`), not raw XML. Chosen over parsing `xl/pivotTables/*.xml` by hand: less code; the underscore attribute is semi-private, so wrap access in `getattr(ws, "_pivots", [])` and degrade to no section on any error.
- KTD5. Pure formatter function takes a pivot object and returns lines, so tests build `TableDefinition` objects in memory instead of crafting a zip.
- KTD6. New section placed after `## Notes`, one block per pivot, field names resolved from `cache.cacheFields` by index. Matches the existing section style (ASCII, one fact per line).

- KTD7. Shared foundations for plans 008 to 012, landed here first (U0) so the five plans do not conflict. Review of the five plans found placement, redaction and untrusted-input gaps; the fixes live in these helpers.
- KTD8. Section placement. Pivots are per sheet, inside the sheet block after `## Notes`. Connections, Queries, External links and Macros are workbook-level, emitted once in a single `# Workbook` block after the last sheet, in fixed order: Connections, Queries, External links, Macros. Each plan appends to a shared `secoes_livro` list instead of editing the join. A heading appears only when its section has content.
- KTD9. One redaction policy for everything new: `redigir(texto)` keeps only the file name of any path (Windows, UNC, `/home`, `file:///`, with `%20` decoded), strips URL query, fragment and userinfo, masks emails and string literals after credential keys (`password`, `pwd`, `key`, `token`, `secret`, `authorization`, `bearer`). Redact first, then truncate. Free text (M code, VBA) is best-effort; the legend tells the user to review before sharing.
- KTD10. Untrusted-input guard: `ler_membro(zf, nome, maximo)` checks `ZipInfo.file_size` before reading and refuses anything over the cap; XML containing `<!DOCTYPE` or `<!ENTITY` is refused before `ElementTree` sees it. Default caps: 5 MB per zip member, 10 MB for decoded blobs.

## Implementation Units

### U0. Shared foundations

**Goal:** helpers that plans 009 to 012 reuse.

**Requirements:** R3, R8

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
- Add `redigir`, `ler_membro`, `xml_seguro` (DOCTYPE/ENTITY refusal then parse), `_injetar(buf, {path: bytes})` (copies every entry of a fixture zip and adds parts, for tests) and the `secoes_livro` list plus the `# Workbook` block emitter.
- Placement: next to `tem_vba` (`excel_ai.py:404`).

**Test scenarios:**
- `redigir("C:\\Users\\ana\\Docs\\taxas.xlsx")` returns `taxas.xlsx`; a UNC path and `file:///x/y%20z.xlsx` likewise.
- `redigir("https://h/p?api_key=abc#f")` has no `abc`, no query, no fragment; userinfo `user:pw@h` removed.
- `redigir('Password = "hunter2"')` hides `hunter2`; an email is masked.
- `ler_membro` refuses a member whose declared size exceeds the cap; `xml_seguro` refuses a DOCTYPE with entities.
- `_injetar` output loads in openpyxl and contains the added part.
- Empty `secoes_livro` emits no `# Workbook` block (output unchanged).

**Verification:** auto-test green; no plan below needs its own redaction or size-guard code.

### U1. Pivot reader and formatter

**Goal:** turn `ws._pivots` into text lines.

**Requirements:** R1, R2, R3, R4

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
- Add `secao_pivots(ws)` next to `secao_extras` (`excel_ai.py:650`), same naming style (Portuguese helper names, English output).
- Resolve field index to name through the cache's field list; guard index out of range and the special "values" pseudo-field index (-2) in row/column lists.
- Source: `cache.cacheSource.worksheetSource` gives sheet and ref, or `name` printed as "table or name" (openpyxl cannot tell them apart); a non-worksheet `cacheSource.type` or a `connectionId` prints the type word (link to plan 009 deferred).
- Data fields: `fld` indexes into `cacheFields`; show name, subtotal function, caption.
- `refreshOnLoad` is on the cache object, not the pivot. A pivot with no cache is only reachable on an in-memory object: on a real load openpyxl raises KeyError on the missing cache before our code runs (out of scope).

**Patterns to follow:** `secao_extras` (`excel_ai.py:650`), table lines in the map (`excel_ai.py:751`).

**Test scenarios:**
- Happy path: pivot with 1 row field, 1 column field, 1 filter and 2 values (sum, count) prints all four roles with function names.
- Source variants: worksheet range source prints `Sheet!A1:F900`; table-name source prints the table.
- Edge: pivot with only values and no row/column fields prints the values only, no empty headers.
- Edge: field index -2 (values axis) in columns prints as `Values`, no crash.
- Error: in-memory pivot with `cache=None` or unreadable source prints `source unknown` and a warning; no exception.
- Size: pivot-cache memory is bounded only by openpyxl's own load; note this in the PR, no extra check possible.
- Warning: `refreshOnLoad` set prints a warning line.

**Verification:** formatter returns expected lines for in-memory objects; no exception on partial objects.

### U2. Wire into `converter()` and the legend

**Goal:** the section appears in output; the AI knows how to read it; nothing changes without pivots.

**Requirements:** R5, R6, R7, R8

**Dependencies:** U0, U1

**Files:** `excel_ai.py` (`converter`, `FORMAT_MD`, `SKILL_MD`, `autoteste`)

**Approach:**
- Call `secao_pivots` per sheet in the same loop that builds the map and extras; append warnings to the existing `avisos` list so they land in `## Warnings`.
- Add one `## Pivots` bullet to `FORMAT_MD` and `SKILL_MD` (they are embedded in the script; `instalar_skill` copies from there).
- Extend `autoteste` with the in-memory pivot cases plus an assertion that the existing fixture output is unchanged.

**Test scenarios:**
- Covers R5. Existing fixture converts to the exact same text as before: assert no `## Pivots` and no `# Workbook` appear, and that header/legend lines change only when a section is present. Merge order for the five plans: 008, 011, 009, 010, 012.
- Integration: a workbook with a pivot injected into the sheet shows `## Pivots` after `## Notes`, and the self-check still passes.
- `--celula` on a cell inside a pivot range still works (pivot section not required in chain output).
- `--instalar-skill` writes a skill that mentions `## Pivots`.

**Verification:** `python excel_ai.py` (auto-test) passes; README section added with one example.

## Risks and Dependencies

- `ws._pivots` is not public API; a future openpyxl may rename it. Mitigation: `getattr` guard plus the auto-test failing loudly.
- Pivots whose cache was dropped by another tool: handled by the "source unknown" path.
- No real pivot workbook available in the repo (confidential). Assumption: validate once on a real file by hand and record the result in the PR.

## Verification Contract

- Auto-test passes on Python 3.8 and current.
- Converting a real workbook with a pivot shows a correct block (manual check, not committed).

## Definition of Done

R1 to R8 met, legend updated, README updated, auto-test green, existing fixture output unchanged.
