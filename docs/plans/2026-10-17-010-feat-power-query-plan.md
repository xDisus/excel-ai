---
title: "feat: Read Power Query (M code)"
type: feat
date: 2026-10-17
artifact_contract: ce-unified-plan/v1
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: standard
---

# feat: Read Power Query (M code)

## Summary

Add a `## Queries` section with the name and M source of each Power Query, plus where its result lands. Third of five plans; depends only on plan 008 U0 (shared helpers). It no longer depends on plan 009. The M code is stored in a `DataMashup` blob that is plain stdlib to unpack (base64, then a small binary header, then a zip holding `Formulas/Section1.m`).

## Problem Frame

Power Query is where the real data logic often lives: sources, filters, merges, type changes. The sheet only shows the loaded result. Without the M code the AI cannot say where numbers come from or what was filtered out.

## Requirements

- R1. For each query: name, M code (verbatim, truncated at 2000 chars with a marker), and load destination when known (sheet table, connection only, data model).
- R2. M code goes through `redigir` (plan 008 U0): local paths reduced to file names, URL query/userinfo stripped, emails masked, credential-position literals masked. Otherwise verbatim. Redact first, truncate after. Best-effort: the legend tells the user to review before sharing.
- R7. Limits: at most 50 queries; blob and inner zip read through `ler_membro` caps; any DOCTYPE/ENTITY in the `customXml` item refuses it.
- R3. Parse failure or unknown blob version prints `Queries present but unreadable` once, with no exception.
- R4. Workbook without a `DataMashup`: output identical to today.
- R5. Legend (`FORMAT_MD`, `SKILL_MD`) documents `## Queries` and says the M code is shown as written, not evaluated.
- R6. Single file, stdlib + openpyxl, Python 3.8, no network. Never runs or refreshes a query.

## Scope Boundaries

- Not evaluating M, not previewing data, not summarizing what a query "does" (guessing).
- Query dependency graph between queries: out. M code already shows `Source = OtherQuery`.

### Deferred to Follow-Up Work

- Data-model (Power Pivot) measures and DAX.

## Key Technical Decisions

- KTD1. Single ASCII file, no new deps, no network (session-settled: user-directed, from earlier plans).
- KTD2. Find the blob by scanning `customXml/item*.xml` for a root `DataMashup` element; decode base64; layout is version (4 bytes) + package length (4 bytes, little endian) + package zip. Read only the package part with `zipfile` on a `BytesIO`. Chosen over any external M parser: only extraction is needed.
- KTD3. M code verbatim, redaction only of credential-position literals. Chosen over summarizing: R3 of earlier plans (facts, never inferred).
- KTD4. Split `Section1.m` into queries on `shared <Name> =` declarations; keep the raw text per query.
- KTD5. Destination (`FillTarget`, `FillEnabled`, `LoadedToAnalysisServices`) comes from the DataMashup metadata part, not from connections. Layout per MS-QDEFF: version, package length, package, permissions length, permissions, metadata length, metadata, bindings; each part is length-prefixed. Parse the metadata XML for the destination; if that fails, omit the destination, never guess. Wrap the whole extraction so any failure yields only the R3 marker.
- KTD6. The blob may be UTF-16 with BOM; decode `utf-16` when a BOM is present. The root element is namespaced and the base64 may hold whitespace: match `{*}DataMashup` and strip whitespace before decoding. Split `Section1.m` on `shared` at line start only (not inside strings/comments).

## Implementation Units

### U1. DataMashup extractor

**Goal:** `consultas(fonte)` returns `[(nome, m_codigo)]`.

**Requirements:** R1, R3, R6

**Dependencies:** none (destination lookup in U2)

**Files:** `excel_ai.py`

**Approach:**
- Zip-open pattern as in `tem_vba` (`excel_ai.py:404`).
- Handle the UTF-8 BOM and UTF-16 variants seen in `customXml` items.
- Unpack in order: base64, header bytes, inner zip, `Formulas/Section1.m`, split on `shared` declarations.

**Test scenarios:**
- Happy path: a hand-built blob (base64 of header + inner zip with a two-query `Section1.m`) returns two queries with exact text.
- Edge: blob with zero queries returns empty list.
- Error: truncated base64 or inner zip not a zip returns the unreadable marker, no exception.
- Edge: query name with spaces/quotes (`#"Vendas 2024"`) is returned without the `#"` wrapper characters dropped incorrectly.

**Verification:** extractor round-trips a synthetic blob built in the test.

### U2. Redaction, section, legend

**Goal:** safe output in the file, legend updated.

**Requirements:** R1, R2, R4, R5

**Dependencies:** U1, plan 008 U0

**Files:** `excel_ai.py` (`converter`, `FORMAT_MD`, `SKILL_MD`, `autoteste`), `README.md`

**Approach:**
- Redact with one regex pass over credential keywords in M text (key = "literal"); keep it small and tested.
- Truncate each query at 2000 chars with a marker; cap 50 queries.
- Append `## Queries` to `secoes_livro` (the `# Workbook` block, after Connections).
- Add a test where M text contains `File.Contents("C:\Users\ana\x.xlsx")` and a URL with `?sig=zzz`: output has `x.xlsx`, no `ana`, no `zzz`.
- Add a test where the destination (metadata XML) says the query loads to a sheet table.

**Test scenarios:**
- Covers R2. M with `[Password = "hunter2"]` and `Web.Contents(..., [Headers=[Authorization="Bearer abc"]])` prints neither `hunter2` nor `abc`.
- Covers R4. Existing fixture output unchanged.
- Integration: fixture zip with injected `customXml/item1.xml` shows `## Queries`; self-check passes.
- Truncation: a 5000-char query prints 2000 chars plus the marker.

**Verification:** auto-test green; planted-secret test green.

## Risks and Dependencies

- `DataMashup` layout is documented (MS-QDEFF) but varies; failure must degrade to the one-line marker (R3).
- Redaction by regex can miss odd credential spellings. Public repo + confidential files: keep a conservative keyword list and say so in the legend; users should still glance at the output before sharing.
- Real workbook needed once for manual validation (not committed).

## Verification Contract

Auto-test passes; planted-secret test passes; manual check on one real workbook.

## Definition of Done

R1 to R6 met, legend and README updated, auto-test green, existing fixture output unchanged.
