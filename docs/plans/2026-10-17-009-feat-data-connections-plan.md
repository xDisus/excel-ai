---
title: "feat: Read data connections"
type: feat
date: 2026-10-17
artifact_contract: ce-unified-plan/v1
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: lightweight
---

# feat: Read data connections

## Summary

Add a `## Connections` section listing the workbook's data connections (database, web, text file, Power Query link): name, type, source, and refresh flags. Second of five plans. openpyxl does not read `xl/connections.xml`, so this reads it straight from the zip with the stdlib.

## Problem Frame

A sheet fed by a connection looks like constants. The AI cannot tell the numbers are loaded from a database or a URL and may be stale. The connection definitions are small XML and hold the facts: description, type, connection string, command (query or table), refresh-on-open.

## Requirements

- R1. One line per connection: id/name, type (ODBC, OLE DB, web, text, Power Query), source, command text (truncated), refresh on open, background refresh.
- R2. Credentials never printed. Strip `Password=`, `PWD=`, `User ID`/`UID` values and any `token=`-like pair from connection strings; print `***`.
- R3. A connection whose provider is the Power Query mashup provider is labeled `Power Query: <name>` and its details are left to plan 010.
- R4. Warn when any connection refreshes on open.
- R5. Workbook without connections: output byte-identical to today.
- R6. Legend (`FORMAT_MD`, `SKILL_MD`) documents `## Connections`.
- R7. Single file, stdlib + openpyxl only, Python 3.8, no network. Never opens the connection.

## Scope Boundaries

- Not testing or refreshing connections. Not reading cached query results.
- Pivot-to-connection linkage: mention only if the pivot plan (008) already exposes the cache source; otherwise deferred.

### Deferred to Follow-Up Work

- Power Query details (010), external links (011), macros (012).

## Key Technical Decisions

- KTD1. Single ASCII file, no new deps, no network (session-settled: user-directed, from earlier plans).
- KTD2. Parse `xl/connections.xml` with `zipfile` + `xml.etree.ElementTree`. Chosen over openpyxl: it does not model connections.
- KTD3. Redaction is an allowlist-by-key approach on the connection string: split on `;`, keep `Data Source`, `Initial Catalog`, `Provider`, `Driver`, `Database`, `Server`; drop everything else. Chosen over a denylist of secret keys: a denylist leaks unknown secret keys. Public repo and confidential work files make a leak the worst failure. Kept values, `webPr url`, `textPr sourceFile`, `odcFile` and command text all go through `redigir` from plan 008 U0 (file name only, no URL query/userinfo); command text is redacted, then truncated.
- KTD5. Depends on plan 008 U0 (`redigir`, `ler_membro`, `xml_seguro`, `_injetar`, `secoes_livro`). The section goes in the `# Workbook` block, first in the fixed order. Mashup provider test: provider string contains `Microsoft.Mashup`.
- KTD4. Namespace-tolerant tag matching (`{*}connection`), because Excel writes the main spreadsheetml namespace.

## Implementation Units

### U1. Connection reader

**Goal:** `conexoes(fonte)` returns a list of formatted lines.

**Requirements:** R1, R2, R3, R4, R7

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
- Next to `tem_vba` (`excel_ai.py:404`), same zip-open pattern (`seek(0)`, `BadZipFile` returns empty).
- Map the numeric `type` attribute to a word (1 ODBC, 2 DAO, 3 web, 4 text, 5 OLE DB; others `type N`).
- Apply the allowlist redaction (KTD3) to `connection` and `odcFile` strings; truncate command text at 200 chars with a marker.

**Test scenarios:**
- Happy path: ODBC connection with `Driver=...;Server=db;Database=x;UID=u;PWD=secret` prints server and database, no `u`, no `secret`.
- Web connection prints its URL; text connection prints the file path.
- Covers R3. Mashup-provider connection prints `Power Query: Name` only.
- Edge: connection string with unknown key `ApiKey=abc` prints no `abc`.
- Covers redaction: `Data Source=C:\Users\ana\x.accdb` prints `x.accdb`; a web URL with `?token=zzz` prints host and path without the token; a text connection prints the file name only.
- Size: oversized `connections.xml` member or a DOCTYPE with entities is refused with one warning; cap 100 connections with a marker.
- Edge: no `xl/connections.xml` in the zip returns empty. Corrupt XML returns empty plus one warning, no exception.
- Warning: `refreshOnLoad="1"` adds a warning line.

**Verification:** no test output contains any planted secret string.

### U2. Section, legend, docs

**Goal:** output, legend and README updated; no change without connections.

**Requirements:** R5, R6

**Dependencies:** U1

**Files:** `excel_ai.py` (`converter`, `FORMAT_MD`, `SKILL_MD`, `autoteste`), `README.md`

**Approach:**
- Call once per workbook (not per sheet) in `converter` and append `## Connections` to `secoes_livro` (the `# Workbook` block).
- Add warning lines to the existing `avisos` list.
- Build the test workbook with `_injetar` (plan 008 U0) adding a hand-made `xl/connections.xml` to the fixture zip. The existing macro test (`excel_ai.py:1613`) only feeds `tem_vba` a standalone zip, so it is not a model for end-to-end tests.

**Test scenarios:**
- Covers R5. Existing fixture output unchanged.
- Integration: fixture with injected `connections.xml` shows the section; self-check still passes.
- `--diff` and folder modes run without error on a workbook with connections.

**Verification:** auto-test green; README gets a short entry.

## Risks and Dependencies

- Secret leakage is the main risk: allowlist (KTD3) plus a test that plants secrets in several key spellings.
- Connection XML varies by Excel version; unknown attributes are ignored, never fatal.

## Verification Contract

Auto-test passes; planted-secret test passes; manual check on one real workbook (not committed).

## Definition of Done

R1 to R7 met, legend and README updated, auto-test green, existing fixture output unchanged.

## Achados de execuções anteriores

- Plano 008 (PR #14) entregou `redigir`, `ler_membro`, `xml_seguro`, `_injetar`, `bloco_livro` e a lista `secoes_livro` em `converter()` (emitida só quando `cadeia is None`). Para usar: `secoes_livro.append(linhas)`.
- `ler_membro` levanta `KeyError` (membro ausente) e `ValueError` (acima do teto); `xml_seguro` levanta `ValueError` (DOCTYPE ou XML ilegível). Capture os dois.
- `redigir` é best-effort: valor sem aspas de `password=`/`key=` vai até `;`, `,`, `&` ou fim de linha; caminhos POSIX só em /home,/Users,/root,/mnt,/media,/var,/tmp,/opt. Nos testes, plante segredos reais (`hunter2`) e confira que sumiram E que o resto ficou.
- Teste só com venv: `/tmp/claude-0/venv/bin/python excel_ai.py` (sem openpyxl no sistema). Mantenha o fonte em ASCII e compatível com Python 3.8.
- Verificação R5: compare `converter()` de fixture, `--celula` e `--diff` contra a saída da `main` antes do plano.
