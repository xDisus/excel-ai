---
title: "feat: List external workbook links"
type: feat
date: 2026-10-17
artifact_contract: ce-unified-plan/v1
status: completed
artifact_readiness: implementation-ready
execution: code
product_contract_source: ce-plan-bootstrap
depth: lightweight
---

# feat: List external workbook links

## Summary

Add a `## External links` section that maps each `[1]`, `[2]` in formulas to the file it points to. Today a formula like `=[1]Taxas!B2` is kept as text and the map only says "N formulas point to other files" (`excel_ai.py:741`), never which file. Fourth of five plans; independent of the others.

## Problem Frame

The AI sees `[1]Taxas!B2` with no way to know `[1]` is `C:\Tabelas\taxas.xlsx`. openpyxl loads these as `wb._external_links`, each with a file link target and cached sheet names, so the facts are available without new code paths for parsing.

## Requirements

- R1. One line per external link, in index order: `[N] <file name or target>` plus the cached sheet names when present.
- R2. Show the file name only (via `redigir`, plan 008 U0), not the full local path. Full paths expose user names and folder structure. URL targets keep host plus file name; query, fragment and userinfo removed. Decode `%20` and strip `file:///`.
- R3. Each line says how many formulas use it, counted per index with `re.findall(r"\[(\d+)\](?=[A-Za-z_'])", t)` into a Counter (the existing `excel_ai.py:741` check is a per-formula boolean, so it is not reused). Count from the unfiltered formulas of each sheet, not the `--celula` chain.
- R4. Warn when a link has no cached values (formula results from it cannot be trusted) or the target is a URL.
- R5. Workbook without external links: output identical to today.
- R6. Legend documents `## External links` and the `[N]` convention.
- R7. Single file, openpyxl only, Python 3.8, no network. Never opens the linked file.

## Scope Boundaries

- Not resolving or reading the linked workbook. Not following links recursively.
- Not rewriting `[N]` inside formulas (the self-check compares formulas and must stay lossless).

### Deferred to Follow-Up Work

- Optional `--paths` flag to show full paths, only if asked.

## Key Technical Decisions

- KTD1. Single ASCII file, no new deps, no network (session-settled: user-directed, from earlier plans).
- KTD2. Use `wb._external_links` (guarded with `getattr`, same caution as pivots). Chosen over parsing `xl/externalLinks/*.xml`: less code; fall back to empty on any error.
- KTD3. File name only (R2). Chosen over full path: public repo output may be shared; paths leak identity.
- KTD4. Index order matters: `[N]` is the 1-based position in the workbook's external-link list; do not sort.

## Implementation Units

### U1. Link listing

**Goal:** `links_externos(wb, textos_por_aba)` returns lines and warnings.

**Requirements:** R1 to R4, R7

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
- Read target from each link's `file_link.Target`; keep the last path component, handling both `/` and `\` separators.
- Target is `link.file_link.Target`. Sheet names are `link.externalBook.sheetNames.sheetName`. Cache present means `externalBook.sheetDataSet` is not `None`.
- Count usage by index from each sheet's formula texts with the R3 pattern; sum across sheets.
- Note: on load, openpyxl raises IndexError if a link's rels file is missing, before our code runs. Out of scope.

**Test scenarios:**
- Happy path: two links list `[1] taxas.xlsx` and `[2] premissas.xlsx` in order with sheet names.
- Covers R2. Target `C:\Users\ana\Docs\taxas.xlsx` prints `taxas.xlsx`, not `ana` or the folder.
- Usage count: three formulas on one sheet using `[1]` and one on another prints `4 formulas` for `[1]`.
- Warning: URL target and a link with no cache each add a warning line.
- Error: link object missing `file_link` prints `[N] unknown target`, no exception.
- Edge: Target with backslashes and with `%20` prints the decoded file name; a URL with `?sig=zzz` and `user:pw@host` prints neither.
- Edge: a structured reference or array constant like `{1,2}` or `T[Col]` is not counted as a link.

**Verification:** no test output contains a planted path component other than the file name.

### U2. Section, legend, docs

**Goal:** section in output, legend and README updated.

**Requirements:** R5, R6

**Dependencies:** U1

**Files:** `excel_ai.py` (`converter`, `FORMAT_MD`, `SKILL_MD`, `autoteste`), `README.md`

**Approach:**
- Place `## External links` after the last of the optional sections; keep the existing per-sheet "point to other files" count in the map unchanged.
- Fixture: build the link in memory (verified to save and reload): `ExternalBook(sheetNames=ExternalSheetNames(["T"]), id="rId1")`, set `link.file_link = Relationship(...)` with the Target, append to `wb._external_links`, save. Formula `=[1]T!B2` survives the round trip.
- Append `## External links` to `secoes_livro` (the `# Workbook` block, after Queries). Depends on plan 008 U0.

**Test scenarios:**
- Covers R5. Existing fixture output unchanged.
- Integration: injected-zip workbook shows the section and the self-check passes.
- Folder and `--diff` modes run without error.

**Verification:** auto-test green.

## Risks and Dependencies

- Path leakage: handled by R2/KTD3, tested.
- `_external_links` is semi-private; guarded.

## Verification Contract

Auto-test passes; manual check on one real workbook with links (not committed).

## Definition of Done

R1 to R7 met, legend and README updated, auto-test green, existing fixture output unchanged.

## Achados de execuções anteriores

- Plano 008 (PR #14) entregou `redigir`, `ler_membro`, `xml_seguro`, `_injetar`, `bloco_livro` e a lista `secoes_livro` em `converter()` (emitida só quando `cadeia is None`). Para usar: `secoes_livro.append(linhas)`.
- `ler_membro` levanta `KeyError` (membro ausente) e `ValueError` (acima do teto); `xml_seguro` levanta `ValueError` (DOCTYPE ou XML ilegível). Capture os dois.
- `redigir` é best-effort: valor sem aspas de `password=`/`key=` vai até `;`, `,`, `&` ou fim de linha; caminhos POSIX só em /home,/Users,/root,/mnt,/media,/var,/tmp,/opt. Nos testes, plante segredos reais (`hunter2`) e confira que sumiram E que o resto ficou.
- Teste só com venv: `/tmp/claude-0/venv/bin/python excel_ai.py` (sem openpyxl no sistema). Mantenha o fonte em ASCII e compatível com Python 3.8.
- Verificação R5: compare `converter()` de fixture, `--celula` e `--diff` contra a saída da `main` antes do plano.

## Resultado da execução

Entregue no PR #15 (base: feat/pivot-tables, depende do #14). `secoes_livro` virou dict com ordem fixa `ORDEM_LIVRO`. Revisão: contagem por fórmula (não por ocorrência) e query de alvo local corrigidas. P2 abertos: `"[1]x"` dentro de string conta como uso; caminho de URL é mantido (só userinfo/query saem); nomes de aba do link não passam por `redigir`. Validação em planilha real: pendente.
