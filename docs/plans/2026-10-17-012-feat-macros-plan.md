---
title: "feat: Read VBA macros"
type: feat
date: 2026-10-17
artifact_contract: ce-unified-plan/v1
status: deferred
artifact_readiness: deferred
execution: code
product_contract_source: ce-plan-bootstrap
depth: deep
---

# feat: Read VBA macros

## Summary

Today the tool only warns "File has VBA macros" (`excel_ai.py:786`). This plan adds a `## Macros` section with module names and, in a second step, the VBA source of each module, so the AI can see which values a macro writes. Fifth of five plans, independent, and the riskiest: `vbaProject.bin` is an OLE (CFB) container with compressed module streams, and there is no stdlib reader.

## Problem Frame

The warning says values "may be written by a macro" but gives nothing to check. Calculators with macros (goal seek loops, copy-paste-values buttons, hardcoded writes) are exactly where formulas stop explaining the numbers.

## Requirements

- R1. Default tier: list module names only (from the CFB `VBA` storage stream names, excluding `dir`, `_VBA_PROJECT`, `__SRP_*`). No types, no procedures, no source.
- R2. Opt-in tier, flag `--macros-fonte`: module types and `Sub`/`Function` names, and module source truncated per module (4000 chars) and in total (20000 chars) with a marker. Source goes through `redigir` (plan 008 U0), which also masks string literals that look like paths, URLs or connection strings. Default stays off because macro source often embeds credentials, servers and paths.
- R3. Highlight sheet-touching writes: list lines that assign to `Range(`/`Cells(`/`.Value` as a short "writes to cells" index. Facts from text, no interpretation.
- R4. Failure to parse any stream degrades to the module-name list or to today's warning; never crash, never block the formulas output.
- R5. Files without macros: output identical to today. The existing warning stays when macros exist.
- R6. Legend documents `## Macros` and states the code is shown as written, not run.
- R7. Single file, stdlib + openpyxl, Python 3.8, no network. Never executes anything.

## Scope Boundaries

- Not running, deobfuscating or analyzing macro behavior. No security verdict: this is a reading aid, not malware analysis. Password-protected or "unviewable" VBA projects: report as protected and stop.
- Not editing macros.

### Deferred to Follow-Up Work

- Mapping macro writes to the exact cells they target when addresses are computed at run time.

## Key Technical Decisions

- KTD1. Single ASCII file, no new deps, no network (session-settled: user-directed, from earlier plans). This rules out `oletools`/`olefile`, the usual choice, so a minimal CFB reader is written in the script.
- KTD2. Two-step delivery. Step A (U1+U2): CFB directory read, which yields module names without decompression. Step B (U3): MS-OVBA decompression for source. Chosen over doing both at once: Step A already gives most of the value and is low risk; Step B can ship or be dropped on its own.
- KTD3. Minimal CFB reader limited to what is needed: header, FAT, mini FAT, directory entries, stream read. Version 3 (512-byte sectors) first; version 4 only if a real file needs it.
- KTD4. `dir` stream decompression and module source offsets follow MS-OVBA; implemented as small pure functions covered by golden tests built from a known compressed sample.
- KTD5. Redaction is the shared `redigir` from plan 008 U0, plus masking of string literals that look like paths, URLs or connection strings. Source stays opt-in (R2).
- KTD6. Decision gate: U3 (decompression and source) ships only after validation on a real `.xlsm`. If no real file is available or the result is flaky, ship U1+U2 only and say in the PR that the source tier is deferred. U1+U2 alone give little over today's warning, so also consider deferring this whole plan until someone asks.
- KTD7. CFB limits: handle DIFAT chains or reject files whose FAT does not fit one DIFAT level; cap sector count and stream size; v4 (4096-byte sectors) accepted only if a real file needs it. Password protection does not encrypt source; policy is to report "protected" and still list module names, and not show source.

## Implementation Units

### U1. CFB container reader

**Goal:** list streams in `xl/vbaProject.bin` and read one by name.

**Requirements:** R4, R7

**Dependencies:** none

**Files:** `excel_ai.py`

**Approach:**
- Use `struct` over the bytes from `zipfile`, as `tem_vba` already opens the zip (`excel_ai.py:404`).
- Implement header parse, FAT chain walk, mini-stream for small streams, directory tree walk (storages `VBA`, streams `dir`, module streams).
- Guard against cycles and absurd sizes (max sectors, max stream size).

**Test scenarios:**
- Happy path: a synthetic CFB built in the test (header, one FAT sector, directory with `VBA/dir` and two module streams) lists the streams and returns exact bytes.
- Edge: stream under 4096 bytes is read through the mini stream, over it through the FAT.
- Error: truncated file, bad signature, or FAT cycle returns "unreadable" and no hang or exception.
- Error: declared size larger than the file is rejected.

**Verification:** reader round-trips the synthetic container.

### U2. Module and procedure listing

**Goal:** `## Macros` section with names and procedures; no decompression yet.

**Requirements:** R1, R4, R5, R6

**Dependencies:** U1

**Files:** `excel_ai.py` (`converter`, `FORMAT_MD`, `SKILL_MD`, `autoteste`), `README.md`

**Approach:**
- Module names come from the CFB stream names under `VBA` (excluding `dir`, `_VBA_PROJECT`, `__SRP_*`); no decompression in U2.
- Keep the current warning line and append `## Macros` to `secoes_livro` (the `# Workbook` block, last in the fixed order). Depends on plan 008 U0.
- Procedure names and types need the decompressed `dir` stream, so they belong to U3 behind `--macros-fonte`.

**Test scenarios:**
- Covers R5. Existing fixture output unchanged. The fake `vbaProject.bin` (`b"x"`) injected with `_injetar` into the fixture and run through `converter` yields the old warning only, with no crash (the existing test at `excel_ai.py:1613` only exercises `tem_vba`).
- Integration: injected synthetic `vbaProject.bin` shows module names and keeps the warning.
- CFB: a v3 container with a mini-stream `dir`, and one needing a DIFAT entry, both read correctly; an oversized sector count is rejected.
- Protected project case: reported as protected, formulas output unaffected.

**Verification:** auto-test green.

### U3. VBA source extraction

**Goal:** decompress module source, list procedures and cell writes.

**Requirements:** R2, R3, R4

**Dependencies:** U1, U2

**Files:** `excel_ai.py`

**Approach:**
- Implement MS-OVBA decompression (chunked, flag-byte, copy tokens) as a pure function.
- Parse the `dir` stream records needed: module name, stream name, text offset.
- Strip the performance cache before `text offset`, decompress the rest, decode with the project code page (fall back to latin-1).
- Extract `Sub`/`Function` names and cell-write lines with simple regexes. Apply credential redaction (KTD5) and truncation caps.

**Test scenarios:**
- Golden: a known MS-OVBA compressed sample decompresses to its known text.
- Happy path: a module with two `Sub`s and a `Range("B2").Value = 5` line lists both procedures and the write.
- Redaction: `password = "abc"` in source prints without `abc`.
- Truncation: a 10000-char module prints 4000 chars plus the marker; total cap enforced across modules.
- Error: corrupt compressed chunk keeps the module name, omits source with a marker, other modules still shown.

**Verification:** golden test and error test pass; output never exceeds the caps.

## Risks and Dependencies

- Largest risk: no real macro workbook in the repo (confidential) and hand-built synthetic containers may not match what Excel writes. Mitigation: validate U1 to U3 on at least one real `.xlsm` by hand before merging; ship U1+U2 alone if U3 proves flaky.
- Hand-written CFB and OVBA code is the largest addition to a one-file tool (rough estimate: 200 lines). If that is too heavy for the single file, drop U3 and keep U1+U2.
- Output size: macros can be big; caps are mandatory.
- Untrusted input: parser must never loop or allocate unbounded memory on crafted files (limits in U1).

## Verification Contract

Auto-test passes including golden and error cases; manual check on one real `.xlsm`.

## Definition of Done

R1 to R7 met for the shipped tiers, caps enforced, legend and README updated, auto-test green, existing fixture output unchanged.

## Achados de execuções anteriores

- Plano 008 (PR #14) entregou `redigir`, `ler_membro`, `xml_seguro`, `_injetar`, `bloco_livro` e a lista `secoes_livro` em `converter()` (emitida só quando `cadeia is None`). Plano 011 mudou isso: `secoes_livro` é um dict, use `secoes_livro["Connections"|"Queries"|"Macros"] = ["## Nome"] + linhas`; a ordem fixa vem de `ORDEM_LIVRO`. Linha de legenda no cabeçalho só quando a seção existe (veja `External links` em `converter`).
- `ler_membro` levanta `KeyError` (membro ausente) e `ValueError` (acima do teto); `xml_seguro` levanta `ValueError` (DOCTYPE ou XML ilegível). Capture os dois.
- `redigir` é best-effort: valor sem aspas de `password=`/`key=` vai até `;`, `,`, `&` ou fim de linha; caminhos POSIX só em /home,/Users,/root,/mnt,/media,/var,/tmp,/opt. Nos testes, plante segredos reais (`hunter2`) e confira que sumiram E que o resto ficou.
- Teste só com venv: `/tmp/claude-0/venv/bin/python excel_ai.py` (sem openpyxl no sistema). Mantenha o fonte em ASCII e compatível com Python 3.8.
- Verificação R5: compare `converter()` de fixture, `--celula` e `--diff` contra a saída da `main` antes do plano.

## Resultado da execução (adiado, não implementado)

Decisão do ship-plans, seguindo o gate KTD6: não há `.xlsm` real no repositório nem na máquina, então nem o leitor CFB (U1) nem a listagem de módulos (U2) poderiam ser validados contra o que o Excel realmente grava; um CFB sintético escrito por nós só prova que lemos o que nós mesmos escrevemos. U1+U2 sozinhos dão só nomes de módulos, pouco além do aviso atual (`File has VBA macros`), e acrescentam ~100 linhas de parser de formato binário em entrada não confiável. U3 (decompressão MS-OVBA e código-fonte) exige o mesmo arquivo real.

**Para retomar:** precisa de um `.xlsm` real (pode ser anonimizado) fornecido pelo usuário, e que ele confirme que quer a leitura de macros. Depois: trocar `status`/`artifact_readiness` de volta para `implementation-ready`, e rodar `/ship-plans`. Os helpers `redigir`, `ler_membro` e `secoes_livro["Macros"]` já existem (plano 008/011).
