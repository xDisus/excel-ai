---
name: excel-ai
description: Read Excel workbooks (.xlsx/.xlsm) efficiently by converting them to compact, lossless text before analysing them. Use whenever the user asks about a spreadsheet, its formulas, calculation logic, cross-sheet references (XLOOKUP/PROCX), financial calculators (IRR/TIR, duration, price schedules, provisioning/PDD), why a cell has a value, what changed between two versions of a workbook, or wants a whole folder of workbooks read. Triggers - planilha, Excel, xlsx, xlsm, formula, PROCX, TIR, calculadora, diff de versoes.
license: MIT
---

# excel-ai: read spreadsheets as compact text

`scripts/excel_ai.py` turns a workbook into a short text file that keeps every formula and constant that matters and drops formatting noise. A 1.6M-token cell-by-cell dump of a large financial calculator becomes about 33k tokens (about 48x smaller). Formulas are verified lossless: the text is expanded back and compared with the original.

Never read an `.xlsx`/`.xlsm` by unzipping it or dumping cells with openpyxl yourself. Run the tool and read its output.

## Setup (once)

```
pip install openpyxl
```

Python 3.8+. One file, no network, nothing leaves the machine. The script is at `scripts/excel_ai.py` in this skill folder. Run it with that path (written below as `excel_ai.py`).

## Pick the mode

| The user wants... | Run |
|---|---|
| To understand or audit one workbook | `python excel_ai.py book.xlsx` |
| One number explained (final IRR, a total, a provision) | `python excel_ai.py book.xlsx --celula "Sheet!B5"` |
| What changed between two versions | `python excel_ai.py old.xlsx new.xlsx --diff` |
| Every workbook in a folder | `python excel_ai.py folder` |

Prefer `--celula` over the full dump whenever the question is about one result: it keeps only the cells that result depends on (for a real calculator, about 6% of the formulas and 2k tokens instead of 33k). Prefer `--diff` over reading two dumps.

Output files are written next to the input:

- full: `book.ai.txt`
- precedents: `book.<Sheet>_<CELL>.ai.txt` (non-alphanumerics become `_`)
- diff: `new.vs_old.diff.ai.txt`
- folder: one `.ai.txt` per workbook (not recursive; `~$*` lock files skipped)

After running, read the generated `.ai.txt` file (not the workbook). The console also prints tokens before/after and the self-check result.

## `--celula` details

- Reference forms: `'Sheet name'!B5`, `Sheet!B5`, `Sheet!$b$5`, or `B5` (active sheet). One cell only, never a range.
- It follows the chain through other sheets, ranges, whole columns/rows (limited to existing cells), defined names, table references (`Table[Col]`, `[@Col]`) and array formulas.
- Not followed, but listed in Warnings: `OFFSET`/`INDIRECT` results, external-file references, data-table (what-if) inputs, names that are not plain ranges, `#REF!`, and LET/LAMBDA variables. If Warnings says "some refs could not be followed", tell the user the chain may be incomplete and offer the full dump.
- Lines `except <cell> not in chain` are gaps inside a block: formulas that exist but are not used by the target.
- The header shows the target's cached value (what Excel last calculated).

## `--diff` details

- Sheets match by exact name. Added/removed sheets are listed with formula counts.
- `## Formulas new or changed` is the NEW version in the normal compact format; `## Formulas before` is the OLD formula of changed or removed cells. Constants appear as `A1 old -> new`, `A1 added X`, `A1 removed (was X)` (200 lines per sheet max). Defined names appear as `added:`/`removed:` (a changed destination is removed + added).
- A pure fill-down with the same relative formula is not a change. Float noise below 1e-9 and empty strings are ignored.
- Inserting rows or columns shows up as many changes (no move detection). Data-table cells are not compared. Say so when the diff looks huge.
- "No changes found." means formulas, constants and names are identical.

## Reading the output

Read `references/format.md` for the full legend. Short version:

- `## Warnings` first: `#REF!`, error values, OFFSET/INDIRECT, macros, uncalculated files. Mention relevant ones to the user.
- `## Map`: each sheet (dimensions, HIDDEN, formula count, which sheets it uses), tables, defined names.
- Per sheet `## Formulas`: `E2:E361 [Juros] (v 12.5) =C2*G2` means the formula of the first cell, copied down/across the range with references shifting. `[Juros]` is the nearest text header (a hint). `(v 12.5)` is the value Excel last calculated for the first cell (a hint, blocks of 2+ cells only). `  except E9 =...` overrides one cell. `same formulas as 'X', except:` means a copy of sheet X with listed differences.
- Per sheet `## Data`: constants, with long number runs summarized as `[100 numbers] min .. max .. first: ..`.

Reason about the logic from the formulas: say what each block computes, which sheets feed it, and where inputs live. Use cached values `(v ...)` only as sanity checks; they may be stale.

## Exit codes and problems

- `0` ok. `1` the lossless self-check failed on some sheet (that sheet is listed one formula per cell, still correct, just larger). `2` bad arguments, missing file, unreadable workbook, or a folder with no workbooks.
- "No calculated values saved": the file was never opened in Excel, so cached values and error checks are missing.
- "File has VBA macros": some values may come from a macro, which this tool does not read.
- `.xls` (old format) and `.xlsb` are not supported; ask the user to save as `.xlsx`.
- Large workbooks take 15-40 seconds. Do not interrupt.

## Rules

1. Run the tool before answering any question about a workbook's contents.
2. Never paste real workbook data into public places (issues, PRs, public repos). Workbooks are usually confidential.
3. Do not edit the workbook with this tool; it is read-only.
4. If the user's question is about one result, start with `--celula`. If it is about differences between versions, use `--diff`.
5. When the output says a part could not be followed or verified, say it plainly instead of guessing.
