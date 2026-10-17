# excel_ai output format

## File layout

```
# excel_ai dump
<how-to-read lines>

## Warnings
- ...

## Map
- 'Sheet' A1:J101, 303 formulas, uses 'Other' (200)
  table Tabela1 A1:B4 columns: Nome, Valor
Defined names:
  TAXA2024 = Premissas!$B$1

# Sheet 'Fluxo'
## Formulas
...
## Data
...
```

## Formula lines

| Line | Meaning |
|---|---|
| `D2:D101 =B2*C2` | Formula of D2; D3..D101 are the same formula filled down (relative references shift as in Excel). |
| `C2:E2 =C1*2` | Same, filled across. |
| `D2:D101 [Juros] =B2*C2` | `[Juros]` is the nearest text header: above for columns, left for rows, up to 2 empty cells away. A hint, not part of the formula. |
| `D2:D101 [Juros] (v 12.5) =B2*C2` | `(v 12.5)` is the cached value (last Excel result) of the FIRST cell. Only blocks of 2+ cells. Hint only. |
| `  except D30 =B30*C30*2` | D30 inside the block above has a different formula. |
| `  except D30 value 5` | D30 is a constant inside the block. |
| `  except D30 blank` | D30 is empty. In `--celula` mode: `not in chain` (exists but not used by the target). In a copy-sheet: `unchanged`. |
| `E1 array E1:E2 {=Z1:Z2*2}` | Array formula anchored at E1 covering E1:E2. Other cells of the range are not listed. |
| `same formulas as 'Fluxo', except:` | The sheet is a copy (80%+ equal) of that sheet; only differing cells follow. `A1 no formula` means the cell has no formula here. |

What-if data-table cells are not listed as lines; they only appear counted in Warnings.

Internal prefixes (`_xlfn.`) are removed: `_xlfn.XLOOKUP` is shown as `XLOOKUP` (PROCX in Portuguese Excel).

## Data lines

| Line | Meaning |
|---|---|
| `A1 Parcela \| B1 Saldo` | Constants on a row, `cell value` separated by `\|`. |
| `A2:A101 [100 numbers] min 1 max 100 first: 1, 2, 3` | Long run of numbers/dates summarized. Runs of up to 10 stay complete. |
| `A2:A69 [68 texts, 68 distinct] first: a, b, c` | Long run of texts (over 50) summarized. Shorter runs stay complete. |
| `A5:A9 x (x5)` | Same text repeated. |

## Warnings you may see

- `N formulas contain #REF!` broken reference.
- `N formulas show an error value (#N/A, #DIV/0!...)` from cached values.
- `N formulas use OFFSET/INDIRECT` target decided at run time; not followable statically.
- `N data-table (what-if) cells`.
- `File has VBA macros`.
- `No calculated values saved in the file`.
- `compact view failed the self-check on N sheet(s)` those sheets are one formula per cell.
- `--celula` only: `unresolved` / `formula not parsed` refs that could not be followed.

## Diff file

```
# excel_ai diff (old -> new)
Summary: N formulas new or changed, M old formulas replaced or removed, K constants changed, S sheet(s) without changes.
## Sheets          (added/removed sheets)
## Defined names   (added:/removed:)
# Sheet 'X'
## Formulas new or changed
## Formulas before (changed or removed cells)
## Constants
```
