## Residual Review Findings

Review done in-process (no subagents, no cross-model peer). Open, all P3:

- P3 excel_ai.py `converter_diff`: data-table (what-if) cells are not compared, so a change there is invisible in the diff.
- P3 excel_ai.py `converter_diff`: inserting rows/columns shows as many changed formulas (no move detection). Documented in README and plan scope.
- P3 excel_ai.py `converter_diff`: no cap on formula sections (only constants are capped at 200 lines).
- P3 excel_ai.py `main_diff`: strict argv shape (`a b --diff` only), same as `--celula`.
