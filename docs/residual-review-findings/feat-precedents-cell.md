# Residual Review Findings — feat/precedents-cell

Source: ce-code-review (correctness, testing, adversarial) + ce-simplify-code, 2026-10-16. Applied fixes are in the branch; these were left open.

- P3 excel_ai.py `main` — output names collide for sheets that differ only in punctuation (`'a b'` vs `'a_b'`); second run overwrites.
- P3 excel_ai.py `precedentes` — data-table (what-if) cells are not expanded; their inputs are not followed (the per-sheet data-table warning still shows).
- P3 excel_ai.py `precedentes` — 3D refs (`S1:S3!A1`) and LET/LAMBDA variable names show as "unresolved reference" warnings (warned, not silent).
- P3 excel_ai.py `__main__` — argument handling is strict: only `file` or `file --celula REF`; `--celula=REF` is rejected.
- P3 excel_ai.py `converter` — in `--celula` mode every formula is still tokenized by `relativa()` before filtering (perf; ~14s on an 80k-formula workbook).
- P3 tests — table `headerRowCount`, external-file refs and the active-sheet default are not pinned by autoteste().
