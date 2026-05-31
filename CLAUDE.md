# CLAUDE.md

Guidance for Claude Code (and other AI assistants) working in this repository.

## What this repository is

This is a **data repository**, not a software application. It contains two
Microsoft Excel (`.xlsx`) workbooks built around **continuous glucose
monitoring (CGM)** data, used to study the effect of **metformin** on glucose
control and variability.

| File | Size | Purpose |
|------|------|---------|
| `LingoCGM_CGM_Report_codex.xlsx` | ~3.5 MB | Single-subject CGM report. Computes glycemic metrics (mean, SD, CV, GMI, TIR, AUC, risk indices) and an **Off Meds vs On Meds** comparison, all via live Excel formulas over a raw-data sheet. |
| `LingoCGM_Metformin_Study_Analysis.xlsx` | ~76 KB | Multi-participant study workbook (intake of pasted CGM exports, per-participant metrics, cohort aggregation, safety QA). |

> ⚠️ The glucose data is **simulated for demonstration** (the workbooks say so in
> their own Review Notes). Treat it as sample data, not clinical evidence.

## File structure

### `LingoCGM_CGM_Report_codex.xlsx` (11 sheets)
- **`INPUT`** (≤20,000 rows) — raw pasted CGM export (timestamp, glucose, etc.).
- **`Daily Log`** — self-reported daily diary (weight in lb, time med taken,
  flavor, taste & aftertaste 1–5). Col A = date, B = weight, C = dose time,
  D = flavor, E = taste, F = aftertaste, G = notes, H = auto phase tag
  (Off/Washout/On from `INPUT!I1`/`I2`). Summary panel in cols J–L (values in
  **K3:K14**) feeds `Study Export` cols **BX–CH**.
- **`_Calc`** (≤20,000 rows) — engine sheet. Col **B** = glucose (mg/dL), col
  **C** = date, col **D** = hour, col **E** = `Off Meds`/`On Meds`/washout tag,
  cols **G–N** = per-row risk/AUC helper terms. **This is the source of truth**
  every other sheet references.
- **`Summary`** — whole-dataset metrics. SD via `STDEVP`, CV, GMI, TIR tiers,
  AUC, MODD, J-Index, M-Value, GRI. (Verified correct.)
- **`Comparative`** — the headline **Off Meds vs On Meds** table. ⚠️ Contains
  broken formulas — see `README.md` § "Issues in the standard-deviation file".
- **`Daily`**, **`AGP`**, **`Time Blocks`** — day-level and time-of-day rollups.
- **`Glossary`** — metric definitions and methodology notes (incl. the
  "Conditional Standard Deviation Workaround").
- **`Study Export`**, **`Review Notes`** — export row and known-limitations log.
  `Study Export` carries 75 glucose fields (A–BW) **plus 11 daily-log fields
  (BX–CH)** for cohort transfer.

Config inputs live in `INPUT!I1` (metformin start date) and `INPUT!I2`
(washout-buffer days, default 4). These also drive the `Daily Log` phase tags.

### `LingoCGM_Metformin_Study_Analysis.xlsx` (8 sheets)
`Dashboard` · `Study_Setup` · `Source_Files` · `File_Intake` · `Pasted_Exports`
· `Participant_Metrics` · `Cohort_Stats` (aggregated mean/SD/CV per group) ·
`Safety_QA`. The daily-log fields ride the existing transfer chain: `File_Intake`
and `Pasted_Exports` mirror cols **BX–CH**; `Participant_Metrics` appends them as
**CO–CY** (the derived columns BX–CN are unchanged); `Cohort_Stats` adds four
paired rows (Weight Change lb / %, Avg Taste, Avg Aftertaste).

## Working with these files

- **Tooling:** Python with `openpyxl` (and/or `pandas`). Not in the base image —
  install with `pip install openpyxl pandas` at session start.
  - `load_workbook(path, data_only=True)` → cached **values** (what Excel last
    saved; a formula that errored is cached as its `IFERROR` fallback, e.g. `—`).
  - `load_workbook(path, data_only=False)` → **formulas**.
  - For the large CGM workbook use `read_only=True` when only reading.
- **Always dump both values *and* formulas** before judging a cell. A value of
  `—` usually means the formula errored, not that the data is missing.
- **`_Calc` is the source of truth.** Recompute any summary statistic from
  `_Calc!B` (glucose) filtered by `_Calc!E` (med tag) rather than trusting a
  summary cell.
- Excel-function gotcha: newer functions must carry the `_xlfn.` prefix in the
  stored XML (e.g. `_xlfn.MINIFS`). A `_xludf.` prefix makes Excel treat the
  function as an unknown user-defined function and return an error.

## Conventions

- **Branch:** develop on `claude/adoring-darwin-KhTd7`; create it locally if
  missing. Never push to `main` without explicit permission.
- **Commits:** clear, descriptive messages; keep doc changes separate from
  binary `.xlsx` changes when practical.
- **Push:** `git push -u origin <branch>`, retrying with exponential backoff on
  network errors. Do **not** open a pull request unless explicitly asked.
