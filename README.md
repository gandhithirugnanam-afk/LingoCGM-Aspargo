# LingoCGM-Aspargo

Excel workbooks for **continuous glucose monitoring (CGM)** analysis, focused on
measuring the effect of **metformin** on glucose control and variability.

> **Disclaimer:** The glucose data is **simulated for demonstration** (stated in
> the workbooks' own Review Notes). Nothing here is validated clinical evidence.

---

## Contents

### `LingoCGM_CGM_Report_codex.xlsx` (~3.5 MB, 10 sheets)
A single-subject CGM report driven entirely by live Excel formulas.
- **`INPUT`** — raw pasted CGM export.
- **`_Calc`** — calculation engine (glucose in col B, med-period tag in col E);
  the source of truth for every other sheet.
- **`Summary`** — whole-dataset metrics: mean, **SD (`STDEVP`)**, CV, GMI, eA1c,
  5-tier Time-in-Range, AUC, MODD, J-Index, M-Value, GRI.
- **`Comparative`** — the headline **Off Meds vs On Meds** comparison.
- **`Daily`**, **`AGP`**, **`Time Blocks`** — day-level / time-of-day rollups.
- **`Glossary`**, **`Study Export`**, **`Review Notes`** — definitions, export
  row, known-limitations log.

Configure the analysis via `INPUT!I1` (metformin start date) and `INPUT!I2`
(washout-buffer days).

### `LingoCGM_Metformin_Study_Analysis.xlsx` (~76 KB, 8 sheets)
A multi-participant study workbook: `Dashboard`, `Study_Setup`, `Source_Files`,
`File_Intake`, `Pasted_Exports`, `Participant_Metrics`, `Cohort_Stats`
(aggregated mean/SD/CV per group), `Safety_QA`.

---

## Working with the files

These are binary `.xlsx` files. Inspect them with Python:

```bash
pip install openpyxl pandas
```

```python
import openpyxl
vals     = openpyxl.load_workbook("LingoCGM_CGM_Report_codex.xlsx", data_only=True)   # cached values
formulas = openpyxl.load_workbook("LingoCGM_CGM_Report_codex.xlsx", data_only=False)  # formulas
```

A cached value of `—` means the formula evaluated to an error (caught by
`IFERROR`), not that the source data is missing. The `_Calc` sheet is the source
of truth — re-derive statistics from `_Calc!B` (glucose) filtered by `_Calc!E`
(`Off Meds` / `On Meds`).

---

## Issues in the standard-deviation file

The standard-deviation analysis lives in **`LingoCGM_CGM_Report_codex.xlsx`**.
The whole-dataset SD on the **`Summary`** sheet is fine
(`STDEVP(_Calc!B…)` = **17.47 mg/dL**, verified against the raw data). **The
problem is on the `Comparative` sheet**, whose entire purpose — showing whether
metformin reduces glucose variability — depends on a per-period (Off Meds vs On
Meds) standard deviation that is **broken**.

Verified by recomputing from `_Calc` (Off Meds n = 765, On Meds n = 2162):

| Comparative metric | Off Meds (shown → correct) | On Meds (shown → correct) |
|---|---|---|
| **Standard Deviation** (B20/C20) | `—` → **15.46** | `—` → **14.85** |
| **Coefficient of Variation** (B21/C21) | `—` → **14.34%** | `—` → **14.51%** |
| **Minimum Glucose** (B13/C13) | 58 ✅ | `—` → **81** |
| **Maximum Glucose** (B14/C14) | 175 ✅ | `—` → **164** |

### 🔴 1. The conditional Standard Deviation returns `—`
`Comparative!B20` and `C20` use the "conditional population SD" workaround the
`Glossary` documents — `SQRT(SUMPRODUCT(…)/COUNTIF(…))` — but in the delivered
file **both cells evaluate to an error** and fall through to `—`. So neither the
Off-Meds nor the On-Meds SD is ever shown, even though the correct values
(15.46 and 14.85 mg/dL) are easily computable from `_Calc`.

### 🔴 2. Coefficient of Variation cascades to broken
`B21`/`C21` are defined as `SD ÷ Mean × 100`, i.e. they reference the broken
`B20`/`C20`. Because SD is `—`, **CV is `—` too**. The CV row — the sheet's
headline "is glucose more stable on metformin?" answer — is therefore blank.

### 🟠 3. On-Meds MIN/MAX use the wrong function prefix
The Off-Meds min/max (`B13`/`B14`) use `_xlfn.MINIFS` / `_xlfn.MAXIFS` and work
(58 / 175). The On-Meds cells (`C13`/`C14`) use **`_xludf.MINIFS` /
`_xludf.MAXIFS`**. The `_xludf.` prefix tells Excel the function is an unknown
user-defined function, so the cell errors and shows `—`. Correct values: min 81,
max 164.

### Root cause
The summary cells were written with **invalid function tokens / a workaround
formula that errors at evaluation time**, rather than being driven by working
formulas over `_Calc`. The same metrics computed on the `Summary` sheet (whole
dataset) are correct — only the per-period `Comparative` versions are broken.

### How to fix
1. **Min/Max:** change `_xludf.MINIFS`/`_xludf.MAXIFS` in `C13`/`C14` to the
   `_xlfn.` prefix used by the working `B13`/`B14`.
2. **SD:** repair the conditional-SD formula in `B20`/`C20` so it evaluates
   (mirroring the documented `SQRT(SUMPRODUCT(...)/COUNTIF(...))` pattern), or
   replace it with an equivalent array formula. Target values: 15.46 / 14.85.
3. **CV** (`B21`/`C21`) needs no change once SD is fixed — it recomputes
   automatically to 14.34% / 14.51%.
