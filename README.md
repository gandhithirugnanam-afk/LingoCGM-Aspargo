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

> **Status: fixed.** The bugs below were repaired in
> `LingoCGM_CGM_Report_codex.xlsx` (see `Standard_Deviation_Analysis.docx` for the
> full write-up). The values in the "shown → correct" column are the original
> broken state and the corrected value now in the file.

The standard-deviation analysis lives in **`LingoCGM_CGM_Report_codex.xlsx`**.
The whole-dataset SD on the **`Summary`** sheet is fine
(`STDEVP(_Calc!B…)` = **17.47 mg/dL**, verified against the raw data). **The
problem was on the `Comparative` sheet**, whose entire purpose — showing whether
metformin reduces glucose variability — depends on a per-period (Off Meds vs On
Meds) standard deviation that was **broken**.

Verified by recomputing from `_Calc` (Off Meds n = 765, On Meds n = 2162):

| Comparative metric | Off Meds (shown → correct) | On Meds (shown → correct) |
|---|---|---|
| **Standard Deviation** (B20/C20) | `—` → **17.82** | `—` → **20.96** |
| **Coefficient of Variation** (B21/C21) | `—` → **16.53%** | `—` → **20.48%** |
| **Minimum Glucose** (B13/C13) | 58 ✅ | `—` → **55** |
| **Maximum Glucose** (B14/C14) | 175 ✅ | `—` → **182** |

### 🔴 1. The conditional Standard Deviation returns `—`
`Comparative!B20` and `C20` use the "conditional population SD" workaround the
`Glossary` documents — `SQRT(SUMPRODUCT(…)/COUNTIF(…))` — but in the delivered
file **both cells were saved with an error value** and fell through to `—`. So
neither the Off-Meds nor the On-Meds SD was shown, even though the correct values
(17.82 and 20.96 mg/dL) are easily computable from `_Calc`.

### 🔴 2. Coefficient of Variation cascades to broken
`B21`/`C21` are defined as `SD ÷ Mean × 100`, i.e. they reference the broken
`B20`/`C20`. Because SD is `—`, **CV is `—` too**. The CV row — the sheet's
headline "is glucose more stable on metformin?" answer — is therefore blank.

### 🟠 3. On-Meds MIN/MAX use the wrong function prefix
The Off-Meds min/max (`B13`/`B14`) use `_xlfn.MINIFS` / `_xlfn.MAXIFS` and work
(58 / 175). The On-Meds cells (`C13`/`C14`) use **`_xludf.MINIFS` /
`_xludf.MAXIFS`**. The `_xludf.` prefix tells Excel the function is an unknown
user-defined function, so the cell errors and shows `—`. Correct values: min 55,
max 182.

### Root cause
The summary cells were written with **invalid function tokens / a workaround
formula that errors at evaluation time**, rather than being driven by working
formulas over `_Calc`. The same metrics computed on the `Summary` sheet (whole
dataset) are correct — only the per-period `Comparative` versions are broken.

### How it was fixed
1. **Min/Max:** changed `_xludf.MINIFS`/`_xludf.MAXIFS` in `C13`/`C14` to the
   `_xlfn.` prefix used by the working `B13`/`B14` → now 55 / 182.
2. **SD:** kept the documented `SQRT(SUMPRODUCT(...)/COUNTIF(...))` formula in
   `B20`/`C20` (it is valid in Excel) and refreshed the saved cached values →
   17.82 / 20.96. The workbook already has `forceFullCalc="1"`, so Excel
   recomputes these on open.
3. **CV** (`B21`/`C21`) needed no change once SD is populated — it recomputes
   to 16.53% / 20.48%.

The fix was applied with surgical edits to `xl/worksheets/sheet3.xml` so all
charts, formatting, and other sheets are preserved.

## Follow-up improvements (robustness + statistics)

A second round of edits to `LingoCGM_CGM_Report_codex.xlsx` (sheets `Summary`
and `Comparative`), again via surgical XML edits that preserve all charts.

**Refreshed stale `—` caches (A).** Several formula cells were saved with a
stale error value (`—`) even though their formulas were valid — they only
recovered when reopened in desktop Excel. Re-populated the correct cached values
so the file reads correctly in any viewer (openpyxl, web preview, etc.):
`Summary!B16` (Hyperglycemia Burden Score), the Comparative HBS / SD / CV /
min-max delta cells (`B16`, `C16`, `D13`, `D14`, `D16`, `D20`, `D21`), and the
J-Index row (`B69`/`C69`/`D69`), which was blank only because it depends on the
SD cells. (A few per-day AUC cells — `Summary!B49`, `Comparative!B42`/`B43` —
were left to Excel's `forceFullCalc` recalculation on open.)

**Hardened the formulas (B).** Rewrote the conditional standard-deviation and
HBS formulas (`B16`/`C16`, `B20`/`C20`) to an **IF-free** `SUMPRODUCT` form. The
original `SUMPRODUCT(…*IF(ISNUMBER(x),…)…)` pattern is exactly what the generating
tool kept mis-evaluating into `—`; the IF-free form evaluates cleanly in Excel
**and** LibreOffice/openpyxl. (Verified safe: column B of `_Calc` contains zero
text cells.) Also confirmed no `_xludf.` tokens remain anywhere.

**Added statistical significance (C).** A new block in `Comparative` rows 78–90
tests the mean-glucose difference (On Meds vs Off Meds), Welch two-sample with a
normal approximation (Welch df is in the thousands, so *t* ≈ *z*):

| Quantity | Value |
|---|---|
| Off / On n | 765 / 2162 readings |
| Off / On sample SD (`STDEV.S`) | 17.83 / 20.96 mg/dL |
| Mean difference (On − Off) | **−5.45 mg/dL** |
| Standard error | 0.79 mg/dL |
| z statistic | −6.93 |
| p-value (two-sided) | ≈ 4×10⁻¹² |
| 95% confidence interval | **[−6.99, −3.91] mg/dL** |

The block uses **sample SD** (more defensible than population SD for a two-group
comparison) and carries an explicit **caveat**: CGM readings minutes apart are
autocorrelated, so the effective sample size is far smaller than *n* and this
p-value overstates significance — treat as exploratory, aggregating to daily
means before any formal inference.
