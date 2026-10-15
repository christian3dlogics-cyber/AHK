# P110 Powder Tracker v20

SLS powder-bucket traceability workbook for the two EOS Formiga P110 machines (Old P110 / New P110).

| File | Purpose |
|---|---|
| `P110_Powder_Tracker_v20.xlsx` | The workbook: v19 data migrated, recalculated, protected (no password), opens on `Checks`. |
| `build_tracker.py` | Generates the workbook. `python3 build_tracker.py` writes the empty template. |
| `migrate_v19.py` | `python3 migrate_v19.py <v19.xlsx>`: copies v19 input values, applies the two logged cleanups, writes the Change Log, recalculates in LibreOffice, writes the v20 file. |
| `ref_model.py` | Reference model (test oracle), verbatim from the build spec §2.4. |
| `lo_calc.py` | LibreOffice headless recalculation over UNO (used by migration and tests). |
| `tests/test_tracker.py` | Tests T1–T11. `python3 tests/test_tracker.py` → writes `TEST_REPORT.md`; exit code 0 only if all pass. |
| `TEST_REPORT.md` | Latest test results, counts and fuzz seed. |

Requirements: Python 3, openpyxl, LibreOffice 7.x+ with Python UNO (`python3-uno`).

Status: tested per `TEST_REPORT.md`. Not validated — formal software validation (ISO 13485 §4.1.6) is for QA to approve.

## Design notes
- Calculation spec = `ref_model.py`. Every machine-sheet row computes, once, in grouped helper columns S…HY:
  normalized inputs → slot state → removal order → 4-lot funnel queue → funnel consumption → slot draws
  (longest-loaded bucket first, ties by bucket key, never by slot) → draw columns → 72-column bucket ledger
  (remaining kg per bucket after the row) → one message column per check → column R.
- Text normalization uses `UNICHAR(160)` for NBSP (`CHAR(160)` is code-page dependent: † on Mac Excel).
  `UNICHAR` and `ISFORMULA` are Excel 2013+ functions, stored with the `_xlfn.` prefix.
- Cached values: LibreOffice 24.2's own xlsx export drops cell styles on later sheets once a workbook holds
  ~400k formula cells, so the delivered file is the openpyxl-generated workbook with the cached results of
  the LibreOffice recalc written into it (`lo_calc.inject_cached_values`). Test T1b checks that every cached
  value equals a fresh recalc. The file is also flagged full-calc-on-load, so Excel recalculates on open.
