"""Migrate v19 -> v20 (spec §6).

    python migrate_v19.py <v19.xlsx> [out_dir]

1. Copies v19 INPUT VALUES only (machine sheets A5:H<last> and P; Bucket Status B:C and G; Settings B3:B6).
   No v19 formula column is copied (in particular not Bucket Status D12/E12, the typed-over M2A- cells).
2. Applies the two §6.2 cleanups and nothing else.
3. Builds the Change Log (§3 defects, §6.2 cleanups, §6.3 QA traceability table). The §6.3 table is generated:
   v19 attribution from v19's cached values, v20 attribution from ref_model.py (the oracle).
4. Writes the workbook with build_tracker.py, recalculates it in LibreOffice headless and stores the
   recalculated values into the file as P110_Powder_Tracker_v20.xlsx (cached values come from that real recalc;
   see lo_calc.inject_cached_values for why LibreOffice's own xlsx export is not used).
"""
import datetime
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_tracker as bt  # noqa: E402
from ref_model import run_machine  # noqa: E402

V19_LAST = 104
CLEANUPS = [  # (sheet, cell, old, new, reason)
    ("New P110", "F21", "whitespace", "EMPTY",
     'Whitespace-only Slot 2 Cycle created a phantom key ("<spaces>H-") on rows 21-23 (v19 defect 3.1). The operator intent '
     "was to empty slot 2 (M3H- removed); v20 uses the EMPTY token. G21 stays blank."),
    ("Old P110", "H23", "whitespace", None,
     "Whitespace-only usage on a row with no Build # (v19 defect 3.1). Cleared."),
]

DEFECTS = [
    ("3.1", "Whitespace accepted: New P110!F21 = \"   \" created the phantom key \"   H-\" on rows 21-23; Old P110!H23 = \" \".",
     "Inputs normalized (UPPER/TRIM/CLEAN/NBSP); whitespace-only = ERROR. EMPTY token for 'slot emptied'."),
    ("3.2", "Formulas overwritten: Bucket Status D12/E12 (M2A-) were typed constants 5.09 / 4.78; the bucket showed 4.78 kg OK "
            "when it is 0.000 Empty.", "All formula cells locked and sheets protected; formula-integrity checks on Checks."),
    ("3.3", "A row with no Build # wiped the loaded buckets, breaking swaps and funnel carry-over.",
     "Slot and funnel state carry through blank rows."),
    ("3.4", "Dashboard listed LOT#s for a pending build (\"\">0 is TRUE in Excel): Old B-2792(3rd) showed M4C/M4B as used.",
     "Draw tests use ISNUMBER/N(); a pending build shows '—'."),
    ("3.5", "All dropdowns had showErrorMessage = False, so they blocked nothing.",
     "Stop-style validation from Lists. Checks sheet is the real safety net (pasting replaces validation)."),
    ("3.6", "Only 100 build rows (5-104).", "1,000 rows (5-1004); WARN at 900 rows used."),
    ("3.7", "Slot-dependent results: slot 1 always drained first; funnel had two slot-bound compartments.",
     "One shared FIFO funnel + longest-loaded-first, ties by bucket key. Tested slot-symmetric (TEST_REPORT.md T6)."),
    ("3.8", "Funnel could be over-filled (F1 capped at full capacity while F2 still held leftover).",
     "One queue, one capacity. Overflow = WARNING and stays on the bucket record."),
    ("3.9", "Funnel leftover could be silently overwritten by a second swap into the same compartment.",
     "Queue holds up to 4 lots; nothing is overwritten; a 5th lot = ERROR."),
    ("3.10", "Column J 'Total LEFT' left out powder still in the funnel.", "J = slot 1 + slot 2 + funnel."),
    ("3.11", "Same bucket on both machines not detected.", "ERROR (row R + Checks)."),
    ("3.12", "TEXT(C,\"yyyymmdd\") locale-dependent; float residue could show Low instead of Empty; text usage silently "
             "excluded from totals.", "Locale-safe LOT#; ROUND only at display/status; non-numeric usage = ERROR."),
    ("3.13", "Dashboard formulas repeated LOOKUP(2,1/…) up to 20× per cell.", "One last-row helper per machine + INDEX."),
]


def read_v19(path):
    wb = openpyxl.load_workbook(path)                    # formulas (we only take input columns)
    wv = openpyxl.load_workbook(path, data_only=True)    # v19 cached values (for the §6.3 comparison)
    data = {"machines": {}, "buckets": {}, "settings": {}}
    for m in bt.MACHINES:
        ws = wb[m]
        last = max([r for r in range(5, V19_LAST + 1)
                    if any(ws.cell(r, c).value is not None for c in list(range(1, 9)) + [16])] or [4])
        data["machines"][m] = [{c: ws["%s%d" % (c, r)].value for c in "ABCDEFGHP"} for r in range(5, last + 1)]
    bs = wb["Bucket Status"]
    for r in range(4, 76):
        data["buckets"][bs["A%d" % r].value] = {c: bs["%s%d" % (c, r)].value for c in "BCG"}
    for c in ("B3", "B4", "B5", "B6"):
        data["settings"][c] = wb["Settings"][c].value
    return data, wv


def apply_cleanups(data):
    log = []
    for sheet, cell, old, new, reason in CLEANUPS:
        col, r = cell[0], int(cell[1:])
        row = data["machines"][sheet][r - 5]
        found = row[col]
        assert isinstance(found, str) and found != "" and found.strip() == "", \
            "%s!%s expected a whitespace-only value, found %r" % (sheet, cell, found)
        row[col] = new
        log.append((sheet + "!" + cell, "%r (%d space%s)" % (found, len(found), "" if len(found) == 1 else "s"),
                    "blank" if new is None else new, reason))
    return log


def oracle_rows(rows):
    return [dict(build=r.get("A"), date=r.get("B"), clean=r.get("C"), s1c=r.get("D"), s1b=r.get("E"),
                 s2c=r.get("F"), s2b=r.get("G"), usage=r.get("H")) for r in rows]


def full_weights(data):
    return {k: v["B"] for k, v in data["buckets"].items() if isinstance(v.get("B"), (int, float))}


def lot_of(data, key):
    d = data["buckets"].get(key, {}).get("C")
    return key + d.strftime("%Y%m%d") if isinstance(d, (datetime.date, datetime.datetime)) else key


def run_oracle(data):
    s = data["settings"]
    out = {}
    for m in bt.MACHINES:
        out[m] = run_machine(oracle_rows(data["machines"][m]), full_weights(data), s["B6"], s["B4"] + s["B5"])
    return out


def consumption_order(o_out, draws, i):
    """LOT order as column Q shows it: funnel lots FIFO, then longest-loaded bucket, then the other."""
    f = [b for ii, b, a, src in draws if ii == i and src == "funnel" and a > 1e-9]
    slots = [(src, b) for ii, b, a, src in draws if ii == i and src != "funnel" and a > 1e-9]
    return f, slots


def v19_attribution(wv, sheet, r):
    """v19 draw order: funnel compartment 1 (Y/AA), compartment 2 (AC/AE), slot 1 (U/AK), slot 2 (V/AL)."""
    ws = wv[sheet]
    out = []
    for kc, qc in (("Y", "AA"), ("AC", "AE"), ("U", "AK"), ("V", "AL")):
        q = ws["%s%d" % (qc, r)].value
        if isinstance(q, (int, float)) and q > 1e-9:
            out.append(ws["%s%d" % (kc, r)].value)
    return out


def qa_table(data, wv, oracle, q_values=None):
    """Builds where v19 and v20 attribute the build's powder to different LOT#s.
    Returns rows (sheet, build, v19 LOTs, v20 LOTs in oracle draw order)."""
    rows = []
    for m in bt.MACHINES:
        out, draws, _ = oracle[m]
        for i, o in enumerate(out):
            if o["total"] is None:
                continue
            r = 5 + i
            v19 = v19_attribution(wv, m, r)
            v20 = [b for ii, b, a, _ in draws if ii == i and a > 1e-9]
            if set(v19) != set(v20):
                rows.append((m, data["machines"][m][i]["A"], r,
                             ", ".join(lot_of(data, k) for k in v19), ", ".join(lot_of(data, k) for k in v20)))
    return rows


def change_log(cleanup_log, qa):
    today = datetime.date.today().isoformat()
    return [
        ("h", "Revision record"),
        ("t", ["Date", "Version", "Change", "Author", "Approver (QA)"],
         [[today, "v20", "Rebuilt from v19: slot-independent powder model, hardened inputs, Checks sheet. Tested per "
                         "TEST_REPORT.md; not validated - formal software validation (ISO 13485 §4.1.6) is for QA to approve.",
           "________", "________"]]),
        ("sp",),
        ("h", "v19 defects fixed in v20"),
        ("t", ["#", "v19 defect (evidence)", "Fix in v20", "", ""], [[a, b, c, "", ""] for a, b, c in DEFECTS]),
        ("sp",),
        ("h", "Migration cleanups (the only data changes made while copying v19 inputs)"),
        ("t", ["Cell", "Old value", "New value", "Reason", "Approved by / date"],
         [[c, o, n, why, "________"] for c, o, n, why in cleanup_log]),
        ("sp",),
        ("h", "QA REVIEW REQUIRED - traceability changes (LOT#s recorded against these builds change)"),
        ("p", "v19 always drained slot 1 first; v20 drains the funnel (oldest leftover first), then the bucket that has "
              "been on the machine longest, then the other (ties by bucket key). Build totals are unchanged on every row; "
              "only which bucket supplied which build changes while two buckets were loaded. LOT#s are listed in the "
              "oracle's draw-record order (funnel, slot 1, slot 2); column Q on the machine sheets lists the same LOT#s "
              "in consumption order."),
        ("t", ["Sheet / row", "Build #", "LOT#s per v19", "LOT#s per v20", "QA disposition / signature"],
         [["%s row %d" % (m, r), b, a, c, "________"] for m, b, r, a, c in qa]),
    ]


def main(v19_path, out_dir=HERE):
    from lo_calc import LO, inject_cached_values, read_all
    data, wv = read_v19(v19_path)
    cleanup_log = apply_cleanups(data)
    oracle = run_oracle(data)
    qa = qa_table(data, wv, oracle)
    wb = bt.build_workbook(data, 1000, change_log(cleanup_log, qa))
    raw = os.path.join(out_dir, "_v20_unrecalculated.xlsx")
    final = os.path.join(out_dir, "P110_Powder_Tracker_v20.xlsx")
    wb.save(raw)
    with LO() as lo:
        doc = lo.open(raw)
        errs = lo.error_cells(doc)
        banner = lo.read(doc, "Checks", "A3:A3")[0][0]
        arrays = read_all(doc)
        doc.close(True)
    if errs:
        raise SystemExit("formula errors after recalc: %s" % errs)
    inject_cached_values(raw, final, arrays)
    os.remove(raw)
    print("wrote", final)
    print("Checks banner:", banner, "| error cells:", errs or "none")
    for row in qa:
        print("QA:", row)
    return final


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else HERE)
