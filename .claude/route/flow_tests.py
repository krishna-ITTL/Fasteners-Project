"""End-to-end flow test: designer input -> PRINT SHEET -> REFRESH -> FINAL + PRODUCTION.

Usage: python flow_tests.py <scratch copy.xlsm> <scratch folder>
Makes designer-style edits on the copy (never the real file). After each edit:
  * PRINT SHEET must follow the input straight away (formulas),
  * PRODUCTION SHEET must NOT change until REFRESH (it is macro-built),
  * after REFRESH, verify_production.py must pass on a saved snapshot
    (PRODUCTION == recomputed from PRINT SHEET, totals == FINAL SHEET),
  * plus a check specific to the edit.
"""
import os
import re
import subprocess
import sys
import win32com.client

src, outdir = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
HERE = os.path.dirname(os.path.abspath(__file__))
fails = 0


def check(name, ok, extra=""):
    global fails
    print(("PASS " if ok else "FAIL ") + name + (f"  [{extra}]" if extra else ""))
    fails += 0 if ok else 1


xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(src)
    th, sd = wb.Sheets("TANK HARDWARE"), wb.Sheets("STD DATA")
    ps, fs, pd = wb.Sheets("PRINT SHEET"), wb.Sheets("FINAL SHEET"), wb.Sheets("PRODUCTION SHEET")
    xl.Run("SetSilentMode", True)
    xl.Run("SetProductionSilent", True)
    xl.Run("RefreshAll")

    def block(title):
        """Rows of a PRODUCTION block as {location: bolt qty}, plus its TOTAL bolt."""
        col = pd.Range("D1:D600").Value
        for r, (v,) in enumerate(col, 1):
            if v == title:
                rows, rr = {}, r + 2
                while pd.Cells(rr, 4).Value != "TOTAL":
                    rows[pd.Cells(rr, 4).Value] = pd.Cells(rr, 6).Value
                    rr += 1
                return rows, pd.Cells(rr, 6).Value
        return None, None

    def final_qty(size_len, material="HDG"):
        """FINAL SHEET bolt/stud quantity for e.g. 'M24X100' (blocks 2 and 3)."""
        for r in range(37, 126):
            if fs.Cells(r, 4).Value == size_len and fs.Cells(r, 6).Value == material:
                return fs.Cells(r, 7).Value
        return None

    def prod_snapshot():
        return pd.Range("D7:K600").Value

    def refresh_and_verify(tag):
        xl.Run("RefreshAll")
        check(f"{tag}: REFRESH reported no problem", xl.Run("ProductionLastProblem") == "",
              xl.Run("ProductionLastProblem")[:120])
        snap = os.path.join(outdir, f"flow_{tag}.xlsm")
        wb.SaveCopyAs(snap)
        r = subprocess.run([sys.executable, os.path.join(HERE, "verify_production.py"), snap],
                           capture_output=True, text=True)
        last = r.stdout.strip().splitlines()[-1] if r.stdout.strip() else r.stderr[-200:]
        check(f"{tag}: PRODUCTION == PRINT SHEET and totals == FINAL", r.returncode == 0, last)

    # --- 1. change a quantity: CURB BOLT ASM 145 -> 150 (TANK HARDWARE C16, M24x100)
    before_prod = prod_snapshot()
    b_rows, b_tot = block("M24X100-HDG")
    th.Range("C16").Value = 150
    check("1 qty: PRINT SHEET follows at once", ps.Range("G31").Value == 150, ps.Range("G31").Value)
    check("1 qty: PRODUCTION unchanged before REFRESH", prod_snapshot() == before_prod)
    refresh_and_verify("1_qty")
    rows, tot = block("M24X100-HDG")
    check("1 qty: M24X100 block 145 -> 150", rows.get("CURB BOLT ASM") == 150 and tot == b_tot + 5, f"{rows} total {tot}")
    check("1 qty: FINAL M24X100 = 150", final_qty("M24X100") == 150, final_qty("M24X100"))

    # --- 2. switch a joint on: LIGHTNING ARRESTER SUPPORT 0 -> 3 (C55, M16x50)
    _, tot_before = block("M16X50-HDG")
    th.Range("C55").Value = 3
    refresh_and_verify("2_on")
    rows, tot = block("M16X50-HDG")
    added = th.Range("N55").Value
    check("2 on: new location appears in M16X50 block",
          rows.get("LIGHTNING ARRESTER SUPPORT") == added and tot == tot_before + added, f"qty {added}, total {tot_before}->{tot}")

    # --- 3. switch a joint off: CABLE TRAY MOUNTING 58 -> 0 (C58, M8x40)
    th.Range("C58").Value = 0
    refresh_and_verify("3_off")
    rows, tot = block("M8X40-HDG")
    check("3 off: location removed from M8X40 block", rows is not None and "CABLE TRAY MOUNTING" not in rows, f"{rows}")

    # --- 4. a size this job never had: FOUNDATION BOLT M24x300 (C242 -> 4)
    check("4 new size: no M24X320 block yet", block("M24X320-HDG")[0] is None)
    th.Range("C242").Value = 4
    refresh_and_verify("4_new")
    rows, tot = block("M24X320-HDG")
    check("4 new size: M24X320 block created", rows is not None and tot == th.Range("N242").Value, f"{rows}")
    r315 = next(r for r in range(300, 330) if ps.Cells(r, 6).Formula == "='TANK HARDWARE'!M242")
    check("4 new size: row now visible on PRINT SHEET", not ps.Rows(r315).Hidden, f"row {r315}")
    check("4 new size: FINAL SHEET row added", final_qty("M24X320") == th.Range("N242").Value, final_qty("M24X320"))

    # --- 4b. same, but rebuilt the way the print macros do it (no REFRESH,
    #         so PRINT SHEET's filter still hides the row): must not be missed
    r243 = next(r for r in range(300, 330) if ps.Cells(r, 6).Formula == "='TANK HARDWARE'!M243")
    th.Range("C243").Value = 2
    check("4b hidden: row still hidden on PRINT SHEET (filter not re-run)", bool(ps.Rows(r243).Hidden))
    ok_p = xl.Run("RebuildSheetSafely", "PRODUCTION SHEET", "RebuildProductionSheet")
    ok_f = xl.Run("RebuildSheetSafely", "FINAL SHEET", "RebuildFinalSheet")
    rows, tot = block("M20X270-HDG")
    check("4b hidden: PRODUCTION picks up the hidden row", ok_p and rows is not None and tot == th.Range("N243").Value, f"{rows}")
    check("4b hidden: FINAL picks up the hidden row", ok_f and final_qty("M20X270") == th.Range("N243").Value, final_qty("M20X270"))
    refresh_and_verify("4b_hidden")

    # --- 5. customer spec: hardware up to M12 in stainless (STD DATA C4 HDG -> SS)
    sd.Range("C4").Value = "SS"
    check("5 spec: TANK HARDWARE copies the spec", th.Range("C4").Value == "SS")
    desigs = [v for (v,) in ps.Range("F9:F330").Value if v]
    small_ss = [d for d in desigs if d.endswith("-SS")]
    check("5 spec: PRINT SHEET M12-and-below now -SS",
          small_ss and all(int(d[1:].split("x")[0].split("-")[0]) <= 12 for d in small_ss)
          and not any(d.endswith("-HDG") and int(d[1:].split("x")[0].split("-")[0]) <= 12 for d in desigs),
          f"{len(small_ss)} SS rows")
    refresh_and_verify("5_spec")
    titles = [v for (v,) in pd.Range("D7:D600").Value
              if isinstance(v, str) and re.match(r"^M\d+(-STUD)?X\d+-", v)]   # block titles only
    check("5 spec: PRODUCTION blocks split by material",
          any(t.endswith("-SS") for t in titles) and any(t.endswith("-HDG") for t in titles)
          and all(t.endswith("-SS") == (int(t[1:].split("X")[0].split("-")[0]) <= 12) for t in titles), f"{titles[:6]}...")
    check("5 spec: FINAL SHEET has SS lines", final_qty("M12X50", "SS") not in (None, 0), final_qty("M12X50", "SS"))

    wb.Close(False)
finally:
    xl.Quit()
print("FAILS:", fails)
sys.exit(1 if fails else 0)
