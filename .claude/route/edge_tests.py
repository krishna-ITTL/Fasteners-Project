"""Edge tests on a scratch copy: idempotent, bad-designation guard, protection.
Usage: python edge_tests.py <scratch copy.xlsm>   (modifies the copy, never saves)"""
import os, sys, win32com.client
p = os.path.abspath(sys.argv[1])
xl = win32com.client.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
def snap(ws):
    return ws.Range("B1:L400").Formula
fails = 0
def check(name, ok):
    global fails
    print(("PASS " if ok else "FAIL ") + name); fails += 0 if ok else 1
try:
    wb = xl.Workbooks.Open(p)
    pd, ps, fs = wb.Sheets("PRODUCTION SHEET"), wb.Sheets("PRINT SHEET"), wb.Sheets("FINAL SHEET")
    xl.Run("SetSilentMode", True); xl.Run("SetProductionSilent", True)
    hdr0 = pd.Range("D2:K6").Formula
    a = snap(pd); xl.Run("RefreshAll"); b = snap(pd); xl.Run("RefreshAll"); c = snap(pd)
    check("idempotent", a == b == c)
    check("header rows 2-6 untouched", pd.Range("D2:K6").Formula == hdr0)
    check("logo still there", pd.Shapes.Count >= 1)
    check("button on PRINT SHEET -> RefreshAll", ps.Shapes("btnRefreshAll").OnAction.endswith("RefreshAll"))
    check("no refresh button on FINAL SHEET", all(s.Name != "btnRefreshFinalSheet" for s in fs.Shapes))
    # protected sheets still refresh and stay protected
    pd.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    fs.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    xl.Run("RefreshAll")
    check("protected: no problem reported", xl.Run("ProductionLastProblem") == "" and xl.Run("LastProblem") == "")
    check("protected: still protected after", pd.ProtectContents and fs.ProtectContents)
    check("protected: same output", snap(pd) == a)
    pd.Unprotect(); fs.Unprotect()
    # bad designation: guard fires, sheet unchanged
    r = next(r for r in range(10, 330) if ps.Cells(r, 7).Value not in (None, 0) and ps.Cells(r, 6).Value)
    ps.Cells(r, 6).Formula = "BROKEN-PART"
    xl.Run("RebuildProductionSheet")
    msg = xl.Run("ProductionLastProblem")
    check("bad designation reported (row %d)" % r, "BROKEN-PART" in msg and "Nothing has been changed" in msg)
    check("bad designation: sheet unchanged", snap(pd) == a)
    # a new non-zero quantity shows up after refresh
    ps.Cells(r, 6).Formula = "M30x200-SS"
    xl.Run("RebuildProductionSheet")
    check("new part appears as a block", any(str(v[0]) == "M30X200-SS" for v in pd.Range("D7:D400").Value))
    # error value in a part name is caught too
    ps.Cells(r, 6).Formula = "=1/0"
    xl.Run("RebuildProductionSheet")
    check("error-value designation reported", "(error value)" in xl.Run("ProductionLastProblem"))
    # shrinking list: zero all quantities except the first 20 data rows -> no stale rows/merges
    ps.Cells(r, 6).Formula = "M12x50-HDG"
    for rr in range(60, 330):
        if ps.Cells(rr, 6).Value:
            ps.Range(ps.Cells(rr, 7), ps.Cells(rr, 11)).Value = 0
    xl.Run("RebuildProductionSheet")
    used = pd.UsedRange; lastused = used.Row + used.Rows.Count - 1
    pa = pd.PageSetup.PrintArea
    last_total = max(rr for rr in range(7, 400) if pd.Cells(rr, 4).Value == "TOTAL")
    check("shrink: nothing below last TOTAL (used to row %d, total %d)" % (lastused, last_total),
          all(pd.Cells(rr, c).Value is None and not pd.Cells(rr, c).MergeCells
              for rr in range(last_total + 1, 400) for c in range(2, 13)))
    check("shrink: print area ends on last TOTAL (%s)" % pa, pa.endswith("$K$%d" % last_total))
    # Excel settings restored, including manual calculation
    xl.Calculation = -4135   # xlCalculationManual
    xl.Run("RefreshAll")
    check("manual calculation restored", xl.Calculation == -4135)
    check("events/screen restored", xl.EnableEvents and xl.ScreenUpdating in (True, False))
    xl.Calculation = -4105
    # password-protected sheet: refused, reported once, no prompt, still protected
    pd.Protect(Password="pw1")
    xl.Run("RefreshAll")
    msg = xl.Run("ProductionLastProblem")
    check("password sheet reported", "password" in msg and msg.startswith("PRODUCTION SHEET:"))
    check("password sheet still protected", pd.ProtectContents)
    pd.Unprotect("pw1")
    fs.Protect(Password="pw2")
    xl.Run("RefreshAll")
    msg = xl.Run("ProductionLastProblem")
    check("FINAL password reported in combined message", msg.startswith("FINAL SHEET:") and "password" in msg)
    fs.Unprotect("pw2")
    # RebuildSheetSafely returns True on success, False on a problem
    check("RebuildSheetSafely True on success", xl.Run("RebuildSheetSafely", "PRODUCTION SHEET", "RebuildProductionSheet") is True)
    # a failing rebuild must not leave a protected sheet unlocked
    pd.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    ok = xl.Run("RebuildSheetSafely", "PRODUCTION SHEET", "NoSuchMacro")
    check("failed rebuild returns False and re-locks", ok is False and pd.ProtectContents)
    pd.Unprotect()
    wb.Close(False)
finally:
    xl.Quit()
print("FAILS:", fails); sys.exit(1 if fails else 0)
