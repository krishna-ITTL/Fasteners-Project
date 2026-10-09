"""VERIFICATION tests on a scratch copy (never saved).

Usage: python verification_tests.py <scratch copy.xlsm>
"""
import os
import sys
import win32com.client

p = os.path.abspath(sys.argv[1])
fails = 0


def check(name, ok, extra=""):
    global fails
    print(("PASS " if ok else "FAIL ") + name + (f"  [{extra}]" if extra else ""))
    fails += 0 if ok else 1


xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(p)
    vs, sdi, th, sd = (wb.Sheets(n) for n in ("VERIFICATION", "STD DATA", "TANK HARDWARE", "STD.DATA"))
    xl.Run("SetVerifySilent", True)

    def run():
        xl.Run("RunVerification")
        rows = []
        for r in range(34, 184):
            if vs.Rows(r).Hidden:
                break
            v = [vs.Cells(r, c).Value for c in range(2, 13)]
            if any(x is not None for x in v):
                rows.append(v)
        counts = [vs.Cells(r, 3).Value for r in (27, 28, 29, 30)]
        return xl.Run("VerifyLastProblem"), rows, counts

    def has(rows, ref=None, item=None, sev=None, text=None):
        return [f for f in rows if (ref is None or str(f[0] or "").startswith(ref))
                and (item is None or f[4] == item) and (sev is None or f[10] == sev)
                and (text is None or text in str(f[9]))]

    prob, base, cnt = run()
    check("runs without problem", prob == "", prob)
    check("sheet is visible", vs.Visible == -1)
    check("summary counts add up", sum(cnt[:3]) == sum(
        (int(str(f[1]).split()[0]) if str(f[1]).endswith(" rows") else 1) for f in base), f"{cnt}")
    check("SD-267 / SD-274 recognised (not UNVERIFIABLE)",
          not [f for f in base if f[2] in ("SD-267", "SD-274") and f[10] == "UNVERIFIABLE"])
    check("STD DATA rows are checked (findings name SD rows)", bool(has(base, ref="SD ")), len(has(base, ref="SD ")))
    check("unused (all-zero) joints are not reported",
          not has(base, text="quantity is 0") or all(any(x not in (None, 0) for x in f[6:9]) for f in has(base, text="quantity is 0")))
    check("totals agree: inputs = PRINT SHEET = FINAL (TH 59/60 now linked)",
          not has(base, item="TOTAL"), [f[9] for f in has(base, item="TOTAL")])

    # a new deviation on STD DATA is caught, with its row
    # 15NB gate valve (row 15, SD-268): nut = bolt -> make it 2 per bolt
    sdi.Range("O15").Formula = "=N15*2"
    _, rows, _ = run()
    check("STD DATA deviation caught as 'SD 15'", bool(has(rows, ref="SD 15", item="NUT", sev="ERROR")),
          [f[0] for f in has(rows, item="NUT")])
    # a ruling accepts it
    vs.Range("B14").Value, vs.Range("C14").Value, vs.Range("D14").Value = "SD-268", "NUT", 2
    _, rows, cnt2 = run()
    check("ruling silences it", not has(rows, ref="SD 15", item="NUT") and cnt2[3] == 1, cnt2)
    vs.Range("B14:D14").ClearContents()
    sdi.Range("O15").Formula = "=N15"

    # an input row that is NOT on PRINT SHEET is caught by the total check
    n59 = th.Range("N59").Formula
    ps = wb.Sheets("PRINT SHEET")
    r59 = next(r for r in range(120, 140) if ps.Cells(r, 7).Formula == "='TANK HARDWARE'!N59")
    ps.Cells(r59, 7).Value = 0                        # unlink the bolt qty
    _, rows, _ = run()
    check("unlinked input row is reported by the total check",
          bool(has(rows, item="TOTAL", text="2030 vs 2026")), [f[9] for f in has(rows, item="TOTAL")])
    ps.Cells(r59, 7).Formula = "='TANK HARDWARE'!N59"

    # standard table moved by an inserted row above it: same result (no drift)
    sd.Rows(100).Insert()
    _, rows, cnt3 = run()
    check("standard table still found after a row is inserted above it",
          cnt3[2] == cnt[2] and not [f for f in rows if f[2] in ("SD-267", "SD-274") and f[10] == "UNVERIFIABLE"], f"{cnt3} vs {cnt}")
    sd.Rows(100).Delete()

    # hardware typed below the STD DATA range is refused, not silently skipped
    sdi.Range("M160").Value = "M12x50-HDG"
    prob, _, _ = run()
    check("STD DATA range guard", "STD DATA" in prob and "below row 150" in prob, prob[:90])
    sdi.Range("M160").ClearContents()

    # an error value in a part name is reported, not a crash
    m20 = sdi.Range("M20").Formula
    sdi.Range("M20").Formula = "=NA()"
    prob, rows, _ = run()
    check("error-value part name reported, run completes",
          prob == "" and bool(has(rows, ref="SD 20", text="error value")), prob[:80])
    sdi.Range("M20").Formula = m20

    # filtered PRINT SHEET: a hidden non-zero row still counts in the total
    th.Range("C243").Value = 2                       # foundation bolt, row hidden by the filter
    _, rows, _ = run()
    # inputs vs PRINT SHEET: the hidden row is counted on both sides, so they agree;
    # PRINT SHEET vs FINAL SHEET: FINAL was not refreshed, so it is flagged as stale.
    check("total check counts filter-hidden PRINT SHEET rows (and flags stale FINAL)",
          [f[9] for f in has(rows, item="TOTAL")] == ["totals disagree: 2032 vs 2030"],
          [f[9] for f in has(rows, item="TOTAL")])
    th.Range("C243").Value = 0
    wb.Close(False)
finally:
    xl.Quit()
print("FAILS:", fails)
sys.exit(1 if fails else 0)
