"""Install PRODUCTION SHEET macros into the output workbook via Excel COM.

Usage: python install_production.py <workbook.xlsm>
Imports whole .bas modules (never line-patching), swaps the REFRESH button
from FINAL SHEET to PRINT SHEET, runs RefreshAll once and saves.
"""
import os
import sys
import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
wb_path = os.path.abspath(sys.argv[1])

xl = win32com.client.DispatchEx("Excel.Application")   # own instance, not the user's
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(wb_path)
    comps = wb.VBProject.VBComponents

    for name in ("Module1", "FinalSheetBuilder", "ProductionSheetBuilder", "VerificationBuilder"):
        try:
            comps.Remove(comps.Item(name))
        except Exception:
            pass
        comps.Import(os.path.join(HERE, name + ".bas"))

    # TANK HARDWARE rows 59 and 60 (added in revision R2) were never linked on
    # PRINT SHEET, so 16 bolts were missing from FINAL / PRODUCTION SHEET.
    # Copy PRINT SHEET row 131 (= TANK HARDWARE row 58) into two new rows
    # below it: relative formulas then point at TANK HARDWARE rows 59 and 60,
    # with the same merges, flag formula and filter. Skipped if already linked.
    ps = wb.Sheets("PRINT SHEET")
    linked = {ps.Cells(r, 6).Formula for r in range(9, 400)}
    if "='TANK HARDWARE'!M59" not in linked and "='TANK HARDWARE'!M60" not in linked:
        assert ps.Range("F131").Formula == "='TANK HARDWARE'!M58", ps.Range("F131").Formula
        ps.Rows("131:131").Copy()
        ps.Rows("132:133").Insert(Shift=-4121)          # xlDown
        xl.CutCopyMode = False
        print("PRINT SHEET: linked TANK HARDWARE rows 59, 60 at rows 132, 133")

    # Button: off FINAL SHEET, onto PRINT SHEET (column N, rows 2-4, outside print area).
    fs = wb.Sheets("FINAL SHEET")
    for shp in list(fs.Shapes):
        if shp.Name == "btnRefreshFinalSheet":
            shp.Delete()
    ps = wb.Sheets("PRINT SHEET")
    for shp in list(ps.Shapes):
        if shp.Name == "btnRefreshAll":
            shp.Delete()
    cell = ps.Range("N2:N4")
    btn = ps.Buttons().Add(cell.Left + 4, cell.Top, 90, cell.Height)
    btn.Name = "btnRefreshAll"
    btn.Caption = "REFRESH"
    btn.OnAction = "RefreshAll"
    btn.Font.Bold = True

    # Small "Page 1 of 10" at the bottom centre of every printed page
    # (&8 = 8pt font, &P = this page, &N = total pages).
    for name in ("PRINT SHEET", "FINAL SHEET", "PRODUCTION SHEET"):
        wb.Sheets(name).PageSetup.CenterFooter = "&8Page &P of &N"

    # VERIFICATION now covers STD DATA too; unhide it and say so.
    vs = wb.Sheets("VERIFICATION")
    vs.Visible = -1                                  # xlSheetVisible
    vs.Range("B3").Value = "Checks TANK HARDWARE and STD DATA against the joint standard in STD.DATA. Report only - nothing is changed."
    vs.Range("B33").Value = "ROW (TH/SD)"

    xl.Run("SetSilentMode", True)
    xl.Run("SetProductionSilent", True)
    xl.Run("SetVerifySilent", True)
    xl.Run("RunVerification")
    print("verify problem:", repr(xl.Run("VerifyLastProblem")), "| result:", vs.Range("C10").Value,
          "| err/info/unv:", [vs.Cells(r, 3).Value for r in (27, 28, 29)])
    xl.Run("SetVerifySilent", False)
    xl.Run("RefreshAll")
    print("final problem:", repr(xl.Run("LastProblem")))
    print("production problem:", repr(xl.Run("ProductionLastProblem")))
    xl.Run("SetSilentMode", False)
    xl.Run("SetProductionSilent", False)

    wb.Sheets("PRODUCTION SHEET").Activate()
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
