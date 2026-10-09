VBA code for Fastener
=====================

The macros inside FINAL_HARDWARE_LIST_KRISH.xlsm, one file per module,
exported from Excel. This folder is the master copy: it is kept identical
to what the workbook runs. Every time the VBA in the workbook is changed,
these files are re-exported and committed.

  Module1.bas                  Print buttons. PRINTOUTPUT (Ctrl+Shift+P) and
                               PRINTPROTECTEDOUTPUT (Ctrl+Shift+Q). Rebuild the
                               sheet being printed first; skip printing if that fails.

  FinalSheetBuilder.bas        RebuildFinalSheet. Writes FINAL SHEET's purchase
                               totals (nuts/washers by size, then bolts, then studs)
                               from PRINT SHEET.

  ProductionSheetBuilder.bas   RefreshAll (the REFRESH button on PRINT SHEET),
                               RebuildSheetSafely (unlock / rebuild / re-lock), and
                               RebuildProductionSheet (one block per part, page
                               breaks, print setup).

  VerificationBuilder.bas      RunVerification (the VERIFY button). Checks every
                               joint on TANK HARDWARE and STD DATA against the
                               standard table on STD.DATA, and cross-checks totals.

The sheet and ThisWorkbook code modules are empty, so they are not exported.

To change a macro
-----------------
1. Edit the .bas file here.
2. Put it into the workbook (Excel closed):
     python .claude/route/update_module.py "Document/New Feauture/FINAL_HARDWARE_LIST_KRISH.xlsm" <ModuleName>
   This imports the whole module and exports it straight back here, so the
   file matches Excel exactly.
3. Run the tests in .claude/route (see the reference PDF in "Document/About this Excel").
4. Commit this folder together with the workbook.

Never edit macros by deleting single lines from code (it once wiped a whole
module). Import the whole .bas instead, as above.
