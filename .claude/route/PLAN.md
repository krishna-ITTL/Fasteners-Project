# PLAN — PRODUCTION SHEET + single REFRESH on PRINT SHEET (2026-10-09)

## Goal
Design team wants a shopfloor-friendly **PRODUCTION SHEET**: one block per fastener
part, listing every joint location that uses it, with a total. One REFRESH button on
PRINT SHEET rebuilds both FINAL SHEET and PRODUCTION SHEET. The macros must be heavily
commented (the user will teach macros from them).

## Files
- Source (designer's file, currently OPEN in Excel — do not modify):
  `Document/New Feauture/6367_HARDWARE LIST.xlsm`
- Output: `Document/New Feauture/6367_HARDWARE LIST_PRODUCTION.xlsm` (a copy, edited).
- Working VBA sources kept in `.claude/route/`: `ProductionSheetBuilder.bas`,
  `Module1.bas` (updated), `RefreshAll` lives in ProductionSheetBuilder.

## What the workbook looks like now (facts verified 2026-10-09)
- Inputs are split over two sheets: `STD DATA` (standard joints: valves, plates, flange
  joints) and `TANK HARDWARE` (tank-specific joints). `PRINT SHEET` pulls rows from BOTH,
  interleaved, rows 9..~320. Not a fixed row offset any more.
- PRINT SHEET columns: B flag (1 if any qty), D joint location (merged D:E, and for
  continuation rows merged vertically, e.g. D101:E105 — read via `MergeArea.Cells(1,1)`;
  verified every non-zero row resolves to a location this way), F designation,
  G bolt/stud qty, H nut, I p.washer, J s.washer, K l.nut, L STD.REF.
  Section/heading rows have B = "-" and no F.
- Designations: `M12x50-HDG`, `M12-STUDx75-HDG` (no grade now). Parse with the same
  regex FinalSheetBuilder uses: `^\s*(M\d+)\s*(-STUD)?\s*x\s*(\d+)\s*-\s*(.+?)\s*$`.
- FinalSheetBuilder is unchanged and measures PRINT SHEET's last row itself — keep it.
- FINAL SHEET has a Forms button `btnRefreshFinalSheet` (at L10) → `RebuildFinalSheet`.
  PRINT SHEET has no button; its print area is `D2:L324`.
- PRODUCTION SHEET template: header rows 2–6 (logo picture at D2, title F2:H4,
  FORMAT/VERSION/DATE I2:K4, WORK ORDER/CUSTOMER/REVISION/RATING formulas rows 5–6 —
  keep these exactly). Example blocks rows 7–20 (to be replaced):
  - title row: `D:K` merged, value e.g. `M4X30-HDG`, font 16, yellow fill FFFF00
  - header row: D:E merged `JOINT LOCATIOIN`, F `BOLT`, G `NUT `, H `P. WASHER`,
    I `S. WASHER`, J `L.NUT`, K `STD.REF` (bold, centred)
  - data rows: D:E merged location (left), F..K centred
  - TOTAL row: D:E `TOTAL` (font 12), F..J = SUM of the block's data rows
  - print area `D2:K28`, portrait, paper 9 (A4).

## Requirements
1. New module `ProductionSheetBuilder` with `Sub RebuildProductionSheet()`.
   - Reads PRINT SHEET rows 9..last row (measure like FinalSheetBuilder: last used row
     in column F, floor 245, hard max 20000).
   - Skip rows with no designation, and rows where G..K are ALL zero.
   - Group rows by designation **including material and the -STUD marker**
     (key = UCase of designation). Material must stay in the key (an SS line must never
     merge into an HDG line — see project history).
   - One output row per contributing PRINT SHEET row (no merging of locations).
   - Block order: all bolts first, then all studs; within each, size ascending
     numerically (M4 < M5 < … < M24), then length ascending numerically, then material.
   - Block title = designation in upper case (e.g. `M12X50-HDG`, `M12-STUDX75-HDG`).
   - Stud blocks: column F header reads `STUD` instead of `BOLT`.
   - Data rows: D = joint location (MergeArea top-left of PRINT SHEET col D),
     F..J = values of PRINT SHEET G..K, K = STD.REF (col L).
   - TOTAL row F..J = `=SUM(...)` live formula over that block's data rows.
   - Formatting written by the macro (thin borders on the block D:K, merges and fonts
     as in the template). One blank row between blocks.
   - Before writing: clear and un-merge everything from row 7 down (cols D:K, and any
     leftover formatting), so a shorter rebuild leaves no stale rows.
   - Print area set to `D2:K<last row written>`; print titles rows 2–6 repeat on
     every page.
   - Guards (run BEFORE anything is cleared): unreadable designation (error value or
     regex mismatch) → message listing up to 8 rows, nothing changed. Same "silent
     mode" test seam as FinalSheetBuilder (`gLastProblem`-style, separate names).
   - Restore ScreenUpdating / EnableEvents / Calculation on every exit path.
   - If PRODUCTION SHEET is protected: unprotect (no password), rebuild, re-protect
     with the same options PRINTPROTECTEDOUTPUT uses.
2. `Sub RefreshAll()`: calls `RebuildFinalSheet` then `RebuildProductionSheet`; if the
   first reports a problem, still runs the second but tells the user both outcomes once.
   FINAL SHEET protection: unprotect/re-protect around RebuildFinalSheet the same way
   (PRINTPROTECTEDOUTPUT leaves sheets protected, which would otherwise make the next
   refresh fail with 1004).
3. Buttons: delete `btnRefreshFinalSheet` from FINAL SHEET. Add a Forms button
   `btnRefreshAll`, caption `REFRESH`, on PRINT SHEET at column N rows 2–4 (outside the
   print area `D2:L324`), OnAction `RefreshAll`.
4. Module1: PRINTOUTPUT / PRINTPROTECTEDOUTPUT also rebuild PRODUCTION SHEET when it is
   the active sheet (same pattern as the FINAL SHEET line). Must keep the
   `Attribute ... VB_Invoke_Func` lines (Ctrl+Shift+P / Q).
5. Comments: every block of ProductionSheetBuilder explained in plain English for a
   macro beginner (what it does and why), not just what the line says.
6. Explainer artifact (HTML, downloadable) teaching how the macros work: the refresh
   flow, each macro's job, a line-by-line walk-through of RebuildProductionSheet, how to
   open the VBA editor / run / step through (F8) / edit safely, and common pitfalls.

## Non-goals
- No changes to VERIFICATION (hidden, older code — separate job later).
- No changes to STD DATA, TANK HARDWARE, PRINT SHEET formulas, FinalSheetBuilder logic.
- No keep-a-block-on-one-page logic (print titles only). ponytail: add manual page
  breaks if shopfloor complains about split blocks.

## Safety (from project history)
- Edit VBA only via Excel COM `VBComponents.Import` of a complete .bas. NEVER
  `CodeModule.DeleteLines` on a single Sub (wiped Module1 on 2026-08-22).
  `AddFromString` drops Attribute lines — don't use it for Module1.
- After writing, re-extract with `olevba` and list procedures to confirm nothing vanished.
- Work on a copy; never touch the file the designer has open.

## Acceptance criteria
1. olevba on the output lists: Module1 (PRINTOUTPUT, PRINTPROTECTEDOUTPUT with both
   Attribute lines), FinalSheetBuilder (unchanged vs source), VerificationBuilder
   (unchanged), ProductionSheetBuilder (RebuildProductionSheet, RefreshAll).
2. Running RefreshAll via COM on the output: no error, no dialog in silent mode.
3. PRODUCTION SHEET content equals an independent Python recomputation from PRINT SHEET
   cached values (blocks, order, rows, values, totals).
4. Sum of every PRODUCTION SHEET block TOTAL for bolts/studs/nuts/washers equals the
   corresponding FINAL SHEET totals.
5. Header rows 2–6 and the logo are untouched; print area = D2:K<last>.
6. FINAL SHEET has no REFRESH button; PRINT SHEET has `btnRefreshAll` → RefreshAll.
7. Re-running RefreshAll twice gives identical output (idempotent); a test with a
   corrupted designation shows the guard message and leaves the sheet unchanged.
8. Protected PRODUCTION/FINAL SHEET: RefreshAll still works and leaves them protected.
