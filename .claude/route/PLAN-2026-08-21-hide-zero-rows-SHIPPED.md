# PLAN: FINAL SHEET — show only sizes actually present in PRINT SHEET

## Goal
On `TANK HARDWARE LIST.xlsm`, sheet `FINAL SHEET`: the `SIZE`, `BOLT SIZE`, and `STUD SIZE`
columns currently list a fixed, generic superset of every size/length/coating combination the
template supports. The user wants only the rows that actually have hardware in `PRINT SHEET`
to be visible — e.g. if `PRINT SHEET` has no `M25X80`, that row must not show in `FINAL SHEET`.

## Background (verified against the live file)
- `FINAL SHEET` has three independent blocks, each already `SUMIF`-driven off `PRINT SHEET`
  (rewired in a prior session, still correct — do not touch these formulas):
  - Block 1 — rows 11:19, cols D (SIZE) / G:J (NUT, P.WASHER, S.WASHER, L.NUT)
  - Block 2 — rows 22:93, col D (BOLT SIZE) / col G (BOLT QUANTITY, visually merged G:J)
  - Block 3 — rows 96:125, col D (STUD SIZE) / col G (STUD QUANTITY, visually merged G:J)
  - Row 20 and row 94 are blank separators between blocks; rows 127-131 are a NOTE + sign-off
    footer immediately after block 3 — footer must not shift.
- Every row's D:E and G:J cells are individually merged (per-row merges), so this sheet
  cannot become a native Excel Table (Tables forbid merged cells in the data body) and a single
  classic `AutoFilter` can't span three separate header rows with gaps between them anyway.
- Because each quantity is a `SUMIF` against `PRINT SHEET`, "size not present in PRINT SHEET"
  and "computed quantity = 0" are the same condition here (verified: e.g. M5X70, M8X30, M8X50,
  M10X50, M12X80, M16X90, M16X100, M16X210, M20X100, M20X120, M20X60 are all currently 0 in both
  the HDG and SS halves of Block 2 on the live file — these are exactly the rows that should
  disappear).
- The workbook already solves an equivalent problem elsewhere (`PRINT SHEET` / hidden
  `PRINT QTY`): a flag column `=IF(qty=0,0,1)` + `AutoFilter` on `filter=['1']`, refreshed via
  `ActiveSheet.AutoFilter.ApplyFilter` inside the existing VBA macros `Module1.PRINTOUTPUT` /
  `Module1.PRINTPROTECTEDOUTPUT` (bound to the sheets' "Print" buttons). That exact mechanism
  doesn't transplant cleanly to `FINAL SHEET` (three blocks + merged cells, see above), so the
  equivalent-but-compatible mechanism is real row-hiding via VBA, triggered from the same place.
- No sheet in the workbook currently has `sheetProtection` enabled (`protection.sheet: False`
  everywhere), so hiding rows is not blocked by protection today; do not introduce new
  protection as part of this change.
- Excel's "Trust access to the VBA project object model" is enabled on this machine
  (`AccessVBOM=1` for Office 16.0), so the VBA project can be edited programmatically via
  `win32com.client` (already used by `hardware_generator.py` and available in `.venv`).
  **`openpyxl` cannot edit VBA source** — it only round-trips `vbaProject.bin` as an opaque
  blob — so this change must be made through Excel COM automation, not openpyxl.

## Chosen approach (confirmed with user)
Add a new VBA macro that hides (not deletes) any row in the three blocks whose quantity is
zero, and call it automatically from the **existing Print button macros**, scoped to when
`FINAL SHEET` is the active sheet — no new buttons, no change to how `TANK HARDWARE` or
`PRINT SHEET` print.

## Implementation steps

1. **Backup first**, same convention as the prior session:
   `TANK HARDWARE LIST_backup_<YYYYMMDD_HHMMSS>.xlsm` alongside the original, before any write.

2. **Add a new Sub to `Module1`** (via COM `VBProject.VBComponents("Module1").CodeModule`),
   named `RefreshFinalSheetVisibility`:
   - Reference the `FINAL SHEET` worksheet explicitly (not `ActiveSheet`) so it can be called
     safely regardless of which sheet triggered it.
   - Unhide every row in ranges `11:19`, `22:93`, `96:125` first (reset — a prior job's hidden
     rows must not leak into the next job's output).
   - Block 1 (rows 11-19): hide the row if `SUM(G:J)` for that row = 0 (all four metrics empty).
   - Block 2 (rows 22-93): hide the row if `G<row>` = 0.
   - Block 3 (rows 96-125): hide the row if `G<row>` = 0.
   - Use `.EntireRow.Hidden = True/False` (real hide, fully reversible, does not move/delete
     cells, does not touch any formula).

3. **Call it from the existing print macros**, guarded so behavior for other sheets is
   unchanged:
   - At the top of `Module1.PRINTOUTPUT` and `Module1.PRINTPROTECTEDOUTPUT`, add:
     `If ActiveSheet.Name = "FINAL SHEET" Then Call RefreshFinalSheetVisibility`
   - Keep all existing lines in both subs exactly as-is (AutoFilter/print calls for the other
     sheets must keep working unchanged).

4. **Do not modify**: any formula on `FINAL SHEET`, `TANK HARDWARE`, or `PRINT SHEET`; the
   `PRINT SHEET`/`PRINT QTY` flag-column + AutoFilter mechanism; sheet visibility/protection;
   the footer/sign-off block (rows 127-131); row 10/21/95 headers.

5. **Verify before/after**, with Excel fully closed except during the scripted COM session:
   - VBA macros preserved (`PRINTOUTPUT`, `PRINTPROTECTEDOUTPUT` still present, plus the new
     `RefreshFinalSheetVisibility`), all 6 sheets intact, all existing formulas byte-for-byte
     unchanged.
   - Programmatically invoke `RefreshFinalSheetVisibility` on the live file's current data and
     confirm: every row currently computing to 0 across its block's quantity column(s) ends up
     `EntireRow.Hidden = True`, every nonzero row stays visible, and the footer rows (127-131)
     are never hidden.
   - Confirm row-hide state resets correctly if run twice in a row (idempotent) and correctly
     un-hides a row that had been hidden by a prior run but now has nonzero data (simulate by
     toggling one PRINT SHEET quantity and re-running).

## Acceptance criteria
- Opening `TANK HARDWARE LIST.xlsm` and pressing the existing Print button on `FINAL SHEET`
  hides every row (across all 3 blocks) whose quantity is 0 for the current job's data, and
  shows every row with nonzero quantity — matching `PRINT SHEET`'s actual contents exactly.
- Pressing Print on `TANK HARDWARE` or `PRINT SHEET` behaves exactly as before (no new hiding
  logic runs there).
- No formula on any sheet changed. No data changed. No sheet added/removed/renamed.
- A pre-change backup exists.
- Verified with a real (non-interactive/headless) run against the live workbook's current data,
  not just code review.

## Non-goals / out of scope
- Not touching the HV Trunk zone bug on `TANK HARDWARE` (explicitly deferred by the user in an
  earlier message this session).
- Not converting the sheet to dynamic arrays/FILTER/UNIQUE (rejected — Excel-version risk and
  merged-cell incompatibility, per the trigger-timing decision already made with the user).
- Not adding a new standalone "Refresh" button (user chose the Print-button-triggered option).
- Not changing what happens when the user merely switches to/views `FINAL SHEET` without
  printing (rows stay in whatever state the last Print left them in — accepted tradeoff).

## Security notes
- Change is confined to this one local `.xlsm`; no external services, no network calls, no
  secrets involved.
- VBA edit is done via local Excel COM automation only, using an already-enabled local Trust
  Center setting — no macro security settings are being weakened as part of this change.
