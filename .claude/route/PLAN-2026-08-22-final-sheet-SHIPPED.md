# PLAN: FINAL SHEET — self-building, material-aware procurement summary

**Target file:** `D:\ITTL_PROJECTS\Fasteners-Project\TANK HARDWARE LIST.xlsm`, sheet `FINAL SHEET`
**Backup taken:** `backups/TANK HARDWARE LIST_backup_20260822_103421_before_final_sheet_rebuild.xlsm`

---

## Goal

`FINAL SHEET` is the one-page, size-wise procurement summary — "how many of each fastener do I
buy". Today its three blocks are driven by a **hand-typed list of sizes**. The sheet can only
ever display combinations someone typed in advance, and its MATERIAL column is not derived from
the data at all.

Make `FINAL SHEET` build its own row list from `PRINT SHEET`, keyed on
**size + length + material**, refreshed by a new **Refresh button** on the sheet (and also
automatically before any print).

---

## Background — verified against the live file

### Current geometry (`FINAL SHEET`)

| Region | Rows | Contents |
|---|---|---|
| Header / job info | 2–8 | Work order, revision, `E7` = material ≤M12, `E8` = material >M12 |
| Block 1 header | 10 | `D10`=SIZE, `F10`=MATERIAL, `G10`=NUT, `H10`=P. WASHER, `I10`=S. WASHER, `J10`=L.NUT |
| **Block 1 data** | **11–19 (9 rows)** | M4, M5, M6, M8, M10, M12, M16, M20, M24 |
| blank | 20 | separator |
| Block 2 header | 21 | `D21`=BOLT SIZE, `F21`=MATERIAL, `G21`=BOLT QUANTITY |
| **Block 2 data** | **22–93 (72 rows)** | 36 sizes × 2 material slots |
| blank | 94 | separator |
| Block 3 header | 95 | `D95`=STUD SIZE, `F95`=MATERIAL, `G95`=STUD QUANTITY |
| **Block 3 data** | **96–125 (30 rows)** | 15 stud sizes × 2 material slots |
| blank | 126 | separator |
| Footer | 127–131 | NOTE / "ALL STUDS ARE FIXED/WELD BY TANK VENDOR." / PREPARED–CHECKED–APPROVED BY |

- **232 merged ranges**, per-row: `D<r>:E<r>` and `G<r>:J<r>` in blocks 2 and 3.
- **Print area is fixed:** `'FINAL SHEET'!$D$2:$J$131`.
- **No other sheet references `FINAL SHEET`** (verified: 0 cross-sheet refs). Geometry is safe
  to rebalance.
- No sheet protection anywhere in the workbook.

### Root cause of the reported bug

`F22` is literally `=$E$7`, and `G22` is
`=SUMIF('PRINT SHEET'!$F:$F, $D22&"-"&$F22, 'PRINT SHEET'!$G:$G)`.

Material is **not read from the data** — it is forced to the global `$E$7` / `$E$8` values.
Rows 22–57 use `$E$7` and rows 58–93 use `$E$8`; these are two *material slots*, not two
*materials*. Since neither global value is `SS`, the designation `M12x60-SS` can never be
matched and its 2 pieces can never appear. Same defect in Block 3.

### Existing VBA (`Module1`)

- `PRINTOUTPUT` (Ctrl+Shift+P) and `PRINTPROTECTEDOUTPUT` (Ctrl+Shift+Q) — bound to Print
  buttons. Both begin with
  `If ActiveSheet.Name = "FINAL SHEET" Then Call RefreshFinalSheetVisibility`.
- `RefreshFinalSheetVisibility` — added in the prior session. **Hides only.** Unhides
  `11:19`, `22:93`, `96:125`, then re-hides rows whose quantity is 0.

This macro is the foundation to extend — it already runs at the right moment, from the right
trigger, against the right ranges.

### Environment

- Excel: Microsoft 365 Business, build **16.0.20228** (Current Channel).
- Dynamic arrays **available** (`UNIQUE`, `FILTER`, `SORT`, `LET` all verified working) — but
  **deliberately not used**, see Decision D1.
- `AccessVBOM = 1` → VBA project is programmatically editable via `win32com.client`.
- `VBAWarnings = 1` → macros run without prompting.
- `openpyxl` **cannot** edit VBA source (opaque `vbaProject.bin` round-trip only) →
  all VBA work must go through Excel COM.
- `.venv` has `openpyxl 3.1.5` + `pywin32 312`. **Note:** `venv/` (the older one) is broken —
  its `python.exe` shim points at a stale user path. Use `.venv` only.

---

## Decisions already settled with the user

| ID | Decision | Rationale |
|---|---|---|
| **D1** | **VBA rebuild, not dynamic-array spill.** | User requires the merged `D:E` / `G:J` look. Excel returns `#SPILL!` when a dynamic array targets merged cells — hard product limit. A macro writing *values* into merged cells is allowed. This exact conflict caused the prior session's work to be reverted. |
| **D2** | **Uniqueness key = size + length + material** (Option B). | `M12X60` exists as 1078 × HDG-8.8 and 2 × SS. `TANK HARDWARE!M223` hand-types `"SS"` — the only such override in 435 rows — because the NRV sits inside the oil. Merging them would silently drop the 2 stainless bolts. User confirmed: *"there are few items that will be forcefully changed in the Tank hardware sheet itself"*. |
| **D3** | **Block 1 splits by material too.** | Same reason. M12 → `HDG-8.8` (2345 nuts) and `SS` (2 nuts). |
| **D4** | **Zero-total studs excluded from Block 3; their nuts/washers stay in Block 1.** | 29 `PRINT SHEET` rows are SD-272 *Welded Pad Joints* — the tank vendor welds the studs on, so IndoTech buys 0 studs but does buy 571 nuts / 582 P.washers / 168 S.washers. Matches the sheet's own note at `D128`. |
| **D5** | **Filter after aggregating, never per row.** | `M12-STUDX85` appears with 6 studs (row 36) *and* 0 studs (rows 124/128/133/137). Its total is 6, so it stays. Only drop a designation whose **total** is 0. |
| **D6** | **Separate Refresh button** on `FINAL SHEET`; Print also refreshes. | User: *"make a separate refresh button in Final sheet"*. Lets them review on screen without printing. Print keeps refreshing so a stale sheet can never be printed. |

### Assumption to confirm (low risk, easily reversed)

**A1 — Row layout.** The user wrote the split rows as full designations
(`M12x60-HDG-8.8` / `M12x60-SS`), but the sheet has a separate MATERIAL column and the user
approved a mock-up showing `SIZE` = `M12X60` with `MATERIAL` = `HDG-8.8` / `SS` on its own row.
**Proceeding with SIZE + separate MATERIAL column**, which preserves the existing headers and
keeps the MATERIAL column meaningful. Flag to user; a one-line change if wrong.

---

## Design problem: Block 1 has no headroom

Block 1 has **9 slots** (rows 11–19). After D3, WO 6449 needs **exactly 9**. One additional
material — or one additional size — overflows it immediately.

Meanwhile Block 2 has 72 slots for ~30 real rows and Block 3 has 30 for ~6.

**Resolution — one-time rebalance, footer and print area unchanged.** Total space rows 11–125
(115 rows) is reallocated:

| Region | New rows | Capacity | Needed (WO 6449) |
|---|---|---|---|
| Block 1 header | 10 | — | — |
| **Block 1 data** | **11–34** | **24** | 9 |
| blank | 35 | — | — |
| Block 2 header | 36 | — | — |
| **Block 2 data** | **37–96** | **60** | 30 |
| blank | 97 | — | — |
| Block 3 header | 98 | — | — |
| **Block 3 data** | **99–125** | **27** | 6 |
| blank | 126 | — | — |
| Footer | **127–131 (unchanged)** | — | — |

Footer stays at 127–131, print area `$D$2:$J$131` stays valid, no cross-sheet refs to fix.

---

## Implementation steps

### 1. Safety

- Backup already taken (path above). Take a second backup immediately before the COM write.
- Excel currently has the workbook open (PID 21116, `Saved = True`). Close it cleanly via COM
  before modifying, or drive the modification through the live instance — do not do both.
- After every write, save and reopen to confirm the file is not corrupt.

### 2. One-time layout rebalance

- Move Block 2 (header + data) from rows 21–93 to rows 36–96.
- Move Block 3 (header + data) from rows 95–125 to rows 98–125.
- Extend Block 1 data to rows 11–34.
- Recreate per-row merges `D<r>:E<r>` and `G<r>:J<r>` across all three blocks' full new
  capacity, and carry the existing borders / fonts / number formats onto the new rows so
  freshly-written rows look identical to existing ones.
- Leave the footer at 127–131 untouched.

### 3. Rewrite `RefreshFinalSheetVisibility` → `RebuildFinalSheet`

New procedure in `Module1`, written via COM
(`VBProject.VBComponents("Module1").CodeModule`). Keep the old name as a thin wrapper that
calls the new one, so existing Ctrl+Shift+P / Ctrl+Shift+Q bindings keep working.

Algorithm:

1. `Application.ScreenUpdating = False`, `Application.EnableEvents = False`. Restore both in an
   error handler so a mid-run failure cannot leave Excel in a broken state.
2. Unhide all rows 11–125. Clear the **values** in `D`, `F`, `G:J` across all three data
   ranges. Do **not** clear formatting and do **not** unmerge.
3. Read `PRINT SHEET` columns F (designation), G (bolt/stud qty), H (nut), I (P.washer),
   J (S.washer), K (L.nut) into a VBA array in one hit — a single `.Value` read, not
   cell-by-cell, for speed.
4. Parse each designation in column F. Two shapes:
   - Bolt: `M12x60-HDG-8.8` → size `M12`, length `60`, material `HDG-8.8`
   - Stud: `M12-STUDx60-HDG-8.8` → size `M12`, length `60`, material `HDG-8.8`, is-stud
   Split material at the **first** `-` that follows the length. Material may itself contain
   `-` (`HDG-8.8`), so take everything after that point as the material.
   Anything unparseable → collect and report, never silently skip (see step 8).
5. Aggregate with `Scripting.Dictionary`:
   - Block 1 key: `size|material` → sum NUT, P.WASHER, S.WASHER, L.NUT
   - Block 2 key: `sizeXlength|material`, bolts only → sum qty
   - Block 3 key: `sizeXlength|material`, studs only → sum qty
6. Apply D4/D5: drop a **Block 3** entry only if its aggregated stud qty = 0. Block 1 keeps
   its nuts/washers regardless. Drop Block 2 entries whose aggregated qty = 0.
7. Sort each block: numeric by size (M4 < M5 < M10 < M12 …, **not** text sort — text puts M10
   before M4), then numeric by length, then material alphabetically.
8. Write rows:
   - `D` = size (Block 1) or `sizeXlength` (Blocks 2/3) — **value**
   - `F` = material — **value**
   - `G:J` (Block 1) / `G` (Blocks 2/3) = live `SUMIF` **formula** matching on the exact
     designation, so quantities stay live between refreshes.
   - **Capacity guard:** if a block's row count exceeds its capacity, abort with a clear
     `MsgBox` naming the block and the required vs available rows. Never truncate silently.
   - **Parse guard:** if any designation failed to parse, report them in the same `MsgBox`.
9. Hide every unused row in each block's range.
10. Restore `ScreenUpdating` / `EnableEvents`.

### 4. Add the Refresh button

- A `Shapes`/Forms button on `FINAL SHEET`, positioned clear of the print area (e.g. anchored
  near column L, or inside `$D$2:$J$131` only if it is set `Print object = False`), captioned
  **"REFRESH"**, assigned to `RebuildFinalSheet`.
- Match the styling of the existing Print button so it does not look bolted on.
- Confirm it does not appear in printed output.

### 5. Also refresh on Print

`PRINTOUTPUT` and `PRINTPROTECTEDOUTPUT` already call `RefreshFinalSheetVisibility` when
`FINAL SHEET` is active. Since that name now forwards to `RebuildFinalSheet`, no edit is
needed — **verify** rather than assume.

### 6. Cleanup

Remove the leftover `_xlpm.*` and `_xleta.*` defined names (14 of them, all resolving to
`#NAME?`) left behind by the prior session's abandoned `LET`/`LAMBDA` experiment.

---

## Acceptance criteria

Run against the live WO 6449 data.

1. **Totals reconcile.** Block 2 bolts + Block 3 studs = **3538**, equal to the `PRINT SHEET`
   column G total. (Today: 6606 — wrong, because the static list double-counts across the two
   material slots.)
2. **Block 2 = 30 rows / 3476 bolts.** **Block 3 = 6 rows / 62 studs.**
3. **Block 1 = 9 rows**, including a distinct `M12 | SS` row.
4. **`M12X60 | SS | 2` appears in Block 2** as its own row, alongside `M12X60 | HDG-8.8 | 1078`.
   This is the headline fix.
5. **Previously invisible hardware now appears:** `M16x75`, `M16x120`, `M20x50` and the two
   extra stud lengths — the 189 pieces the prior session found and then lost on revert.
6. **No zero-quantity row is visible** in any block.
7. **`M12-STUDX85` is present in Block 3** (total 6 across mixed rows) — proves D5 aggregate-
   then-filter, not per-row filter.
8. **Block 1 M12 nut count includes stud-derived nuts.** The 571 nuts / 582 P.washers /
   168 S.washers from zero-stud welded-pad joints must still be counted.
9. **Merged cells intact** — `D:E` and `G:J` merged on every visible row of blocks 2 and 3.
10. **Footer still at rows 127–131**; print area still `$D$2:$J$131`; printed page looks
    unchanged in structure.
11. **Refresh button works** from a cold open and is not printed.
12. **Idempotent:** clicking Refresh 3× in a row produces byte-identical results.
13. **Reacts to change:** delete a size from `PRINT SHEET`, click Refresh → row disappears.
    Add one back → row reappears.
14. **Quantities are live formulas**, not pasted values — editing a qty in `TANK HARDWARE`
    updates `FINAL SHEET` without pressing Refresh.
15. **File still opens clean** in Excel with macros enabled; no repair prompt.

---

## Non-goals — do not do these

- Do **not** touch `TANK HARDWARE`, `PRINT SHEET`, `STD.DATA`, `PRINT QTY`, or `NON.STD.DATA`.
- Do **not** unmerge `D:E` or `G:J`. That was explicitly reverted once already.
- Do **not** introduce dynamic-array formulas on `FINAL SHEET`.
- Do **not** add sheet protection.
- Do **not** change the Print button behaviour or its keyboard shortcuts.
- Do **not** move the footer or alter the print area range.
- Do **not** "improve" `hardware_generator.py` or anything outside this sheet.
- Do **not** delete rows to hide them — hide only, fully reversible.

---

## Risks

| Risk | Mitigation |
|---|---|
| COM write corrupts the workbook | Backup before every write; reopen and verify after |
| Designation parsing misses an edge case | Parse guard reports unparseable strings rather than dropping them |
| Text sort puts M10 before M4 | Explicit numeric sort on size then length |
| Macro clears Excel's undo stack | Known and accepted; inherent to VBA writes |
| Block overflows on a future job | Capacity guard aborts with a named, actionable message |
| Layout rebalance breaks pagination | Footer + print area held fixed; verify printed output |
