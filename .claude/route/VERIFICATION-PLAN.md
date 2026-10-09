# SPEC: VERIFICATION sheet — checking the hardware list against the standard

**Target:** `TANK HARDWARE LIST.xlsm`, new sheet `VERIFICATION`, positioned immediately
after `FINAL SHEET`.

---

## Goal

Give the engineer one sheet that answers three questions about the hardware list:

1. **Does the summary faithfully represent the input?** (reconciliation)
2. **Does each joint carry the hardware its standard demands?** (joint standard)
3. **Is the chosen bolt length long enough?** (stack-up)

Findings must be short and actionable. A report that cries wolf gets switched off.

---

## Settled decisions

| ID | Decision | Source |
|---|---|---|
| V1 | Source is **`TANK HARDWARE`**, not `PRINT SHEET` — it carries the joint type and the location text that make a finding actionable. | user |
| V2 | **Ignore columns AA:DF** on `TANK HARDWARE`. Only B, C, D, F, K, L, M, N, O, P, Q, R are read. | user |
| V3 | The standard in `STD.DATA` is **authoritative** and may be used for verification. | designer, via user |
| V4 | The **rulings table ships empty**; the user fills it in as the designer rules. | user |
| V5 | Sheet sits **immediately after `FINAL SHEET`**. | user |
| V6 | Bulk deviations are **grouped into one line**; individual outliers are **listed row by row**. | recommended, accepted |
| V7 | Manual **`VERIFY` button** only. Nothing auto-runs. | design |
| V8 | Source range is `TANK HARDWARE` **rows 15-250**, not 36-245. Row 13 is the header, 14 a section title, data starts at 15. Starting at 36 skipped the NB flange joints - 20 rows, 168 pieces. Check A caught this on its first run. A `RangeGuard` now refuses to run if hardware appears past row 250. | found in testing |
| V9 | A ruling **redefines** the accepted value, it does not merely mute rows. After `SD-269 P.WASHER = 2`, a row carrying 1 washer becomes a deviation and row 208's message reads "vs standard 2". This is intended. | found in testing |

### Evidence for V3

`SD-268 Gasket Joint` — 50 rows, nut and plain washer match the standard **exactly** on every
row. The standard is followed where it is followed, so deviations elsewhere are real signals.
This also disproves the "shop always adds a washer under the head" theory: if that were
universal, SD-268 would show 3 plain washers, not 2.

---

## The standard (`STD.DATA` rows 222–227, header at row 221)

Columns: `E`=joint name, `G`=BOLT/STUD, `H`=NUT, `I`=PLAIN WASHER, `J`=SPRING WASHER,
`K`=REF-STD. Read live every run, so a designer edit to the standard changes the check.

| Ref | Joint | Bolt/Stud | Nut | P.Washer | S.Washer |
|---|---|---|---|---|---|
| SD-267 | Tank curb Joint | 1 | 1 | 1 | 0 |
| SD-268 | Gasket Joint | 1 | 1 | 2 | 0 |
| SD-269 | Metal Joint | 1 | 1 | 1 | 1 |
| SD-270/REF-1 | Radiator Joint (Stud) | 1 | 3 | 3 | 2 |
| SD-270/REF-2 | Radiator Joint (Bolt) | 1 | 2 | 3 | 1 |
| SD-272 | Welded Pad Joint | 0 | 1 | 1 | 0 |

---

## Sheet layout

| Region | Rows | Contents |
|---|---|---|
| Title / job info | 2–6 | Title, work order, revision (mirrored from `TANK HARDWARE`), last-run stamp, **overall verdict** |
| RULINGS (input) | 9–18 | `Joint Ref │ Item │ Accepted per bolt │ Approved by │ Date │ Note` — **ships empty** |
| SUMMARY | 21–26 | One line per check: checked / passed / failed |
| FINDINGS | 29–178 | `TH Row │ Location │ Joint │ Designation │ Bolt │ Nut │ P.Wsh │ S.Wsh │ Finding │ Severity` |

Findings capacity 150 rows; unused rows hidden. Overflow aborts with a named message rather
than truncating, same guard pattern as `RebuildFinalSheet`.

---

## Checks

### A — Reconciliation
- **A1** `TANK HARDWARE` Σ(N) == `PRINT SHEET` Σ(G) == `FINAL SHEET` blocks 2+3
- **A2** every designation with qty > 0 appears in `FINAL SHEET`
- **A3** per-designation quantity agrees between `TANK HARDWARE` and `FINAL SHEET`

*Expected today: clean.* This is the guard against the `RebuildFinalSheet` macro going wrong.

### B — Joint standard
For each source row with a designation:

- Joint ref absent from `STD.DATA` → **UNVERIFIABLE** (today: `PSD-510`)
- Standard Bolt/Stud > 0 → denominator is `N`; compare `O/N` vs nut, `Q/N` vs p.washer,
  `R/N` vs s.washer
- Standard Bolt/Stud = 0 (welded pad) → a per-bolt ratio is undefined. Verify `N` = 0, then
  compare washers against the **nut** count using the standard's own nut ratio.
- A matching **ruling** replaces the standard's value for that joint + item.

*Expected today: 2 grouped bulk patterns + 11 individual outliers.*

### C — Length
- **C1** `L` (chosen) ≥ `K` (required stack-up) → **ERROR** if short. *Expected today: clean.*
- **C3** size named in the description text disagrees with the designation → **INFO**, never
  ERROR. The four foundation bolts are all deliberately +20 mm (`M24x300`→`M24x320`,
  `M20x250`→`M20x270`, `M16x200`→`M16x220`, `M12x200`→`M12x220`) for grouting projection.

### Severity
`ERROR` contradicts the standard with no ruling · `INFO` needs a human eye ·
`UNVERIFIABLE` joint type absent from the standard.

---

## Grouping rule (V6)

Group findings by `joint + item + actual ratio + expected value`. A group of **5 or more**
rows prints as one line carrying the row count; anything smaller prints row by row with its
location text. If a bulk pattern later shrinks below 5, it starts listing individually on its
own.

---

## Implementation

New VBA module `VerificationBuilder`, entry point `RunVerification`, following the pattern
already proven in `FinalSheetBuilder`:

- bulk `.Value` read of the source, one hit
- `Scripting.Dictionary` aggregation
- capacity guard, host-state save/restore, `gSilent` test seam
- values written into the sheet, unused rows hidden
- injected via `.bas` **Import** (never `DeleteLines` — see the Module1 incident)

---

## Acceptance criteria

1. Findings match a Python-computed expected set **exactly** — same rows, same severities.
2. `PSD-510` reported UNVERIFIABLE, not ERROR.
3. The four foundation bolts appear as INFO, not ERROR.
4. Check A reports clean against the current workbook.
5. Adding a ruling `SD-269 │ P.WASHER │ 2` removes the 96-row group, re-bases row 208's message
   to "vs standard 2", and turns the single 1-washer row into a deviation - net 95 fewer errors.
   Verified by running the reference implementation with the same ruling.
6. Removing that ruling brings the group back.
7. A deliberately injected error (change a nut count) is caught and named.
8. Idempotent: three consecutive runs identical.
9. `FINAL SHEET`, `TANK HARDWARE`, `PRINT SHEET` and `STD.DATA` unmodified.
10. `Module1` print macros still intact afterwards.
11. Workbook opens clean, no repair prompt.

## Non-goals

- Do not modify any existing sheet.
- Do not read or write `TANK HARDWARE` AA:DF.
- Do not auto-run the check.
- Do not pre-populate the rulings table.
- Do not "fix" anything the check finds — report only.
