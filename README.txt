
INDOTECH HARDWARE AUTOMATION PLATFORM
======================================

1) HOW TO USE
--------------
1. Click Select Hardware PDF.
2. Choose PDF.
3. Select output folder.
4. Excel & PDF auto generated.

2) IF APPLICATION FAILS
-----------------------
- Close already opened Excel file.
- Ensure PDF is not scanned image.
- Ensure Excel is installed.

3) WEIGHT TABLE
---------------
Only HDG coating values are active.
MAGNI, BLACK, SS, HT will show N/A until
their values are added to WEIGHT_TABLE.

To add MAGNI rates later, open the code and add:
  "M16X50-MAGNI": 27,
under the WEIGHT_TABLE section.

4) ROUNDING RULE
----------------
Odd lengths (X25, X35, X45...) are automatically
rounded up to nearest 10 (X30, X40, X50...).
Output displays the rounded size.

5) LATEST NEWS
---------------
2026-08-24 - Added a full backup of TANK HARDWARE LIST_FINAL_KRISH.xlsm
under backups/Final BackUp/, together with the three real macros
(print buttons, FINAL SHEET builder, VERIFICATION checker) saved as
plain-text files with line-by-line, plain-English explanations - so
the automation logic can be reviewed by anyone, not just someone who
reads VBA. The older dated .xlsm-only snapshots from 22-Aug (before
the FINAL SHEET rebuild and verification work) stay alongside it in
backups/.

======================================
IndoTech Internal Software
======================================
