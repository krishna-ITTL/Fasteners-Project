====================================================================
READ ME FIRST - Plain-English guide to the TANK HARDWARE LIST macros
====================================================================

This folder is a backup of the "behind the scenes" automation (called
"macros" or "VBA code") that lives inside the file:

    TANK HARDWARE LIST_FINAL_KRISH.xlsm

A macro is just a small program that Excel runs when you click a
button or press a keyboard shortcut. It does, automatically and
instantly, the same clicking/typing/copying you would otherwise have
to do by hand - but without the risk of a manual mistake.

There are only THREE files worth reading here. Excel workbooks also
carry a lot of empty "container" code that Excel creates by itself
for every sheet, even if nothing was ever written in it - those are
not included because there is nothing in them to explain.

--------------------------------------------------------------------
1. "1 - Print Buttons (Module1).txt"
--------------------------------------------------------------------
What it's for: the two print shortcuts, Ctrl+Shift+P and
Ctrl+Shift+Q.

In plain terms: when you press one of these, the macro makes sure
FINAL SHEET is up to date, then sends the current sheet to the
printer. The "Q" version also re-locks (protects) the sheet
afterwards so nobody can accidentally type over it.

--------------------------------------------------------------------
2. "2 - FINAL SHEET Builder (FinalSheetBuilder).txt"
--------------------------------------------------------------------
What it's for: the REFRESH button on FINAL SHEET.

In plain terms: FINAL SHEET is your procurement summary - the short
list of "buy this many of this size, in this material" that a
purchaser actually orders from. This macro is the thing that builds
that list. It reads every single row of hardware from PRINT SHEET,
adds up how many bolts/studs/nuts/washers of each size-and-material
combination are needed across the whole tank, and writes that
summary onto FINAL SHEET - including refreshing the live totals and
hiding the rows that aren't needed for this job.

This is the most important macro in the workbook. If FINAL SHEET
ever looks wrong, this is the code to look at.

--------------------------------------------------------------------
3. "3 - Verification Checker (VerificationBuilder).txt"
--------------------------------------------------------------------
What it's for: the VERIFY button on the VERIFICATION sheet.

In plain terms: this is a second pair of eyes. It goes through every
hardware line on TANK HARDWARE, compares it against the company
standard (STD.DATA) for that type of joint, and flags anything that
looks off - wrong number of nuts or washers for the bolts ordered,
a bolt shorter than the joint requires, a joint type it doesn't
recognise, or the grand totals on TANK HARDWARE / PRINT SHEET /
FINAL SHEET not matching each other. It never changes any of your
data - it only writes a report.

--------------------------------------------------------------------
Why this backup exists
--------------------------------------------------------------------
The workbook file itself (the .xlsm) already contains all of this
code - you don't need these text files for the workbook to work.
This folder is purely a paper trail: a plain-English record of what
the automation does, kept separately in case the workbook is ever
lost, corrupted, or the macro code needs to be reviewed by someone
without opening Excel's code editor.

Each .txt file below contains the REAL code exactly as it exists in
the workbook today, with extra explanation lines added above the
tricky parts. Every explanation line starts with a single quote (')
followed by ">> PLAIN ENGLISH:" so you can tell at a glance which
lines are the original programmer's code and which lines are the
plain-English explanation added for this backup.
