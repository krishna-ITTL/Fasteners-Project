"""Print PRODUCTION SHEET to PDF and check no block is split across pages.

Usage: python check_pagebreaks.py <workbook.xlsm> <out.pdf>
Opens the workbook read-only in a separate Excel. A block is "split" if its
title and its TOTAL line land on different pages. Exits non-zero if any are.
"""
import os
import re
import sys
import pdfplumber
import win32com.client

wb_path, pdf = os.path.abspath(sys.argv[1]), os.path.realpath(sys.argv[2])
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    # The PC's default printer may not do A4 (it exported Letter here), so
    # check with Microsoft Print to PDF, which honours the sheet's A4 setting.
    wb = xl.Workbooks.Open(wb_path, ReadOnly=True)
    ws = wb.Sheets("PRODUCTION SHEET")
    print("paper:", ws.PageSetup.PaperSize, "| zoom:", ws.PageSetup.Zoom, "| manual breaks:",
          [ws.HPageBreaks(i).Location.Row for i in range(1, ws.HPageBreaks.Count + 1)
           if ws.HPageBreaks(i).Type == -4135])
    # ExportAsFixedFormat always writes Letter on this PC, so the workbook
    # under test must be set to Letter paper too (edge copy) - otherwise the
    # macro plans A4 pages and the PDF shows Letter pages.
    ws.ExportAsFixedFormat(0, pdf)
    wb.Close(False)
finally:
    xl.Quit()

title = re.compile(r"^M\d+(-STUD)?X\d+-\S+")
split, blocks = [], 0
with pdfplumber.open(pdf) as doc:
    for n, page in enumerate(doc.pages, 1):
        open_block = None
        lines = (page.extract_text() or "").splitlines()
        if not any(l.startswith("JOINT LOCATION") for l in lines):
            print(f"page {n}: no blocks")
        for line in lines:
            if title.match(line):
                if open_block:
                    split.append((n, open_block))
                open_block, blocks = line.split()[0], blocks + 1
            elif line.startswith("TOTAL"):
                if open_block is None:
                    split.append((n, "(TOTAL without title at top of page)"))
                open_block = None
        if open_block:
            split.append((n, open_block))
        print(f"page {n}: {sum(1 for l in lines if title.match(l))} blocks")
    pages = len(doc.pages)

print(f"{pages} pages, {blocks} block titles, split: {split}")
sys.exit(1 if split else 0)
