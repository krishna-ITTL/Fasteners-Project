"""Replace one VBA module in a workbook with its .bas from this folder.

Usage: python update_module.py <workbook.xlsm> <ModuleName> [<ModuleName> ...]
Whole-module Import (never line-patching); saves the workbook.
"""
import os
import sys
import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
wb_path, names = os.path.abspath(sys.argv[1]), sys.argv[2:]
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(wb_path)
    comps = wb.VBProject.VBComponents
    for name in names:
        comps.Remove(comps.Item(name))
        comps.Import(os.path.join(HERE, name + ".bas"))
        print("imported", name)
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
