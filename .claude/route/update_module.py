"""Replace VBA modules in a workbook from "VBA code for Fastener", then
export them back so that folder always matches what Excel holds.

Usage: python update_module.py <workbook.xlsm> <ModuleName> [<ModuleName> ...]
Whole-module Import (never line-patching); saves the workbook.
"""
import os
import sys
import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
VBA_DIR = os.path.abspath(os.path.join(HERE, "..", "..", "VBA code for Fastener"))
wb_path, names = os.path.abspath(sys.argv[1]), sys.argv[2:]
xl = win32com.client.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(wb_path)
    comps = wb.VBProject.VBComponents
    for name in names:
        comps.Remove(comps.Item(name))
        comps.Import(os.path.join(VBA_DIR, name + ".bas"))
        comps.Item(name).Export(os.path.join(VBA_DIR, name + ".bas"))
        print("imported + re-exported", name)
    wb.Save()
    wb.Close(False)
finally:
    xl.Quit()
