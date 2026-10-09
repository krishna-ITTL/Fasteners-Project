Attribute VB_Name = "Module1"
Sub PRINTOUTPUT()
Attribute PRINTOUTPUT.VB_Description = "PRINT OUTPUT PAGE"
Attribute PRINTOUTPUT.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' PRINTOUTPUT Macro
' PRINT OUTPUT PAGE
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    If ActiveSheet.Name = "FINAL SHEET" Then Call RefreshFinalSheetVisibility
    ' FINAL SHEET has no AutoFilter - calling ApplyFilter on Nothing
    ' raised run-time error 91 and stopped the print.
    If Not ActiveSheet.AutoFilter Is Nothing Then ActiveSheet.AutoFilter.ApplyFilter
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
Sub PRINTPROTECTEDOUTPUT()
Attribute PRINTPROTECTEDOUTPUT.VB_Description = "PRINT PROTECTED PAGE"
Attribute PRINTPROTECTEDOUTPUT.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' PRINTPROTECTEDOUTPUT Macro
' PRINT PROTECTED PAGE
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
    If ActiveSheet.Name = "FINAL SHEET" Then Call RefreshFinalSheetVisibility
    ActiveSheet.Unprotect
    If Not ActiveSheet.AutoFilter Is Nothing Then ActiveSheet.AutoFilter.ApplyFilter
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
