"""One-off: measure PRINT SHEET by its used area (filter-proof) in all three
builders, plus the small VERIFICATION review fixes. Run from .claude/route/."""


def patch(p, pairs):
    s = open(p, encoding='latin-1').read().replace('\r\n', '\n')
    for old, new in pairs:
        assert s.count(old) == 1, (p, old[:60], s.count(old))
        s = s.replace(old, new)
    open(p, 'w', encoding='latin-1', newline='\r\n').write(s)


patch('FinalSheetBuilder.bas', [(
'''    ' Find, not End(xlUp): PRINT SHEET is filtered to hide zero rows, and
    ' End(xlUp) skips hidden rows - a joint that became non-zero at the
    ' bottom of the list (e.g. foundation bolts) was silently left out.
    ' LookIn:=xlFormulas also searches hidden rows.
    Dim lastCell As Range
    Set lastCell = ps.Columns("F").Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastRow = PS_FIRST Else lastRow = lastCell.Row
''',
'''    ' Bottom of the sheet's used area. PRINT SHEET is filtered to hide zero
    ' rows, and both End(xlUp) and Find skip rows the filter has hidden - a
    ' joint that became non-zero at the bottom of the list (e.g. foundation
    ' bolts) was silently left out. UsedRange ignores the filter. Extra
    ' rows below the list (notes, signatures) have no part name and are
    ' skipped by the loop below.
    lastRow = ps.UsedRange.Row + ps.UsedRange.Rows.Count - 1
''')])

patch('ProductionSheetBuilder.bas', [(
'''' Uses Find, not Ctrl+Up (End(xlUp)): PRINT SHEET is filtered to hide
' zero rows, and End(xlUp) jumps over hidden rows - so a joint that just
' became non-zero at the bottom of the list (e.g. foundation bolts) was
' silently left out. Find with LookIn:=xlFormulas also searches hidden rows.
Private Function LastSourceRow(ByRef ps As Worksheet) As Long
    Dim lastCell As Range
    Set lastCell = ps.Columns("F").Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then LastSourceRow = PS_FIRST Else LastSourceRow = lastCell.Row
''',
'''' Uses the bottom of the sheet's used area. PRINT SHEET is filtered to
' hide zero rows, and both Ctrl+Up (End(xlUp)) and Find skip rows the
' filter has hidden - so a joint that just became non-zero at the bottom
' of the list (e.g. foundation bolts) was silently left out. UsedRange
' ignores the filter. Rows below the list (notes, signatures) have no
' part name and are skipped anyway.
Private Function LastSourceRow(ByRef ps As Worksheet) As Long
    LastSourceRow = ps.UsedRange.Row + ps.UsedRange.Rows.Count - 1
''')])

patch('VerificationBuilder.bas', [
(
'''    ' Find, not End(xlUp): PRINT SHEET is filtered to hide zero rows and
    ' End(xlUp) skips hidden rows, cutting off the bottom of the list.
    Set lastCell = ps.Columns("F").Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastPs = 9 Else lastPs = lastCell.Row
''',
'''    ' Bottom of the used area: PRINT SHEET is filtered to hide zero rows,
    ' and End(xlUp) and Find both skip filter-hidden rows, cutting off the
    ' bottom of the list. Blank rows below it add 0.
    lastPs = ps.UsedRange.Row + ps.UsedRange.Rows.Count - 1
'''),
(
'''    Dim r As Long, lastPs As Long, lastCell As Range
''',
'''    Dim r As Long, lastPs As Long
'''),
# End(xlDown) could run past a gap in the table: walk down instead.
(
'''    If Len(Trim$(CStr(sd.Cells(firstRow, SD_REF_COL).Value & ""))) = 0 Then
        lastRow = firstRow - 1
    Else
        lastRow = sd.Cells(firstRow, SD_REF_COL).End(xlDown).Row
    End If
''',
'''    ' Walk down to the first blank REF-STD cell (End(xlDown) would jump
    ' past a one-row table to whatever is next in column K).
    lastRow = firstRow - 1
    Do While Len(Trim$(CStrSafe(sd.Cells(lastRow + 1, SD_REF_COL).Value))) > 0
        lastRow = lastRow + 1
    Loop
'''),
# error values in a designation: report, don't crash
(
'''    Dim s As Variant, act As Double, want As Double

    For i = 1 To UBound(src, 1)
        des = Trim$(CStr(src(i, 12) & ""))
        If Len(des) = 0 Then GoTo NextRow
        r = firstRow + i - 1
''',
'''    Dim s As Variant

    For i = 1 To UBound(src, 1)
        r = firstRow + i - 1
        If IsError(src(i, 12)) Then
            ' A #REF!/#N/A part name would otherwise stop the whole run.
            AddFinding fnd, r, CStrSafe(src(i, 1)), CStrSafe(src(i, 3)), "(error value)", "-", _
                       "ERROR", "part name shows an error value", "|DESIG|ERROR|", 0, 0, 0, 0
            GoTo NextRow
        End If
        des = Trim$(CStr(src(i, 12) & ""))
        If Len(des) = 0 Then GoTo NextRow
'''),
(
'''        joint = Trim$(CStr(src(i, 3) & ""))
        loc = Trim$(CStr(src(i, 1) & ""))
        n = Nz(src(i, 13)): nut = Nz(src(i, 14))''',
'''        joint = Trim$(CStrSafe(src(i, 3)))
        loc = Trim$(CStrSafe(src(i, 1)))
        n = Nz(src(i, 13)): nut = Nz(src(i, 14))'''),
(
'''    For i = 1 To UBound(src, 1)
        des = Trim$(CStr(src(i, 12) & ""))
        If Len(des) = 0 Then GoTo NextRow
        r = firstRow + i - 1
        loc = Trim$(CStr(src(i, 1) & ""))''',
'''    For i = 1 To UBound(src, 1)
        des = Trim$(CStrSafe(src(i, 12)))
        If Len(des) = 0 Then GoTo NextRow
        r = firstRow + i - 1
        loc = Trim$(CStrSafe(src(i, 1)))'''),
(
'''        If Len(Trim$(CStr(th.Cells(r, 13).Value & ""))) > 0 Then _''',
'''        If Len(Trim$(CStrSafe(th.Cells(r, 13).Value))) > 0 Then _'''),
(
'''        If Len(Trim$(CStr(sdi.Cells(r, 13).Value & ""))) > 0 Then _''',
'''        If Len(Trim$(CStrSafe(sdi.Cells(r, 13).Value))) > 0 Then _'''),
(
'''                ' Name the actual TANK HARDWARE rows, compressed into''',
'''                ' Name the actual TANK HARDWARE / STD DATA rows, compressed into'''),
(
'''Private Function SepComma(ByVal s As String) As String
    If Len(s) > 0 Then SepComma = ","
End Function
''',
'''' Text of a cell value; an error value (#N/A, #REF!) becomes "" instead of
' stopping the macro with "Type mismatch".
Private Function CStrSafe(ByVal v As Variant) As String
    If IsError(v) Then
        CStrSafe = ""
    Else
        CStrSafe = CStr(v & "")
    End If
End Function
'''),
])
print("patched")
