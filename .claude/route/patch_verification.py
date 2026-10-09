"""One-off: adapt the designers' VerificationBuilder.bas (Oct 2026) to check
STD DATA joints too, restore the header-based standard-table lookup, and fix
the reconciliation total. Run once from .claude/route/."""
p = 'VerificationBuilder.bas'
s = open(p, encoding='latin-1').read().replace('\r\n', '\n')


def rep(old, new, count=1):
    global s
    assert s.count(old) == count, (old, s.count(old))
    s = s.replace(old, new)


rep("""' Reads TANK HARDWARE (the source with joint type and location text) and
' compares each line against its joint's entry in STD.DATA.""",
"""' Reads both input sheets - TANK HARDWARE and STD DATA (the standard
' joints: valves, plates, flanges; split out in Oct 2026) - and compares
' each line against its joint's entry in the standard table on the
' hidden STD.DATA sheet (note the dot: a different sheet from STD DATA).
' Findings name the sheet and row, e.g. "TH 60" or "SD 44".""")

rep("""Public Const SD_FIRST As Long = 222
Public Const SD_LAST  As Long = 227
""", """
' --- STD DATA input sheet (same columns as TANK HARDWARE) ---
Public Const SDI_FIRST As Long = 15
Public Const SDI_LAST  As Long = 150
Public Const SDI_GUARD As Long = 220

' --- STD.DATA standard table ---
' Found by its "REF-STD" header in column K, not by fixed row numbers:
' the designers added SD-267 and SD-274 above/below the old rows
' 222-227, and the fixed rows silently skipped them, reporting 21 joints
' as UNVERIFIABLE for no reason. See LoadStandard.
Public Const SD_REF_COL    As Long = 11        ' column K: REF-STD
Public Const SD_REF_HEADER As String = "REF-STD"
""")

rep("""Public gvSilent As Boolean
Public gvLastProblem As String
""", """Public gvSilent As Boolean
Public gvLastProblem As String

' Which input sheet the checks are reading right now: "TH" or "SD".
' AddFinding stamps it on every finding so the report can say where.
Private mTag As String
""")

rep("""    Dim ws As Worksheet, th As Worksheet, sd As Worksheet
""", """    Dim ws As Worksheet, th As Worksheet, sd As Worksheet, sdi As Worksheet
    Dim src2 As Variant, sdMsg As String
""")
rep("""    Set sd = ThisWorkbook.Sheets("STD.DATA")
""", """    Set sd = ThisWorkbook.Sheets("STD.DATA")      ' standard table
    Set sdi = ThisWorkbook.Sheets("STD DATA")     ' standard-joint inputs
""")
rep("""    Set std = LoadStandard(sd)
    Set rul = LoadRulings(ws)
    Set fnd = New Collection

    ' Columns B..R of TANK HARDWARE. AA:DF are deliberately never read.
    src = th.Range("B" & TH_FIRST & ":R" & TH_LAST).Value

    msg = RangeGuard(th)
    If Len(msg) > 0 Then GoTo Abort

    CheckJointStandard src, std, rul, fnd
    CheckLength src, fnd
    CheckReconciliation fnd
""", """    Set std = LoadStandard(sd, sdMsg)
    If Len(sdMsg) > 0 Then
        msg = sdMsg
        GoTo Abort
    End If
    Set rul = LoadRulings(ws)
    Set fnd = New Collection

    ' Columns B..R of each input sheet (same layout on both).
    ' AA:DF are deliberately never read.
    src = th.Range("B" & TH_FIRST & ":R" & TH_LAST).Value
    src2 = sdi.Range("B" & SDI_FIRST & ":R" & SDI_LAST).Value

    msg = RangeGuard(th, TH_LAST, TH_GUARD)
    If Len(msg) = 0 Then msg = RangeGuard(sdi, SDI_LAST, SDI_GUARD)
    If Len(msg) > 0 Then GoTo Abort

    mTag = "TH"
    CheckJointStandard src, TH_FIRST, std, rul, fnd
    CheckLength src, TH_FIRST, fnd
    mTag = "SD"
    CheckJointStandard src2, SDI_FIRST, std, rul, fnd
    CheckLength src2, SDI_FIRST, fnd
    mTag = ""
    CheckReconciliation fnd
""")

a = s.index("' STD.DATA rows 222-227. E=name")
b = s.index("' Rulings the user has entered")
s = s[:a] + '''' Finds the standard table by its own "REF-STD" header instead of fixed
' row numbers, so it survives rows being inserted or deleted anywhere
' above it on STD.DATA. Reads down from the header to the first blank
' REF-STD cell. E=name, G=bolt/stud, H=nut, I=p.washer, J=s.washer,
' K=ref. Read live so a designer edit changes the check.
' errMsg comes back non-empty if the table is missing or empty - the
' caller stops rather than marking every joint UNVERIFIABLE.
Private Function LoadStandard(ByRef sd As Worksheet, ByRef errMsg As String) As Object
    Dim d As Object, r As Long, ref As String
    Dim hdr As Range, headerRow As Long, firstRow As Long, lastRow As Long

    Set d = CreateObject("Scripting.Dictionary")
    errMsg = ""

    Set hdr = sd.Columns(SD_REF_COL).Find(What:=SD_REF_HEADER, LookIn:=xlFormulas, _
                  LookAt:=xlWhole, MatchCase:=False)
    If hdr Is Nothing Then
        errMsg = "Could not find the STD.DATA standard table: no """ & _
            SD_REF_HEADER & """ header found in column K." & vbLf & vbLf & _
            "VERIFICATION cannot check anything until STD.DATA's REF-STD " & _
            "column and header are restored. Nothing has been written."
        Set LoadStandard = d
        Exit Function
    End If

    headerRow = hdr.Row
    firstRow = headerRow + 1
    If Len(Trim$(CStr(sd.Cells(firstRow, SD_REF_COL).Value & ""))) = 0 Then
        lastRow = firstRow - 1
    Else
        lastRow = sd.Cells(firstRow, SD_REF_COL).End(xlDown).Row
    End If

    For r = firstRow To lastRow
        ref = Trim$(CStr(sd.Cells(r, SD_REF_COL).Value & ""))
        If Len(ref) > 0 Then
            d(UCase$(ref)) = Array(Nz(sd.Cells(r, 7).Value), Nz(sd.Cells(r, 8).Value), _
                                   Nz(sd.Cells(r, 9).Value), Nz(sd.Cells(r, 10).Value), _
                                   CStr(sd.Cells(r, 5).Value & ""))
        End If
    Next r

    If d.Count = 0 Then
        errMsg = "The STD.DATA standard table under the """ & SD_REF_HEADER & _
            """ header (row " & headerRow & ") has no entries." & vbLf & vbLf & _
            "VERIFICATION cannot check anything until it is restored. " & _
            "Nothing has been written."
    End If

    Set LoadStandard = d
End Function


''' + s[b:]

rep("""Private Sub CheckJointStandard(ByRef src As Variant, ByRef std As Object, _
                               ByRef rul As Object, ByRef fnd As Collection)""",
"""Private Sub CheckJointStandard(ByRef src As Variant, ByVal firstRow As Long, _
                               ByRef std As Object, ByRef rul As Object, _
                               ByRef fnd As Collection)""")
rep("""        r = TH_FIRST + i - 1
""", """        r = firstRow + i - 1
""", 2)
rep("""Private Sub CheckLength(ByRef src As Variant, ByRef fnd As Collection)""",
"""Private Sub CheckLength(ByRef src As Variant, ByVal firstRow As Long, _
                        ByRef fnd As Collection)""")

rep("""    Dim th As Worksheet, ps As Worksheet, fs As Worksheet
    Dim tTot As Double, pTot As Double, fTot As Double
    Dim r As Long, lastPs As Long

    Set th = ThisWorkbook.Sheets("TANK HARDWARE")""",
"""    Dim th As Worksheet, ps As Worksheet, fs As Worksheet, sdi As Worksheet
    Dim tTot As Double, pTot As Double, fTot As Double
    Dim r As Long, lastPs As Long, lastCell As Range

    Set th = ThisWorkbook.Sheets("TANK HARDWARE")
    Set sdi = ThisWorkbook.Sheets("STD DATA")""")
rep("""            tTot = tTot + Nz(th.Cells(r, 14).Value)
    Next r

    lastPs = ps.Cells(ps.Rows.Count, "F").End(xlUp).Row
""", """            tTot = tTot + Nz(th.Cells(r, 14).Value)
    Next r
    ' PRINT SHEET now carries the STD DATA joints too, so they count here.
    For r = SDI_FIRST To SDI_LAST
        If Len(Trim$(CStr(sdi.Cells(r, 13).Value & ""))) > 0 Then _
            tTot = tTot + Nz(sdi.Cells(r, 14).Value)
    Next r

    ' Find, not End(xlUp): PRINT SHEET is filtered to hide zero rows and
    ' End(xlUp) skips hidden rows, cutting off the bottom of the list.
    Set lastCell = ps.Columns("F").Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastPs = 9 Else lastPs = lastCell.Row
""")
rep('''        AddFinding fnd, 0, "TANK HARDWARE vs PRINT SHEET", "", "", "TOTAL", "ERROR", _''',
    '''        AddFinding fnd, 0, "TANK HARDWARE + STD DATA vs PRINT SHEET", "", "", "TOTAL", "ERROR", _''')

rep("""    fnd.Add Array(r, loc, joint, des, item, sev, txt, gkey, n, nut, pw, sw)""",
    """    fnd.Add Array(r, loc, joint, des, item, sev, txt, gkey, n, nut, pw, sw, mTag)""")
rep("""            If f(0) > 0 Then ws.Cells(r, 2).Value = f(0)""",
    """            If f(0) > 0 Then ws.Cells(r, 2).Value = f(12) & " " & f(0)   ' e.g. "SD 44\"""")

rep("""            If (b(0) < a(0)) Or (b(0) = a(0) And CStr(b(4)) < CStr(a(4))) Then""",
    """            If (RowKey(b) < RowKey(a)) Or (RowKey(b) = RowKey(a) And CStr(b(4)) < CStr(a(4))) Then""")
rep("""Private Sub WriteSummary(""", """' Sort position of a finding: TANK HARDWARE rows first, then STD DATA.
Private Function RowKey(ByVal f As Variant) As Double
    RowKey = IIf(f(12) = "SD", 100000#, 0#) + CDbl(f(0))
End Function


Private Sub WriteSummary(""")

rep("""Private Function RangeGuard(ByRef th As Worksheet) As String""",
"""Private Function RangeGuard(ByRef th As Worksheet, ByVal lastRow As Long, _
                            ByVal guardRow As Long) As String""")
rep("""    For r = TH_LAST + 1 To TH_GUARD""", """    For r = lastRow + 1 To guardRow""")
rep("""        RangeGuard = "There are " & n & " hardware line(s) on TANK HARDWARE below row " & _
            TH_LAST & ", outside the range this check reads." & vbLf & vbLf & _""",
"""        RangeGuard = "There are " & n & " hardware line(s) on " & th.Name & " below row " & _
            lastRow & ", outside the range this check reads." & vbLf & vbLf & _""")

a = s.index("' Turns the rows behind a grouped finding")
b = s.index("Private Function RangeChunk")
s = s[:a] + """' Turns the rows behind a grouped finding into "TH 15-22,28; SD 44-49".
' The collection is in scan order: all TANK HARDWARE rows, then STD DATA.
Private Function RowRanges(ByRef c As Collection) As String
    Dim i As Long, r As Long, startR As Long, prevR As Long
    Dim s As String, f As Variant, tag As String, prevTag As String

    startR = -1: prevR = -1
    For i = 1 To c.Count
        f = c(i)
        r = CLng(f(0))
        tag = CStr(f(12))
        If r > 0 Then
            If startR = -1 Then
                startR = r: prevR = r
                s = s & IIf(Len(s) > 0, "; ", "") & tag & " "
            ElseIf tag <> prevTag Then              ' moved on to the next sheet
                s = s & RangeChunk(startR, prevR) & "; " & tag & " "
                startR = r: prevR = r
            ElseIf r = prevR + 1 Then
                prevR = r
            Else
                s = s & RangeChunk(startR, prevR) & ","
                startR = r: prevR = r
            End If
            prevTag = tag
        End If
    Next i
    If startR <> -1 Then s = s & RangeChunk(startR, prevR)
    RowRanges = s
End Function


""" + s[b:]

open(p, 'w', encoding='latin-1', newline='\r\n').write(s)
print("patched")
