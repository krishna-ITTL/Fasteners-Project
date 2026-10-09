Attribute VB_Name = "FinalSheetBuilder"
'=======================================================================
' FINAL SHEET - self-building procurement summary
'
' Builds the three blocks of FINAL SHEET directly from PRINT SHEET,
' keyed on SIZE + LENGTH + MATERIAL. Replaces the old hand-typed size
' list, which could only ever show combinations someone typed in
' advance and forced MATERIAL to the global E7/E8 values (so an SS
' line such as M12x60-SS could never appear).
'
' Writes SIZE and MATERIAL as values, QUANTITY as a live SUMIF formula,
' then hides every unused row. Values are written into merged cells,
' which is allowed - a spilling dynamic array into merged cells is not.
'=======================================================================

' --- Block geometry. Update here if the sheet layout ever changes. ---
Public Const FS_B1_FIRST As Long = 11
Public Const FS_B1_LAST  As Long = 34
Public Const FS_B2_FIRST As Long = 37
Public Const FS_B2_LAST  As Long = 96
Public Const FS_B3_FIRST As Long = 99
Public Const FS_B3_LAST  As Long = 125

' --- PRINT SHEET source range ---
' PS_MIN_LAST is a floor, not a limit: the real last row is measured
' every run, so the summary keeps working if PRINT SHEET grows.
Public Const PS_FIRST    As Long = 9
Public Const PS_MIN_LAST As Long = 245
Public Const PS_HARD_MAX As Long = 20000

' --- Test seam. When gSilent is True the macro records its message in
'     gLastProblem instead of showing a dialog, so the guard paths can
'     be exercised automatically. Production use leaves gSilent False.
Public gSilent As Boolean
Public gLastProblem As String


' Diagnostics. SetSilentMode True suppresses the dialog so an automated
' check can read the outcome back from LastProblem instead.
Public Sub SetSilentMode(ByVal quiet As Boolean)
    gSilent = quiet
End Sub

Public Function LastProblem() As String
    LastProblem = gLastProblem
End Function


'-----------------------------------------------------------------------
' Back-compat wrapper. The Print macros (PRINTOUTPUT /
' PRINTPROTECTEDOUTPUT, Ctrl+Shift+P / Ctrl+Shift+Q) call this name.
'-----------------------------------------------------------------------
Sub RefreshFinalSheetVisibility()
    Call RebuildFinalSheet
End Sub


'-----------------------------------------------------------------------
' Main entry point. Assigned to the REFRESH button on FINAL SHEET.
'-----------------------------------------------------------------------
Sub RebuildFinalSheet()
    Dim ws As Worksheet, ps As Worksheet
    Dim src As Variant
    Dim i As Long, lastRow As Long
    Dim d1 As Object, d2 As Object, d3 As Object
    Dim re As Object, mc As Object
    Dim desig As String, sz As String, mat As String
    Dim lng As Long, isStud As Boolean
    Dim k As String, cell As Variant
    Dim bad As String, badN As Long
    Dim msg As String
    Dim prevCalc As Long, calcSaved As Boolean
    Dim prevScreen As Boolean, prevEvents As Boolean

    gLastProblem = ""

    ' Capture host state BEFORE arming the handler that restores it,
    ' so the handler can never write back an uninitialised value.
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    calcSaved = True

    On Error GoTo Fail

    Set ws = ThisWorkbook.Sheets("FINAL SHEET")
    Set ps = ThisWorkbook.Sheets("PRINT SHEET")

    ' Make sure PRINT SHEET is current before we read its values.
    If Application.Calculation <> xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic
    End If
    ps.Calculate

    ' Measure the source every run. A hardcoded end row would silently
    ' ignore hardware added past it, under-ordering with no warning.
    ' Bottom of the sheet's used area. PRINT SHEET is filtered to hide zero
    ' rows, and both End(xlUp) and Find skip rows the filter has hidden - a
    ' joint that became non-zero at the bottom of the list (e.g. foundation
    ' bolts) was silently left out. UsedRange ignores the filter. Extra
    ' rows below the list (notes, signatures) have no part name and are
    ' skipped by the loop below.
    lastRow = ps.UsedRange.Row + ps.UsedRange.Rows.Count - 1
    If lastRow < PS_MIN_LAST Then lastRow = PS_MIN_LAST
    If lastRow > PS_HARD_MAX Then
        msg = "PRINT SHEET appears to run to row " & lastRow & _
              ", which is beyond the " & PS_HARD_MAX & " row limit." & _
              vbLf & vbLf & "Nothing has been changed."
        GoTo Abort
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' One bulk read: F=designation, G=qty, H=nut, I=p.washer,
    ' J=s.washer, K=l.nut
    src = ps.Range("F" & PS_FIRST & ":K" & lastRow).Value

    Set d1 = CreateObject("Scripting.Dictionary")   ' size|material -> nut,pw,sw,ln
    Set d2 = CreateObject("Scripting.Dictionary")   ' sizeXlen|mat  -> bolt qty
    Set d3 = CreateObject("Scripting.Dictionary")   ' sizeXlen|mat  -> stud qty

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    ' M12x60-HDG-8.8   /   M12-STUDx60-HDG-8.8
    ' Material is everything after the length separator, so it may
    ' itself contain hyphens (HDG-8.8).
    re.Pattern = "^\s*(M\d+)\s*(-STUD)?\s*x\s*(\d+)\s*-\s*(.+?)\s*$"

    For i = 1 To UBound(src, 1)
        cell = src(i, 1)

        ' IsError must be tested first: concatenating an error value
        ' raises a type mismatch, and VBA's And does not short-circuit.
        If IsError(cell) Then
            badN = badN + 1
            If badN <= 8 Then bad = bad & vbLf & "    row " & _
                (PS_FIRST + i - 1) & ":  (error value)"
        ElseIf Not IsEmpty(cell) Then
            desig = Trim$(CStr(cell))
            If Len(desig) > 0 Then
                If re.Test(desig) Then
                    Set mc = re.Execute(desig)(0)
                    sz = UCase$(mc.SubMatches(0))
                    isStud = (Len(mc.SubMatches(1)) > 0)
                    lng = CLng(mc.SubMatches(2))
                    mat = Trim$(mc.SubMatches(3))

                    ' Block 1 - nuts and washers, by size + material.
                    ' Deliberately includes stud rows whose stud qty is
                    ' 0 (welded pad joints): the studs are not bought,
                    ' but their nuts and washers are.
                    k = sz & "|" & mat
                    If Not d1.Exists(k) Then d1(k) = Array(0#, 0#, 0#, 0#)
                    Dim v As Variant
                    v = d1(k)
                    v(0) = v(0) + Num(src(i, 3))
                    v(1) = v(1) + Num(src(i, 4))
                    v(2) = v(2) + Num(src(i, 5))
                    v(3) = v(3) + Num(src(i, 6))
                    d1(k) = v

                    ' Blocks 2 / 3 - the fastener itself.
                    ' Studs keep their "-STUD" marker in the key:
                    ' without it a stud row would be written as e.g.
                    ' M16X75 and its SUMIF would match the *bolt*
                    ' M16x75 instead.
                    If isStud Then
                        k = sz & "-STUDX" & lng & "|" & mat
                        d3(k) = Num(d3(k)) + Num(src(i, 2))
                    Else
                        k = sz & "X" & lng & "|" & mat
                        d2(k) = Num(d2(k)) + Num(src(i, 2))
                    End If
                Else
                    badN = badN + 1
                    If badN <= 8 Then bad = bad & vbLf & "    row " & _
                        (PS_FIRST + i - 1) & ":  " & desig
                End If
            End If
        End If
    Next i

    ' Report unreadable designations rather than silently dropping them.
    If badN > 0 Then
        msg = "FINAL SHEET could not read " & badN & _
              " designation(s) on PRINT SHEET:" & bad
        If badN > 8 Then msg = msg & vbLf & "    ... and " & _
            (badN - 8) & " more"
        msg = msg & vbLf & vbLf & _
              "Expected format:  M12x60-HDG-8.8  or  M12-STUDx60-HDG-8.8" & _
              vbLf & vbLf & "Nothing has been changed."
        GoTo Abort
    End If

    ' Capacity check before writing anything.
    msg = CapacityMsg("SIZE", CountNonZero1(d1), FS_B1_LAST - FS_B1_FIRST + 1)
    If Len(msg) = 0 Then msg = CapacityMsg("BOLT SIZE", _
        CountNonZero(d2), FS_B2_LAST - FS_B2_FIRST + 1)
    If Len(msg) = 0 Then msg = CapacityMsg("STUD SIZE", _
        CountNonZero(d3), FS_B3_LAST - FS_B3_FIRST + 1)
    If Len(msg) > 0 Then GoTo Abort

    ' Refuse if one material name is a suffix of another - the Block 1
    ' wildcard SUMIF (size & "x*-" & material) would count both.
    msg = SuffixClashMsg(d1)
    If Len(msg) > 0 Then GoTo Abort

    ClearBlocks ws

    WriteBlock1 ws, d1, lastRow
    WriteBlock23 ws, d2, FS_B2_FIRST, FS_B2_LAST, lastRow
    WriteBlock23 ws, d3, FS_B3_FIRST, FS_B3_LAST, lastRow

    Application.Calculation = xlCalculationAutomatic
    ws.Calculate
    If prevCalc <> xlCalculationAutomatic Then Application.Calculation = prevCalc

    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Exit Sub

Abort:
    ' The guards all run before anything is written, so the sheet still
    ' holds the previous, internally consistent list. It is left alone
    ' on purpose - the message says nothing has changed, and a blanked
    ' sheet would be worse to find than a slightly stale one.
    If calcSaved Then Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ReportProblem msg
    Exit Sub

Fail:
    If calcSaved Then Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ReportProblem "FINAL SHEET refresh failed:" & vbLf & vbLf & _
                  Err.Number & " - " & Err.Description
End Sub


'--------------------------- helpers -----------------------------------

Private Sub ReportProblem(ByVal msg As String)
    gLastProblem = msg
    If Not gSilent Then
        MsgBox msg, vbExclamation, "FINAL SHEET"
    End If
End Sub


Private Function Num(ByVal v As Variant) As Double
    If IsEmpty(v) Then
        Num = 0
    ElseIf IsError(v) Then
        Num = 0
    ElseIf IsNumeric(v) Then
        Num = CDbl(v)
    Else
        Num = 0
    End If
End Function


Private Function CountNonZero(ByRef d As Object) As Long
    Dim k As Variant, n As Long
    For Each k In d.keys
        If Num(d(k)) <> 0 Then n = n + 1
    Next k
    CountNonZero = n
End Function


Private Function CountNonZero1(ByRef d As Object) As Long
    Dim k As Variant, v As Variant, n As Long
    For Each k In d.keys
        v = d(k)
        If v(0) <> 0 Or v(1) <> 0 Or v(2) <> 0 Or v(3) <> 0 Then n = n + 1
    Next k
    CountNonZero1 = n
End Function


Private Function CapacityMsg(ByVal blockName As String, _
                             ByVal needed As Long, _
                             ByVal capacity As Long) As String
    If needed > capacity Then
        CapacityMsg = "The " & blockName & " block needs " & needed & _
            " rows but only has room for " & capacity & "." & vbLf & _
            vbLf & "Nothing has been changed. The block needs to be " & _
            "made larger before this job can be summarised."
    End If
End Function


' Block 1 sums with a wildcard on the length (size & "x*-" & material).
' If one material name ends with another - e.g. "SS" and "HDG-SS" -
' that wildcard would match both and over-count. Detect and refuse.
Private Function SuffixClashMsg(ByRef d1 As Object) As String
    Dim k As Variant, mats As Object, a As Variant, b As Variant
    Set mats = CreateObject("Scripting.Dictionary")
    For Each k In d1.keys
        mats(Mid$(k, InStr(k, "|") + 1)) = 1
    Next k
    For Each a In mats.keys
        For Each b In mats.keys
            If a <> b Then
                If Len(a) > Len(b) Then
                    If Right$(UCase$(a), Len(b) + 1) = "-" & UCase$(b) Then
                        SuffixClashMsg = "Cannot summarise safely: the " & _
                          "material """ & b & """ is also the ending of """ & _
                          a & """." & vbLf & vbLf & "The SIZE block totals " & _
                          "would count both. Nothing has been changed."
                        Exit Function
                    End If
                End If
            End If
        Next b
    Next a
End Function


Private Sub ClearBlocks(ByRef ws As Worksheet)
    ws.Rows(FS_B1_FIRST & ":" & FS_B3_LAST).Hidden = False
    ws.Range("D" & FS_B1_FIRST & ":J" & FS_B1_LAST).ClearContents
    ws.Range("D" & FS_B2_FIRST & ":J" & FS_B2_LAST).ClearContents
    ws.Range("D" & FS_B3_FIRST & ":J" & FS_B3_LAST).ClearContents
End Sub


Private Sub WriteBlock1(ByRef ws As Worksheet, ByRef d As Object, _
                        ByVal lastRow As Long)
    Dim keys As Variant, i As Long, r As Long, j As Long
    Dim sz As String, mat As String, v As Variant
    Dim col As Variant, srcCol As Variant

    keys = SortedKeys(d, True)
    r = FS_B1_FIRST
    col = Array("G", "H", "I", "J")
    srcCol = Array("H", "I", "J", "K")

    For i = 0 To UBound2(keys)
        v = d(keys(i))
        If v(0) <> 0 Or v(1) <> 0 Or v(2) <> 0 Or v(3) <> 0 Then
            SplitKey CStr(keys(i)), sz, mat
            ws.Range("D" & r).Value = sz
            ws.Range("F" & r).Value = mat
            For j = 0 To 3
                ws.Range(col(j) & r).Formula = _
                    "=SUMIF('PRINT SHEET'!$F$" & PS_FIRST & ":$F$" & lastRow & _
                    ",$D" & r & "&""x*-""&$F" & r & _
                    ",'PRINT SHEET'!$" & srcCol(j) & "$" & PS_FIRST & _
                    ":$" & srcCol(j) & "$" & lastRow & ")" & _
                    "+SUMIF('PRINT SHEET'!$F$" & PS_FIRST & ":$F$" & lastRow & _
                    ",$D" & r & "&""-STUDx*-""&$F" & r & _
                    ",'PRINT SHEET'!$" & srcCol(j) & "$" & PS_FIRST & _
                    ":$" & srcCol(j) & "$" & lastRow & ")"
            Next j
            r = r + 1
        End If
    Next i

    If r <= FS_B1_LAST Then ws.Rows(r & ":" & FS_B1_LAST).Hidden = True
End Sub


Private Sub WriteBlock23(ByRef ws As Worksheet, ByRef d As Object, _
                         ByVal firstRow As Long, ByVal lastBlockRow As Long, _
                         ByVal lastRow As Long)
    Dim keys As Variant, i As Long, r As Long
    Dim sz As String, mat As String

    keys = SortedKeys(d, False)
    r = firstRow

    For i = 0 To UBound2(keys)
        If Num(d(keys(i))) <> 0 Then
            SplitKey CStr(keys(i)), sz, mat
            ws.Range("D" & r).Value = sz
            ws.Range("F" & r).Value = mat
            ws.Range("G" & r).Formula = _
                "=SUMIF('PRINT SHEET'!$F$" & PS_FIRST & ":$F$" & lastRow & _
                ",$D" & r & "&""-""&$F" & r & _
                ",'PRINT SHEET'!$G$" & PS_FIRST & ":$G$" & lastRow & ")"
            r = r + 1
        End If
    Next i

    If r <= lastBlockRow Then ws.Rows(r & ":" & lastBlockRow).Hidden = True
End Sub


Private Sub SplitKey(ByVal k As String, ByRef sz As String, _
                     ByRef mat As String)
    Dim p As Long
    p = InStr(k, "|")
    sz = Left$(k, p - 1)
    mat = Mid$(k, p + 1)
End Sub


Private Function UBound2(ByVal a As Variant) As Long
    On Error Resume Next
    UBound2 = -1
    UBound2 = UBound(a)
End Function


' Sort by numeric size, then numeric length, then material.
' A text sort would put M10 before M4.
Private Function SortedKeys(ByRef d As Object, _
                            ByVal sizeOnly As Boolean) As Variant
    Dim keys As Variant, rank As Variant
    Dim i As Long, j As Long, tmp As Variant, n As Long

    If d.Count = 0 Then
        SortedKeys = Array()
        Exit Function
    End If

    keys = d.keys
    n = UBound(keys)
    ReDim rank(0 To n)
    For i = 0 To n
        rank(i) = SortRank(CStr(keys(i)), sizeOnly)
    Next i

    For i = 0 To n - 1
        For j = i + 1 To n
            If rank(j) < rank(i) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
                tmp = rank(i): rank(i) = rank(j): rank(j) = tmp
            End If
        Next j
    Next i

    SortedKeys = keys
End Function


' Returns a sortable string: size zero-padded, length zero-padded,
' then material.
Private Function SortRank(ByVal k As String, _
                          ByVal sizeOnly As Boolean) As String
    Dim sz As String, mat As String, body As String
    Dim nSize As Long, nLen As Long, p As Long

    SplitKey k, body, mat

    If sizeOnly Then
        nSize = CLng(Mid$(body, 2))
        nLen = 0
    Else
        p = InStr(2, body, "X")
        If p = 0 Then
            nSize = CLng(Mid$(body, 2))
            nLen = 0
        Else
            sz = Mid$(body, 2, p - 2)
            ' strip a "-STUD" suffix off the size part
            If InStr(sz, "-") > 0 Then sz = Left$(sz, InStr(sz, "-") - 1)
            nSize = CLng(sz)
            nLen = CLng(Mid$(body, p + 1))
        End If
    End If

    SortRank = Format$(nSize, "000") & Format$(nLen, "0000") & mat
End Function


