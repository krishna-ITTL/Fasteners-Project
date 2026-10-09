Attribute VB_Name = "VerificationBuilder"
'=======================================================================
' VERIFICATION - checks the hardware list against the standard
'
' Reads TANK HARDWARE (the source with joint type and location text) and
' compares each line against its joint's entry in STD.DATA. Reports only
' deviations. Bulk patterns are grouped to one line; individual outliers
' are listed with their row and location so they can be acted on.
'
' Reports. Never modifies any other sheet.
'=======================================================================

Public Const VS_NAME As String = "VERIFICATION"

' --- source ---
' Data starts at row 15 (row 13 is the header, row 14 a section title).
' Starting at 36 silently skipped the NB flange joints - 20 rows, 168
' pieces - which is exactly what check A caught on its first run.
Public Const TH_FIRST As Long = 15
Public Const TH_LAST  As Long = 250
' Rows past TH_LAST are checked for stray data so the range can never
' silently truncate again.
Public Const TH_GUARD As Long = 320
Public Const SD_FIRST As Long = 222
Public Const SD_LAST  As Long = 227

' --- VERIFICATION sheet layout ---
' Every block starts at column B and stacks top to bottom, so how wide
' the row-range column gets can never push anything off screen.
Public Const RUL_FIRST As Long = 14      ' rulings input rows
Public Const RUL_LAST  As Long = 23
Public Const SUM_FIRST As Long = 27      ' summary rows
Public Const SUM_LAST  As Long = 30
Public Const FND_FIRST As Long = 34      ' findings rows
Public Const FND_LAST  As Long = 183

' A finding is grouped once this many rows share its nature.
Public Const GROUP_MIN As Long = 5

Public gvSilent As Boolean
Public gvLastProblem As String

Public Sub SetVerifySilent(ByVal quiet As Boolean)
    gvSilent = quiet
End Sub

Public Function VerifyLastProblem() As String
    VerifyLastProblem = gvLastProblem
End Function


'-----------------------------------------------------------------------
' Entry point. Assigned to the VERIFY button.
'-----------------------------------------------------------------------
Sub RunVerification()
    Dim ws As Worksheet, th As Worksheet, sd As Worksheet
    Dim std As Object, rul As Object, groups As Object
    Dim src As Variant
    Dim fnd As Collection
    Dim i As Long, r As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As Long, calcSaved As Boolean
    Dim msg As String
    Dim nErr As Long, nInfo As Long, nUnv As Long

    gvLastProblem = ""
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    calcSaved = True

    On Error GoTo Fail

    Set ws = ThisWorkbook.Sheets(VS_NAME)
    Set th = ThisWorkbook.Sheets("TANK HARDWARE")
    Set sd = ThisWorkbook.Sheets("STD.DATA")

    If Application.Calculation <> xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set std = LoadStandard(sd)
    Set rul = LoadRulings(ws)
    Set fnd = New Collection

    ' Columns B..R of TANK HARDWARE. AA:DF are deliberately never read.
    src = th.Range("B" & TH_FIRST & ":R" & TH_LAST).Value

    msg = RangeGuard(th)
    If Len(msg) > 0 Then GoTo Abort

    CheckJointStandard src, std, rul, fnd
    CheckLength src, fnd
    CheckReconciliation fnd

    Set groups = GroupFindings(fnd)

    ClearFindings ws
    r = WriteFindings(ws, groups, nErr, nInfo, nUnv)

    If r < 0 Then
        msg = "The FINDINGS area holds " & (FND_LAST - FND_FIRST + 1) & _
              " rows, which is not enough for this job." & vbLf & vbLf & _
              "Nothing has been written. Make the area larger first."
        GoTo Abort
    End If

    WriteSummary ws, fnd.Count, nErr, nInfo, nUnv, rul.Count

    Application.Calculation = xlCalculationAutomatic
    If prevCalc <> xlCalculationAutomatic Then Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Exit Sub

Abort:
    If calcSaved Then Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    VerifyReport msg
    Exit Sub

Fail:
    If calcSaved Then Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    VerifyReport "Verification failed:" & vbLf & vbLf & _
                 Err.Number & " - " & Err.Description
End Sub


'--------------------------- inputs ------------------------------------

' STD.DATA rows 222-227. E=name, G=bolt/stud, H=nut, I=p.washer,
' J=s.washer, K=ref. Read live so a designer edit changes the check.
Private Function LoadStandard(ByRef sd As Worksheet) As Object
    Dim d As Object, r As Long, ref As String
    Set d = CreateObject("Scripting.Dictionary")
    For r = SD_FIRST To SD_LAST
        ref = Trim$(CStr(sd.Cells(r, 11).Value & ""))
        If Len(ref) > 0 Then
            d(UCase$(ref)) = Array(Nz(sd.Cells(r, 7).Value), Nz(sd.Cells(r, 8).Value), _
                                   Nz(sd.Cells(r, 9).Value), Nz(sd.Cells(r, 10).Value), _
                                   CStr(sd.Cells(r, 5).Value & ""))
        End If
    Next r
    Set LoadStandard = d
End Function


' Rulings the user has entered: joint ref, item, accepted value.
' Ships empty; each line silences exactly the deviation it names.
Private Function LoadRulings(ByRef ws As Worksheet) As Object
    Dim d As Object, r As Long, j As String, it As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For r = RUL_FIRST To RUL_LAST
        j = Trim$(CStr(ws.Cells(r, 2).Value & ""))
        it = Trim$(CStr(ws.Cells(r, 3).Value & ""))
        v = ws.Cells(r, 4).Value
        If Len(j) > 0 And Len(it) > 0 And IsNumeric(v) Then
            d(UCase$(j) & "|" & UCase$(it)) = CDbl(v)
        End If
    Next r
    Set LoadRulings = d
End Function


'--------------------------- checks ------------------------------------

' src is 1-based over columns B..R, so:
'   1=B location   3=D joint   5=F size   10=K required   11=L chosen
'   12=M designation  13=N qty  14=O nut  15=P l.nut  16=Q p.wsh  17=R s.wsh
Private Sub CheckJointStandard(ByRef src As Variant, ByRef std As Object, _
                               ByRef rul As Object, ByRef fnd As Collection)
    Dim i As Long, r As Long
    Dim joint As String, loc As String, des As String
    Dim n As Double, nut As Double, pw As Double, sw As Double
    Dim s As Variant, act As Double, want As Double

    For i = 1 To UBound(src, 1)
        des = Trim$(CStr(src(i, 12) & ""))
        If Len(des) = 0 Then GoTo NextRow
        r = TH_FIRST + i - 1
        joint = Trim$(CStr(src(i, 3) & ""))
        loc = Trim$(CStr(src(i, 1) & ""))
        n = Nz(src(i, 13)): nut = Nz(src(i, 14))
        pw = Nz(src(i, 16)): sw = Nz(src(i, 17))

        If Not std.Exists(UCase$(joint)) Then
            AddFinding fnd, r, loc, IIf(Len(joint) = 0, "(blank)", joint), des, "-", _
                       "UNVERIFIABLE", "joint type not in STD.DATA", _
                       joint & "|-|UNVERIFIABLE|", n, nut, pw, sw
            GoTo NextRow
        End If

        s = std(UCase$(joint))

        If s(0) > 0 Then
            If n = 0 Then
                AddFinding fnd, r, loc, joint, des, "BOLT/STUD", "INFO", _
                           "standard expects a fastener but quantity is 0", _
                           joint & "|BOLT/STUD-ZERO|INFO|", n, nut, pw, sw
                GoTo NextRow
            End If
            CmpRatio fnd, r, loc, joint, des, "NUT", nut / n, _
                     Want2(rul, joint, "NUT", s(1)), "bolt", n, nut, pw, sw
            CmpRatio fnd, r, loc, joint, des, "P.WASHER", pw / n, _
                     Want2(rul, joint, "P.WASHER", s(2)), "bolt", n, nut, pw, sw
            CmpRatio fnd, r, loc, joint, des, "S.WASHER", sw / n, _
                     Want2(rul, joint, "S.WASHER", s(3)), "bolt", n, nut, pw, sw
        Else
            ' Welded pad: no fastener is bought, so a per-bolt ratio is
            ' undefined. Compare the washers against the nut count.
            If n <> 0 Then
                AddFinding fnd, r, loc, joint, des, "BOLT/STUD", "ERROR", _
                           "standard buys no fastener, qty is " & FmtNum(n), _
                           joint & "|BOLT/STUD-PRESENT|ERROR|", n, nut, pw, sw
            End If
            If nut > 0 Then
                CmpRatio fnd, r, loc, joint, des, "P.WASHER", pw / nut, _
                         Want2(rul, joint, "P.WASHER", s(2) / s(1)), "nut", n, nut, pw, sw
                CmpRatio fnd, r, loc, joint, des, "S.WASHER", sw / nut, _
                         Want2(rul, joint, "S.WASHER", s(3) / s(1)), "nut", n, nut, pw, sw
            End If
        End If
NextRow:
    Next i
End Sub


Private Function Want2(ByRef rul As Object, ByVal joint As String, _
                       ByVal item As String, ByVal dflt As Double) As Double
    Dim k As String
    k = UCase$(joint) & "|" & UCase$(item)
    If rul.Exists(k) Then
        Want2 = rul(k)
    Else
        Want2 = dflt
    End If
End Function


Private Sub CmpRatio(ByRef fnd As Collection, ByVal r As Long, ByVal loc As String, _
                     ByVal joint As String, ByVal des As String, ByVal item As String, _
                     ByVal act As Double, ByVal want As Double, ByVal per As String, _
                     ByVal n As Double, ByVal nut As Double, ByVal pw As Double, ByVal sw As Double)
    act = Round(act, 4)
    If act = want Then Exit Sub
    AddFinding fnd, r, loc, joint, des, item, "ERROR", _
               FmtNum(act) & " per " & per & " vs standard " & FmtNum(want), _
               joint & "|" & item & "|ERROR|" & FmtNum(act), n, nut, pw, sw
End Sub


Private Sub CheckLength(ByRef src As Variant, ByRef fnd As Collection)
    Dim i As Long, r As Long, des As String, loc As String
    Dim k As Variant, L As Variant
    Dim re As Object, m As Object, mm As Object
    Dim wantSz As String, wantLn As String

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True

    For i = 1 To UBound(src, 1)
        des = Trim$(CStr(src(i, 12) & ""))
        If Len(des) = 0 Then GoTo NextRow
        r = TH_FIRST + i - 1
        loc = Trim$(CStr(src(i, 1) & ""))
        k = src(i, 10): L = src(i, 11)

        If IsNumeric(k) And IsNumeric(L) Then
            If Not IsEmpty(k) And Not IsEmpty(L) Then
                If CDbl(L) < CDbl(k) Then
                    AddFinding fnd, r, loc, "", des, "LENGTH", "ERROR", _
                        "chosen length " & FmtNum(CDbl(L)) & " shorter than required " & _
                        FmtNum(CDbl(k)), "|LENGTH|ERROR|", 0, 0, 0, 0
                End If
            End If
        End If

        ' Size named in the description vs the designation. INFO only:
        ' the foundation bolts are deliberately +20mm for grouting.
        If Len(loc) > 0 Then
            re.Global = False
            re.Pattern = "^M(\d+)\s*(?:-STUD)?[xX](\d+)"
            If re.Test(des) Then
                Set m = re.Execute(des)(0)
                wantSz = m.SubMatches(0): wantLn = m.SubMatches(1)
                re.Global = True
                re.Pattern = "M(\d+)\s*[xX]\s*(\d+)"
                If re.Test(loc) Then
                    For Each mm In re.Execute(loc)
                        If mm.SubMatches(0) <> wantSz Or mm.SubMatches(1) <> wantLn Then
                            AddFinding fnd, r, Left$(loc, 40), "", des, "TEXT", "INFO", _
                                "description says M" & mm.SubMatches(0) & "x" & mm.SubMatches(1), _
                                "|TEXT|INFO|", 0, 0, 0, 0
                        End If
                    Next mm
                End If
            End If
        End If
NextRow:
    Next i
End Sub


' Guards the FINAL SHEET summary against drift.
Private Sub CheckReconciliation(ByRef fnd As Collection)
    Dim th As Worksheet, ps As Worksheet, fs As Worksheet
    Dim tTot As Double, pTot As Double, fTot As Double
    Dim r As Long, lastPs As Long

    Set th = ThisWorkbook.Sheets("TANK HARDWARE")
    Set ps = ThisWorkbook.Sheets("PRINT SHEET")
    Set fs = ThisWorkbook.Sheets("FINAL SHEET")

    For r = TH_FIRST To TH_LAST
        If Len(Trim$(CStr(th.Cells(r, 13).Value & ""))) > 0 Then _
            tTot = tTot + Nz(th.Cells(r, 14).Value)
    Next r

    lastPs = ps.Cells(ps.Rows.Count, "F").End(xlUp).Row
    For r = 9 To lastPs
        pTot = pTot + Nz(ps.Cells(r, 7).Value)
    Next r

    For r = 37 To 96
        If Not fs.Rows(r).Hidden Then fTot = fTot + Nz(fs.Cells(r, 7).Value)
    Next r
    For r = 99 To 125
        If Not fs.Rows(r).Hidden Then fTot = fTot + Nz(fs.Cells(r, 7).Value)
    Next r

    If tTot <> pTot Then
        AddFinding fnd, 0, "TANK HARDWARE vs PRINT SHEET", "", "", "TOTAL", "ERROR", _
            "totals disagree: " & FmtNum(tTot) & " vs " & FmtNum(pTot), _
            "|TOTAL|ERROR|A1", 0, 0, 0, 0
    End If
    If pTot <> fTot Then
        AddFinding fnd, 0, "PRINT SHEET vs FINAL SHEET", "", "", "TOTAL", "ERROR", _
            "totals disagree: " & FmtNum(pTot) & " vs " & FmtNum(fTot), _
            "|TOTAL|ERROR|A2", 0, 0, 0, 0
    End If
End Sub


'--------------------------- grouping / output -------------------------

Private Sub AddFinding(ByRef fnd As Collection, ByVal r As Long, ByVal loc As String, _
                       ByVal joint As String, ByVal des As String, ByVal item As String, _
                       ByVal sev As String, ByVal txt As String, ByVal gkey As String, _
                       ByVal n As Double, ByVal nut As Double, ByVal pw As Double, ByVal sw As Double)
    fnd.Add Array(r, loc, joint, des, item, sev, txt, gkey, n, nut, pw, sw)
End Sub


Private Function GroupFindings(ByRef fnd As Collection) As Object
    Dim d As Object, i As Long, f As Variant, k As String
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To fnd.Count
        f = fnd(i)
        k = CStr(f(7))
        If Not d.Exists(k) Then d.Add k, New Collection
        d(k).Add f
    Next i
    Set GroupFindings = d
End Function


Private Sub ClearFindings(ByRef ws As Worksheet)
    ws.Rows(FND_FIRST & ":" & FND_LAST).Hidden = False
    ws.Range("B" & FND_FIRST & ":L" & FND_LAST).ClearContents
    ' A plain list like "156,160,173" is otherwise read as a number and
    ' redisplayed with digit grouping. Keep the row column as text.
    ws.Range("B" & FND_FIRST & ":B" & FND_LAST).NumberFormat = "@"
End Sub


' Returns the next free row, or -1 if the area overflowed.
Private Function WriteFindings(ByRef ws As Worksheet, ByRef groups As Object, _
                               ByRef nErr As Long, ByRef nInfo As Long, _
                               ByRef nUnv As Long) As Long
    Dim k As Variant, c As Collection, f As Variant
    Dim r As Long, i As Long, j As Long
    Dim keys As Variant, cnt As Variant, tmpK As Variant, tmpC As Variant

    r = FND_FIRST

    ' Grouped blocks first, largest first.
    If groups.Count > 0 Then
        keys = groups.keys
        ReDim cnt(0 To UBound(keys))
        For i = 0 To UBound(keys)
            cnt(i) = groups(keys(i)).Count
        Next i
        For i = 0 To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If cnt(j) > cnt(i) Then
                    tmpK = keys(i): keys(i) = keys(j): keys(j) = tmpK
                    tmpC = cnt(i): cnt(i) = cnt(j): cnt(j) = tmpC
                End If
            Next j
        Next i

        For i = 0 To UBound(keys)
            Set c = groups(keys(i))
            If c.Count >= GROUP_MIN Then
                If r > FND_LAST Then WriteFindings = -1: Exit Function
                f = c(1)
                ' Name the actual TANK HARDWARE rows, compressed into
                ' ranges - "(grouped)" hid the one thing you need to act.
                ws.Cells(r, 2).Value = RowRanges(c)
                ws.Cells(r, 3).Value = c.Count & " rows"
                ws.Cells(r, 4).Value = f(2)
                ws.Cells(r, 6).Value = f(4)
                ws.Cells(r, 11).Value = f(6)
                ws.Cells(r, 12).Value = f(5)
                Tally CStr(f(5)), c.Count, nErr, nInfo, nUnv
                r = r + 1
            End If
        Next i

        ' Then individual findings, in source-row order.
        Dim flat As Collection
        Set flat = New Collection
        For i = 0 To UBound(keys)
            Set c = groups(keys(i))
            If c.Count < GROUP_MIN Then
                For j = 1 To c.Count
                    flat.Add c(j)
                Next j
            End If
        Next i
        SortByRow flat
        For j = 1 To flat.Count
            If r > FND_LAST Then WriteFindings = -1: Exit Function
            f = flat(j)
            If f(0) > 0 Then ws.Cells(r, 2).Value = f(0)
            ws.Cells(r, 3).Value = f(1)
            ws.Cells(r, 4).Value = f(2)
            ws.Cells(r, 5).Value = f(3)
            ws.Cells(r, 6).Value = f(4)
            If f(8) <> 0 Or f(9) <> 0 Then
                ws.Cells(r, 7).Value = f(8)
                ws.Cells(r, 8).Value = f(9)
                ws.Cells(r, 9).Value = f(10)
                ws.Cells(r, 10).Value = f(11)
            End If
            ws.Cells(r, 11).Value = f(6)
            ws.Cells(r, 12).Value = f(5)
            Tally CStr(f(5)), 1, nErr, nInfo, nUnv
            r = r + 1
        Next j
    End If

    If r <= FND_LAST Then ws.Rows(r & ":" & FND_LAST).Hidden = True
    If r > FND_FIRST Then ws.Rows(FND_FIRST & ":" & (r - 1)).AutoFit
    WriteFindings = r
End Function


Private Sub Tally(ByVal sev As String, ByVal n As Long, ByRef nErr As Long, _
                  ByRef nInfo As Long, ByRef nUnv As Long)
    Select Case sev
        Case "ERROR": nErr = nErr + n
        Case "INFO": nInfo = nInfo + n
        Case Else: nUnv = nUnv + n
    End Select
End Sub


Private Sub SortByRow(ByRef c As Collection)
    Dim i As Long, j As Long, a As Variant, b As Variant
    For i = 1 To c.Count - 1
        For j = i + 1 To c.Count
            a = c(i): b = c(j)
            If (b(0) < a(0)) Or (b(0) = a(0) And CStr(b(4)) < CStr(a(4))) Then
                c.Add b, Before:=i
                c.Remove j + 1
            End If
        Next j
    Next i
End Sub


Private Sub WriteSummary(ByRef ws As Worksheet, ByVal total As Long, _
                         ByVal nErr As Long, ByVal nInfo As Long, _
                         ByVal nUnv As Long, ByVal nRul As Long)
    ws.Range("C9").Value = Now
    ws.Range("C10").Value = IIf(nErr = 0, "PASSED - no errors found", _
                                nErr & " error(s) need attention")
    ws.Cells(SUM_FIRST, 3).Value = nErr
    ws.Cells(SUM_FIRST + 1, 3).Value = nInfo
    ws.Cells(SUM_FIRST + 2, 3).Value = nUnv
    ws.Cells(SUM_FIRST + 3, 3).Value = nRul
End Sub


'--------------------------- small helpers -----------------------------

' Refuse to run if real hardware sits past the scanned range.
Private Function RangeGuard(ByRef th As Worksheet) As String
    Dim r As Long, v As Variant, s As String, n As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^M\d+\s*(?:-STUD)?\s*x\s*\d+\s*-"
    For r = TH_LAST + 1 To TH_GUARD
        v = th.Cells(r, 13).Value
        If Not IsError(v) Then
            s = Trim$(CStr(v & ""))
            If Len(s) > 0 Then
                If re.Test(s) Then n = n + 1
            End If
        End If
    Next r
    If n > 0 Then
        RangeGuard = "There are " & n & " hardware line(s) on TANK HARDWARE below row " & _
            TH_LAST & ", outside the range this check reads." & vbLf & vbLf & _
            "Nothing has been written. The scanned range must be extended first."
    End If
End Function


' Turns the rows behind a grouped finding into "15-22,28,30-34".
' The collection is already in ascending scan order.
Private Function RowRanges(ByRef c As Collection) As String
    Dim i As Long, r As Long, startR As Long, prevR As Long
    Dim s As String, f As Variant

    startR = -1: prevR = -1
    For i = 1 To c.Count
        f = c(i)
        r = CLng(f(0))
        If r > 0 Then
            If startR = -1 Then
                startR = r: prevR = r
            ElseIf r = prevR + 1 Then
                prevR = r
            Else
                s = s & SepComma(s) & RangeChunk(startR, prevR)
                startR = r: prevR = r
            End If
        End If
    Next i
    If startR <> -1 Then s = s & SepComma(s) & RangeChunk(startR, prevR)
    RowRanges = s
End Function


Private Function RangeChunk(ByVal a As Long, ByVal b As Long) As String
    If a = b Then
        RangeChunk = CStr(a)
    Else
        RangeChunk = CStr(a) & "-" & CStr(b)
    End If
End Function


Private Function SepComma(ByVal s As String) As String
    If Len(s) > 0 Then SepComma = ","
End Function


Private Function Nz(ByVal v As Variant) As Double
    If IsEmpty(v) Then
        Nz = 0
    ElseIf IsError(v) Then
        Nz = 0
    ElseIf IsNumeric(v) Then
        Nz = CDbl(v)
    Else
        Nz = 0
    End If
End Function


Private Function FmtNum(ByVal v As Double) As String
    ' Format$(2, "0.####") yields "2." - keep whole numbers clean.
    If v = Int(v) Then
        FmtNum = CStr(CLng(v))
    Else
        FmtNum = Format$(v, "0.####")
    End If
End Function


Private Sub VerifyReport(ByVal msg As String)
    gvLastProblem = msg
    If Not gvSilent Then MsgBox msg, vbExclamation, "VERIFICATION"
End Sub


