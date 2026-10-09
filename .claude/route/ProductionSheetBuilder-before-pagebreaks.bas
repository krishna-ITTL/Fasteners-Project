Attribute VB_Name = "ProductionSheetBuilder"
'=======================================================================
' PRODUCTION SHEET - shopfloor list, one block per fastener part
'
' WHAT IT DOES
'   Reads every hardware row on PRINT SHEET and groups them by part
'   (for example "M12x50-HDG"). For each part it writes a block:
'
'       M12X50-HDG                                   <- yellow title bar
'       JOINT LOCATION | BOLT | NUT | P.W | S.W | L.NUT | STD.REF
'       25NB GATE VALVE ...  128   128   256    0     0    SD-268
'       NO OF 25NB FLANGE    152   152   304    0     0    SD-268
'       TOTAL                280   280   560    0     0
'
'   so the person fitting M12x50 bolts sees every place they go.
'
' WHY PRINT SHEET AND NOT TANK HARDWARE
'   Since Oct 2026 the inputs are split: STD DATA holds the standard
'   joints (valves, plates, flange joints) and TANK HARDWARE holds the
'   tank-specific ones. PRINT SHEET is the only sheet that has BOTH,
'   and it is also what FINAL SHEET reads - so the two sheets always
'   agree with each other.
'
' HOW IT IS RUN
'   The REFRESH button on PRINT SHEET runs RefreshAll (bottom of this
'   module), which rebuilds FINAL SHEET and then PRODUCTION SHEET.
'   The print macros (Ctrl+Shift+P / Ctrl+Shift+Q) also rebuild it
'   when PRODUCTION SHEET is the sheet being printed.
'
' WHY A MACRO AND NOT FORMULAS
'   The number of blocks and rows changes from job to job, and the
'   layout uses merged cells. Excel formulas cannot add or remove
'   blocks, and "spilling" formulas refuse to write into merged cells.
'   So the macro wipes the old blocks and writes fresh ones each time.
'=======================================================================

Option Explicit     ' every variable must be declared with Dim - catches typos

' --- Where things are. Change these if the sheet layout ever moves. ---
Private Const PROD_SHEET   As String = "PRODUCTION SHEET"
Private Const SRC_SHEET    As String = "PRINT SHEET"
Private Const PD_FIRST_ROW As Long = 7       ' first row the blocks may use;
                                             ' rows 2-6 are the fixed header
Private Const PD_TITLE_ROWS As String = "$2:$6"   ' repeated on every printed page

' PRINT SHEET: hardware starts on row 9. The last row is measured each
' run (see LastSourceRow) so new rows added by the designers are never
' missed. PS_MIN_LAST is only a floor; PS_HARD_MAX is a safety stop.
Private Const PS_FIRST     As Long = 9
Private Const PS_MIN_LAST  As Long = 245
Private Const PS_HARD_MAX  As Long = 20000

' --- Test seam: when gPdSilent is True, problems are stored in
'     gPdLastProblem instead of popping up a message box, so an
'     automated test can run the macro and read what happened.
'     Normal use leaves it False. ---
Public gPdSilent As Boolean
Public gPdLastProblem As String

Public Sub SetProductionSilent(ByVal quiet As Boolean)
    gPdSilent = quiet
End Sub

Public Function ProductionLastProblem() As String
    ProductionLastProblem = gPdLastProblem
End Function


'-----------------------------------------------------------------------
' RefreshAll - the ONE button the user presses (on PRINT SHEET).
' Rebuilds FINAL SHEET first, then PRODUCTION SHEET. Each sheet is
' unprotected for the rebuild and protected again afterwards if it was
' protected before (the "print protected" macro leaves sheets locked,
' and a locked sheet cannot be written to).
'-----------------------------------------------------------------------
Sub RefreshAll()
    Dim wasSilent As Boolean, wasFinalSilent As Boolean
    Dim p1 As String, p2 As String, msg As String

    ' While both sheets rebuild, keep their own pop-ups quiet and collect
    ' what each one has to say, so the user gets ONE message at the end.
    wasSilent = gPdSilent
    wasFinalSilent = gSilent            ' FINAL SHEET's own quiet switch
    gPdSilent = True
    SetSilentMode True                  ' FINAL SHEET's quiet switch

    RebuildSheetSafely "FINAL SHEET", "RebuildFinalSheet", p1
    RebuildSheetSafely PROD_SHEET, "RebuildProductionSheet", p2

    SetSilentMode wasFinalSilent
    gPdSilent = wasSilent

    If Len(p1) = 0 And Len(p2) = 0 Then
        gPdLastProblem = ""
        If Not gPdSilent Then MsgBox "FINAL SHEET and PRODUCTION SHEET are up to date.", _
                                     vbInformation, "REFRESH"
    Else
        If Len(p1) > 0 Then msg = "FINAL SHEET:" & vbLf & p1
        If Len(p2) > 0 Then
            If Len(msg) > 0 Then msg = msg & vbLf & vbLf
            msg = msg & "PRODUCTION SHEET:" & vbLf & p2
        End If
        ReportPd msg
    End If
End Sub


'-----------------------------------------------------------------------
' RebuildSheetSafely - unlock a sheet, run its rebuild macro, re-lock.
'   sheetName : the sheet that will be rewritten
'   macroName : the macro that rewrites it (run by name)
'   problem   : (optional) comes back holding the problem message, or ""
' Returns True if the sheet was rebuilt without a problem. The print
' macros in Module1 use that to avoid printing an out-of-date sheet.
'-----------------------------------------------------------------------
Public Function RebuildSheetSafely(ByVal sheetName As String, _
                                   ByVal macroName As String, _
                                   Optional ByRef problem As String) As Boolean
    Dim ws As Worksheet
    Dim wasLocked As Boolean

    problem = ""
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets(sheetName)
    wasLocked = ws.ProtectContents

    If wasLocked Then
        ' Password:="" means "try with no password". If the sheet HAS a
        ' password this raises an error instead of popping up a password
        ' box (a box would let someone type it, and then the re-lock below
        ' would quietly drop their password).
        On Error Resume Next
        ws.Unprotect Password:=""
        On Error GoTo Fail
        If ws.ProtectContents Then
            problem = sheetName & " is protected with a password, so it " & _
                      "cannot be refreshed." & vbLf & _
                      "Unprotect it (Review > Unprotect Sheet) and try again."
            GoTo Done
        End If
    End If

    Application.Run macroName           ' run the rebuild macro by its name

    ' Ask the builder whether it had a problem (it stores its message).
    If macroName = "RebuildFinalSheet" Then
        problem = LastProblem()          ' from FinalSheetBuilder
    Else
        problem = gPdLastProblem
    End If

    GoTo Done

Fail:
    problem = sheetName & " refresh failed: " & Err.Number & " - " & Err.Description
Done:
    ' Lock it again (same options PRINTPROTECTEDOUTPUT uses) - also after
    ' a failure, so an error can never leave the sheet unlocked.
    If wasLocked And Not ws Is Nothing Then
        If Not ws.ProtectContents Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    ' Called on its own (from a print macro) the message is shown here;
    ' inside RefreshAll it is quiet and RefreshAll shows one combined one.
    ' A builder that already showed its own message is not repeated.
    If Len(problem) > 0 And Not gPdSilent Then
        If problem <> LastProblem() And problem <> gPdLastProblem Then
            MsgBox problem, vbExclamation, sheetName
        End If
    End If
    RebuildSheetSafely = (Len(problem) = 0)
End Function


'-----------------------------------------------------------------------
' RebuildProductionSheet - the main macro.
' Steps:  1 read PRINT SHEET   2 group rows by part   3 check for bad
'         part names   4 sort the parts   5 wipe old blocks
'         6 write new blocks   7 set the print area
' Steps 1-4 change nothing, so if a check fails the sheet is left
' exactly as it was.
'-----------------------------------------------------------------------
Sub RebuildProductionSheet()
    Dim ws As Worksheet, ps As Worksheet
    Dim lastRow As Long, r As Long, i As Long, c As Long
    Dim desig As Variant, key As String
    Dim lastWritten As Long
    Dim qty(1 To 5) As Double, anyQty As Boolean
    Dim groups As Object            ' part name -> Collection of rows
    Dim rec As Variant              ' one output row: Array(location, 5 qtys, std.ref)
    Dim keys() As String
    Dim re As Object
    Dim bad As String, badN As Long
    Dim msg As String
    Dim outRow As Long
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As Long

    gPdLastProblem = ""

    ' Remember Excel's settings so we can put them back at the end,
    ' even if something goes wrong (see the Fail: label).
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo Fail              ' any unexpected error jumps to Fail:

    Set ws = ThisWorkbook.Sheets(PROD_SHEET)
    Set ps = ThisWorkbook.Sheets(SRC_SHEET)

    ' Make sure every formula on PRINT SHEET has its latest value.
    If Application.Calculation <> xlCalculationAutomatic Then
        Application.Calculation = xlCalculationAutomatic
    End If
    ps.Calculate

    lastRow = LastSourceRow(ps)
    If lastRow > PS_HARD_MAX Then
        msg = "PRINT SHEET appears to run to row " & lastRow & _
              ", beyond the " & PS_HARD_MAX & " row limit." & _
              vbLf & vbLf & "Nothing has been changed."
        GoTo Abort
    End If

    ' Turn off screen redraw while we work - much faster, no flicker.
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' The pattern a part name must follow:
    '   M12x50-HDG        (bolt)
    '   M12-STUDx75-HDG   (stud)
    ' (M\d+) = size, (-STUD)? = optional stud marker, (\d+) = length,
    ' (.+?) = material (may contain hyphens, e.g. HDG-8.8).
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Pattern = "^\s*(M\d+)\s*(-STUD)?\s*x\s*(\d+)\s*-\s*(.+?)\s*$"

    ' A Dictionary is a lookup table: key = part name, value = the list
    ' (Collection) of rows that use that part.
    Set groups = CreateObject("Scripting.Dictionary")

    '--- Steps 1 + 2: read each PRINT SHEET row and file it under its part
    For r = PS_FIRST To lastRow
        desig = ps.Cells(r, "F").Value          ' F = part name

        If IsError(desig) Then
            ' A #REF! / #N/A in the part name: remember it, report later.
            badN = badN + 1
            If badN <= 8 Then bad = bad & vbLf & "    row " & r & ":  (error value)"

        ElseIf Len(Trim$(CStr(desig))) > 0 Then
            ' Read the five quantities: G=bolt/stud, H=nut, I=p.washer,
            ' J=s.washer, K=l.nut. Skip the row if they are all zero -
            ' nothing is fitted there on this job.
            anyQty = False
            For c = 1 To 5
                qty(c) = Num(ps.Cells(r, 6 + c).Value)
                If qty(c) <> 0 Then anyQty = True
            Next c

            If anyQty Then
                If Not re.Test(CStr(desig)) Then
                    badN = badN + 1
                    If badN <= 8 Then bad = bad & vbLf & "    row " & r & ":  " & desig
                Else
                    ' Rebuild the name from its pieces, so small typing
                    ' differences ("M12 x50-hdg", "M12x050-HDG") still land in
                    ' the same block - the same way FINAL SHEET groups them.
                    ' Material stays in the name on purpose: an SS bolt must
                    ' never be added to the HDG bolts of the same size.
                    key = PartKey(re, CStr(desig))
                    If Not groups.Exists(key) Then groups.Add key, New Collection

                    rec = Array(JointLocation(ps, r), qty(1), qty(2), qty(3), _
                                qty(4), qty(5), ps.Cells(r, "L").Value)
                    groups(key).Add rec
                End If
            End If
        End If
    Next r

    '--- Step 3: refuse to continue if any part name could not be read
    If badN > 0 Then
        msg = "PRODUCTION SHEET could not read " & badN & _
              " part name(s) on PRINT SHEET:" & bad
        If badN > 8 Then msg = msg & vbLf & "    ... and " & (badN - 8) & " more"
        msg = msg & vbLf & vbLf & _
              "Expected format:  M12x50-HDG  or  M12-STUDx75-HDG" & _
              vbLf & vbLf & "Nothing has been changed."
        GoTo Abort
    End If

    '--- Step 4: put the parts in shopfloor order
    keys = SortedPartKeys(groups, re)

    '--- Step 5: wipe the old blocks (row 7 downwards). Rows 2-6, the
    '    header with the logo, are never touched.
    ClearOldBlocks ws

    '--- Step 6: write one block per part
    outRow = PD_FIRST_ROW
    If groups.Count > 0 Then
        For i = 0 To UBound(keys)
            outRow = WriteBlock(ws, keys(i), groups(keys(i)), outRow)
            outRow = outRow + 1             ' one empty row between blocks
        Next i
    Else
        ws.Cells(outRow, "D").Value = "No hardware quantities on PRINT SHEET."
        outRow = outRow + 1
    End If

    '--- Step 7: print setup - print only what we wrote, header on every page
    lastWritten = outRow - 1                 ' last row with something on it
    If groups.Count > 0 Then lastWritten = outRow - 2   ' drop the spacer row
    ws.PageSetup.PrintArea = "$D$2:$K$" & lastWritten
    ws.PageSetup.PrintTitleRows = PD_TITLE_ROWS
    ' D:K is wider than an A4 page; without this the STD.REF column
    ' prints on a separate sheet of paper. Fit the width, any height.
    With ws.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    Application.Calculation = xlCalculationAutomatic
    ws.Calculate                            ' work out the TOTAL formulas
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Exit Sub

Abort:
    ' A check failed before anything was written. Put Excel back and say why.
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ReportPd msg
    Exit Sub

Fail:
    ' Something unexpected went wrong. Put Excel back and show the error.
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ReportPd "PRODUCTION SHEET refresh failed:" & vbLf & vbLf & _
             Err.Number & " - " & Err.Description
End Sub


'--------------------------- helpers -----------------------------------

' Last row on PRINT SHEET that has a part name in column F.
' Never less than PS_MIN_LAST.
Private Function LastSourceRow(ByRef ps As Worksheet) As Long
    LastSourceRow = ps.Cells(ps.Rows.Count, "F").End(xlUp).Row
    If LastSourceRow < PS_MIN_LAST Then LastSourceRow = PS_MIN_LAST
End Function


' The joint location for row r. On PRINT SHEET one location can cover
' several rows (the cell is merged downwards, e.g. "OLTC CONSERVATOR"
' spans 5 part rows). Only the top cell of a merged area holds the text,
' so we always read the top-left cell of the merged area.
' For a cell that is not merged, MergeArea is just the cell itself.
Private Function JointLocation(ByRef ps As Worksheet, ByVal r As Long) As String
    Dim v As Variant
    v = ps.Cells(r, "D").MergeArea.Cells(1, 1).Value
    If IsError(v) Then
        JointLocation = "(error in PRINT SHEET row " & r & ")"
    Else
        JointLocation = Trim$(CStr(v))
    End If
End Function


' Turn a cell value into a number. Blank, text or error counts as 0.
Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Then
        Num = 0
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        Num = CDbl(v)
    Else
        Num = 0
    End If
End Function


' Wipe everything from row 7 to the bottom of the used area:
' un-merge first (Excel cannot clear half of a merged cell), then clear
' values AND formatting, then put row heights back to normal.
Private Sub ClearOldBlocks(ByRef ws As Worksheet)
    Dim lastUsed As Long
    Dim rng As Range

    lastUsed = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastUsed < PD_FIRST_ROW Then lastUsed = PD_FIRST_ROW

    Set rng = ws.Range(ws.Cells(PD_FIRST_ROW, "B"), ws.Cells(lastUsed, "L"))
    rng.UnMerge
    rng.Clear                           ' values + formats + borders
    rng.EntireRow.RowHeight = ws.StandardHeight
End Sub


' Write one part block starting at row r. Returns the row after the
' TOTAL row. Layout copied from the designers' template:
'   title row  : D:K merged, size 16, yellow
'   header row : D:E merged, bold, centred
'   data rows  : D:E merged location, F..K centred numbers
'   TOTAL row  : F..J = SUM of the data rows (a live formula)
Private Function WriteBlock(ByRef ws As Worksheet, ByVal partKey As String, _
                           ByRef items As Collection, ByVal r As Long) As Long
    Dim firstRow As Long, dataFirst As Long, dataLast As Long
    Dim rec As Variant, c As Long
    Dim heads As Variant

    firstRow = r

    ' --- Title bar: the part name
    With ws.Range(ws.Cells(r, "D"), ws.Cells(r, "K"))
        .Merge
        .Value = partKey
        .Font.Size = 16
        .Interior.Color = RGB(255, 255, 0)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 24.95
    r = r + 1

    ' --- Column headings. A stud block says STUD instead of BOLT.
    If InStr(partKey, "-STUD") > 0 Then
        heads = Array("STUD", "NUT", "P. WASHER", "S. WASHER", "L.NUT", "STD.REF")
    Else
        heads = Array("BOLT", "NUT", "P. WASHER", "S. WASHER", "L.NUT", "STD.REF")
    End If
    ws.Range(ws.Cells(r, "D"), ws.Cells(r, "E")).Merge
    ws.Cells(r, "D").Value = "JOINT LOCATION"
    For c = 0 To 5
        ws.Cells(r, 6 + c).Value = heads(c)       ' column F is number 6
    Next c
    With ws.Range(ws.Cells(r, "D"), ws.Cells(r, "K"))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    r = r + 1

    ' --- One row per joint that uses this part
    dataFirst = r
    For Each rec In items
        ws.Range(ws.Cells(r, "D"), ws.Cells(r, "E")).Merge
        ws.Cells(r, "D").Value = rec(0)
        ws.Cells(r, "D").HorizontalAlignment = xlLeft
        ' Long names (e.g. "100NB GATE VALVE ASSEMBLY (BOTTOM/QUICK DRAIN)")
        ' are shrunk to fit the cell instead of being cut off on paper.
        ws.Cells(r, "D").ShrinkToFit = True
        For c = 1 To 5
            ws.Cells(r, 5 + c).Value = rec(c)     ' F..J quantities
        Next c
        ws.Cells(r, "K").Value = rec(6)           ' STD.REF
        ws.Range(ws.Cells(r, "F"), ws.Cells(r, "K")).HorizontalAlignment = xlCenter
        r = r + 1
    Next rec
    dataLast = r - 1

    ' --- TOTAL row: live SUM formulas, so a hand-edit still adds up
    ws.Range(ws.Cells(r, "D"), ws.Cells(r, "E")).Merge
    ws.Cells(r, "D").Value = "TOTAL"
    For c = 6 To 10                               ' F..J
        ws.Cells(r, c).Formula = "=SUM(" & _
            ws.Cells(dataFirst, c).Address(False, False) & ":" & _
            ws.Cells(dataLast, c).Address(False, False) & ")"
    Next c
    With ws.Range(ws.Cells(r, "D"), ws.Cells(r, "K"))
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 20.1

    ' --- Thin borders round every cell, text centred top-to-bottom
    ws.Range(ws.Cells(firstRow, "D"), ws.Cells(r, "K")).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(firstRow, "D"), ws.Cells(r, "K")).Borders.LineStyle = xlContinuous

    WriteBlock = r + 1
End Function


' The block name for a part, built from its pieces:
'   size + "-STUD" (studs only) + "X" + length + "-" + material
' e.g. "m12 x050-HDG" -> "M12X50-HDG". Only called after re.Test passed.
' The material is kept exactly as typed (only trimmed), because that is
' how FINAL SHEET groups it - so both sheets always split parts the same way.
Private Function PartKey(ByRef re As Object, ByVal desig As String) As String
    Dim m As Object
    Set m = re.Execute(desig)(0)
    PartKey = UCase$(m.SubMatches(0)) & IIf(Len(m.SubMatches(1)) > 0, "-STUD", "") & _
              "X" & CLng(m.SubMatches(2)) & "-" & Trim$(m.SubMatches(3))
End Function


' Return the part names in shopfloor order:
'   all bolts first, then all studs;
'   inside each: size (M4, M5 ... M24), then length, then material.
' Each name gets a sortable code, e.g. bolt M8x40-HDG -> "0|0008|00040|HDG",
' so that M8 comes before M10 (plain text sorting would put M10 first).
Private Function SortedPartKeys(ByRef groups As Object, ByRef re As Object) As String()
    Dim keys() As String, codes() As String
    Dim n As Long, i As Long, j As Long
    Dim k As Variant, m As Object
    Dim tmp As String

    n = groups.Count
    If n = 0 Then
        ReDim keys(0)
        SortedPartKeys = keys
        Exit Function
    End If
    ReDim keys(n - 1)
    ReDim codes(n - 1)

    i = 0
    For Each k In groups.keys
        Set m = re.Execute(k)(0)        ' every key already passed re.Test
        keys(i) = k
        codes(i) = IIf(Len(m.SubMatches(1)) > 0, "1", "0") & "|" & _
                   Format$(CLng(Mid$(m.SubMatches(0), 2)), "0000") & "|" & _
                   Format$(CLng(m.SubMatches(2)), "00000") & "|" & _
                   m.SubMatches(3)
        i = i + 1
    Next k

    ' Simple insertion sort - there are only a few dozen parts.
    For i = 1 To n - 1
        For j = i To 1 Step -1
            If codes(j - 1) <= codes(j) Then Exit For
            tmp = codes(j): codes(j) = codes(j - 1): codes(j - 1) = tmp
            tmp = keys(j):  keys(j) = keys(j - 1):   keys(j - 1) = tmp
        Next j
    Next i

    SortedPartKeys = keys
End Function


' Show a problem to the user - or, in test mode, just store it.
Private Sub ReportPd(ByVal msg As String)
    gPdLastProblem = msg
    If Not gPdSilent Then MsgBox msg, vbExclamation, "PRODUCTION SHEET"
End Sub
