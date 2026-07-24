Attribute VB_Name = "WPXDashboard"
' WPXDashboard.bas
' Adds a visible WPX History line/section to the existing "Player Dashboard" sheet
' in GTStatsFINAL.xlsm for the player selected in B3, if that player's roster ID
' is also in the WPX workbook.
'
' It does NOT overwrite any existing GT data on the dashboard or in Arena Power.
' It only writes to rows 8-13, columns B-G, and leaves the rest of the sheet alone.
'
' To use:
'   1. Import this module into GTStatsFINAL.xlsm (Alt+F11).
'   2. Optionally run ImportWPXData() once to embed the WPX sheets as hidden copies.
'   3. Run ShowWPXHistoryOnDashboard() while the Player Dashboard sheet is active
'      and a cross-team player is selected in B3.
'
' To auto-update when B3 changes, add this to the Player Dashboard sheet module:
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
'           Application.EnableEvents = False
'           ShowWPXHistoryOnDashboard
'           Application.EnableEvents = True
'       End If
'   End Sub

Option Explicit

Public Sub ShowWPXHistoryOnDashboard()
    Dim wsDash As Worksheet
    Dim wsGTRoster As Worksheet
    Dim wsWpxRoster As Worksheet
    Dim wsWpxAP As Worksheet
    Dim wbWpx As Workbook

    Set wsDash = ThisWorkbook.Sheets("Player Dashboard")
    Dim player As String
    player = Trim(CStr(Nz(wsDash.Range("B3").Value, "")))
    If player = "" Then
        MsgBox "Select a player in B3 first.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsGTRoster = ThisWorkbook.Sheets("Roster")
    On Error GoTo 0
    If wsGTRoster Is Nothing Then
        MsgBox "Roster sheet not found.", vbExclamation
        Exit Sub
    End If

    Dim pID As Long
    pID = GetRosterID(wsGTRoster, player)
    If pID = 0 Then
        Call ClearWPXArea(wsDash)
        Exit Sub
    End If

    ' Try embedded hidden WPX sheets first; if not present, prompt for the file.
    Dim wpxWasOpen As Boolean
    wpxWasOpen = False
    On Error Resume Next
    Set wsWpxRoster = ThisWorkbook.Sheets("WPX_Roster")
    Set wsWpxAP = ThisWorkbook.Sheets("WPX_ArenaPower")
    On Error GoTo 0

    If wsWpxRoster Is Nothing Or wsWpxAP Is Nothing Then
        Dim f As Variant
        f = Application.GetOpenFilename( _
            "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , _
            "Select WPXStatsFinal.xlsm")
        If VarType(f) = vbBoolean Then Exit Sub
        If f = "False" Then Exit Sub
        Application.ScreenUpdating = False
        Set wbWpx = Workbooks.Open(CStr(f), ReadOnly:=True)
        Set wsWpxRoster = wbWpx.Sheets("Roster")
        Set wsWpxAP = wbWpx.Sheets("Arena Power")
        wpxWasOpen = True
    End If

    Dim wpxName As String, wpxRow As Long
    wpxName = GetWpxNameFromID(wsWpxRoster, pID)
    If wpxName = "" Then
        Call ClearWPXArea(wsDash)
        If wpxWasOpen Then wbWpx.Close SaveChanges:=False
        Exit Sub
    End If

    wpxRow = FindNameRow(wsWpxAP, wpxName)
    If wpxRow = 0 Then
        Call ClearWPXArea(wsDash)
        If wpxWasOpen Then wbWpx.Close SaveChanges:=False
        Exit Sub
    End If

    ' Read WPX current values and baseline
    Dim curDate As String, curLevel As Variant
    Dim curArena As Variant, curHQ As Variant
    Dim baseDate As String, baseLevel As Variant
    Dim baseArena As Variant, baseHQ As Variant

    curDate = FormatDateNz(wsWpxAP.Cells(wpxRow, 2).Value)
    curLevel = Nz(wsWpxAP.Cells(wpxRow, 3).Value, "-")
    curArena = ParsePower(wsWpxAP.Cells(wpxRow, 4).Value)
    curHQ = ParsePower(wsWpxAP.Cells(wpxRow, 5).Value)

    Call FindWpxBaseline(wsWpxAP, wpxRow, baseDate, baseLevel, baseArena, baseHQ)

    ' Write to Player Dashboard rows 8-13 (below the GT top stats, does not touch them)
    Call ClearWPXArea(wsDash)
    With wsDash
        .Range("B8").Value = "WPX HISTORY"
        .Range("B8").Font.Bold = True
        .Range("B8").Font.Color = RGB(255, 200, 0)

        .Range("B9").Value = "Baseline"
        .Range("C9").Value = baseDate
        .Range("D9").Value = "Level " & baseLevel
        .Range("E9").Value = "Arena " & FormatPower(baseArena)
        .Range("F9").Value = "HQ " & FormatPower(baseHQ)

        .Range("B10").Value = "WPX Current"
        .Range("C10").Value = curDate
        .Range("D10").Value = "Level " & curLevel
        .Range("E10").Value = "Arena " & FormatPower(curArena)
        .Range("F10").Value = "HQ " & FormatPower(curHQ)

        .Range("B11").Value = "Level "
        .Range("C11").Value = FormatDeltaNum(SafeDiff(curLevel, baseLevel))
        .Range("D11").Value = "Arena "
        .Range("E11").Value = FormatDelta(CalcDelta(curArena, baseArena))
        .Range("F11").Value = "HQ "
        .Range("G11").Value = FormatDelta(CalcDelta(curHQ, baseHQ))
    End With

    If wpxWasOpen Then
        wbWpx.Close SaveChanges:=False
        Application.ScreenUpdating = True
    End If
End Sub

' Copies WPX Roster and Arena Power sheets into this workbook as hidden sheets
' so ShowWPXHistoryOnDashboard() can use them without prompting every time.
Public Sub ImportWPXData()
    Dim f As Variant
    Dim wbWpx As Workbook

    f = Application.GetOpenFilename( _
        "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , _
        "Select WPXStatsFinal.xlsm")
    If VarType(f) = vbBoolean Then Exit Sub
    If f = "False" Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbWpx = Workbooks.Open(CStr(f), ReadOnly:=True)

    On Error Resume Next
    ThisWorkbook.Sheets("WPX_ArenaPower").Delete
    ThisWorkbook.Sheets("WPX_Roster").Delete
    On Error GoTo 0

    wbWpx.Sheets("Arena Power").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "WPX_ArenaPower"
    ThisWorkbook.Sheets("WPX_ArenaPower").Visible = xlSheetVeryHidden

    wbWpx.Sheets("Roster").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "WPX_Roster"
    ThisWorkbook.Sheets("WPX_Roster").Visible = xlSheetVeryHidden

    wbWpx.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "WPX data imported. Run ShowWPXHistoryOnDashboard to view a player.", vbInformation
End Sub

' Clears the WPX History display area on the Player Dashboard.
Private Sub ClearWPXArea(ws As Worksheet)
    ws.Range("B8:G13").ClearContents
    On Error Resume Next
    ws.Range("B8:G13").Interior.ColorIndex = xlNone
    On Error GoTo 0
End Sub

Private Sub FindWpxBaseline(ws As Worksheet, r As Long, ByRef baseDate As String, _
    ByRef baseLevel As Variant, ByRef baseArena As Variant, ByRef baseHQ As Variant)

    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column

    baseDate = "-"
    baseLevel = "-"
    baseArena = Empty
    baseHQ = Empty

    ' History groups start at column K (11) in groups of 4: Date, Level, Arena, HQ
    For c = 11 To lastCol Step 4
        If Not IsEmpty(ws.Cells(r, c).Value) And CStr(ws.Cells(r, c).Value) <> "" And CStr(ws.Cells(r, c).Value) <> "-" Then
            If IsDate(ws.Cells(r, c).Value) Then
                baseDate = Format(ws.Cells(r, c).Value, "mm/dd/yyyy")
            Else
                baseDate = CStr(ws.Cells(r, c).Value)
            End If
        End If
        If Not IsEmpty(ws.Cells(r, c + 1).Value) And CStr(ws.Cells(r, c + 1).Value) <> "" And CStr(ws.Cells(r, c + 1).Value) <> "-" Then
            baseLevel = ws.Cells(r, c + 1).Value
        End If
        If IsNumeric(ParsePower(ws.Cells(r, c + 2).Value)) Then
            baseArena = ParsePower(ws.Cells(r, c + 2).Value)
        End If
        If IsNumeric(ParsePower(ws.Cells(r, c + 3).Value)) Then
            baseHQ = ParsePower(ws.Cells(r, c + 3).Value)
        End If
    Next c
End Sub

Private Function GetRosterID(ws As Worksheet, player As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(LCase(CStr(Nz(ws.Cells(r, 2).Value, "")))) = LCase(player) Then
            GetRosterID = CLng(Nz(ws.Cells(r, 1).Value, 0))
            Exit Function
        End If
        If Trim(LCase(CStr(Nz(ws.Cells(r, 6).Value, "")))) = LCase(player) Then
            GetRosterID = CLng(Nz(ws.Cells(r, 5).Value, 0))
            Exit Function
        End If
        If Trim(LCase(CStr(Nz(ws.Cells(r, 10).Value, "")))) = LCase(player) Then
            GetRosterID = CLng(Nz(ws.Cells(r, 9).Value, 0))
            Exit Function
        End If
        If Trim(LCase(CStr(Nz(ws.Cells(r, 14).Value, "")))) = LCase(player) Then
            GetRosterID = CLng(Nz(ws.Cells(r, 13).Value, 0))
            Exit Function
        End If
    Next r
    GetRosterID = 0
End Function

Private Function GetWpxNameFromID(ws As Worksheet, pID As Long) As String
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If CLng(Nz(ws.Cells(r, 1).Value, 0)) = pID Then
            GetWpxNameFromID = CStr(Nz(ws.Cells(r, 2).Value, ""))
            Exit Function
        End If
        If CLng(Nz(ws.Cells(r, 5).Value, 0)) = pID Then
            GetWpxNameFromID = CStr(Nz(ws.Cells(r, 6).Value, ""))
            Exit Function
        End If
        If CLng(Nz(ws.Cells(r, 9).Value, 0)) = pID Then
            GetWpxNameFromID = CStr(Nz(ws.Cells(r, 10).Value, ""))
            Exit Function
        End If
        If CLng(Nz(ws.Cells(r, 13).Value, 0)) = pID Then
            GetWpxNameFromID = CStr(Nz(ws.Cells(r, 14).Value, ""))
            Exit Function
        End If
    Next r
    GetWpxNameFromID = ""
End Function

Private Function FindNameRow(ws As Worksheet, name As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(LCase(CStr(Nz(ws.Cells(r, 1).Value, "")))) = LCase(name) Then
            FindNameRow = r
            Exit Function
        End If
    Next r
    FindNameRow = 0
End Function

Private Function ParsePower(v As Variant) As Variant
    If IsEmpty(v) Or v = "" Or v = "-" Then
        ParsePower = Empty
        Exit Function
    End If
    If IsNumeric(v) Then
        ParsePower = CDbl(v)
        Exit Function
    End If

    Dim s As String
    s = CStr(v)
    s = Trim(UCase(s))
    s = Replace(s, "M", "", , , vbTextCompare)
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")

    If IsNumeric(s) Then
        ParsePower = CDbl(s)
    Else
        ParsePower = Empty
    End If
End Function

Private Function FormatPower(v As Variant) As String
    If IsNumeric(v) Then
        FormatPower = Format(v, "0.00") & " M"
    Else
        FormatPower = "-"
    End If
End Function

Private Function CalcDelta(cur As Variant, base As Variant) As Variant
    If IsNumeric(cur) And IsNumeric(base) Then
        CalcDelta = CDbl(cur) - CDbl(base)
    Else
        CalcDelta = Empty
    End If
End Function

Private Function FormatDelta(diff As Variant) As String
    If Not IsNumeric(diff) Then
        FormatDelta = "-"
    ElseIf diff > 0 Then
        FormatDelta = "+" & Format(diff, "0.00") & " M"
    ElseIf diff < 0 Then
        FormatDelta = Format(diff, "0.00") & " M"
    Else
        FormatDelta = "0.00 M"
    End If
End Function

Private Function FormatDeltaNum(diff As Variant) As String
    If Not IsNumeric(diff) Then
        FormatDeltaNum = "-"
    ElseIf diff > 0 Then
        FormatDeltaNum = "+" & CStr(CLng(diff))
    Else
        FormatDeltaNum = CStr(CLng(diff))
    End If
End Function

Private Function SafeDiff(cur As Variant, base As Variant) As Variant
    If IsNumeric(cur) And IsNumeric(base) Then
        SafeDiff = CLng(cur) - CLng(base)
    Else
        SafeDiff = Empty
    End If
End Function

Private Function FormatDateNz(v As Variant) As String
    If IsDate(v) Then
        FormatDateNz = Format(v, "mm/dd/yyyy")
    Else
        FormatDateNz = "-"
    End If
End Function

Private Function Nz(v As Variant, defaultVal As Variant) As Variant
    If IsEmpty(v) Or v = "" Then
        Nz = defaultVal
    Else
        Nz = v
    End If
End Function
