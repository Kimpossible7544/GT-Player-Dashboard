Attribute VB_Name = "WPXDashboard"
' WPXDashboard.bas
' Adds a visible WPX data row and WPX History section to the existing "Player Dashboard" sheet
' in GTStatsFINAL.xlsm for the player selected in B3, resolving WPX identity through
' the WPX ID Map, WPX roster ID, archived WPX ID, then the GT roster AKA name.
'
' It does NOT overwrite any existing GT data on the dashboard or in Arena Power.
' It writes the WPX overall row at A7:F7 and the WPX History section at rows 9-12.
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
    Dim wsWpxArchived As Worksheet
    Dim wsWpxAP As Worksheet
    Dim wsWpxPT As Worksheet
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

    Dim pID As String
    pID = GetRosterID(wsGTRoster, player)
    If pID = "" Or pID = "0" Then
        Call ClearWPXArea(wsDash)
        Exit Sub
    End If

    ' Try embedded hidden WPX sheets first; if not present, prompt for the file.
    Dim wpxWasOpen As Boolean
    wpxWasOpen = False
    On Error Resume Next
    Set wsWpxRoster = ThisWorkbook.Sheets("WPX_Roster")
    Set wsWpxArchived = ThisWorkbook.Sheets("WPX_Archived")
    Set wsWpxAP = ThisWorkbook.Sheets("WPX_ArenaPower")
    Set wsWpxPT = ThisWorkbook.Sheets("WPX_PlayerTracking")
    On Error GoTo 0

    If wsWpxRoster Is Nothing Or wsWpxAP Is Nothing Or wsWpxPT Is Nothing Then
        Dim f As Variant
        f = Application.GetOpenFilename( _
            "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , _
            "Select WPXStatsFinal.xlsm")
        If VarType(f) = vbBoolean Then Exit Sub
        If f = "False" Then Exit Sub
        Application.ScreenUpdating = False
        Set wbWpx = Workbooks.Open(CStr(f), ReadOnly:=True)
        Set wsWpxRoster = wbWpx.Sheets("Roster")
        On Error Resume Next
        Set wsWpxArchived = wbWpx.Sheets("Archived Players")
        On Error GoTo 0
        Set wsWpxAP = wbWpx.Sheets("Arena Power")
        Set wsWpxPT = wbWpx.Sheets("Player Tracking")
        wpxWasOpen = True
    End If

    Dim wpxName As String, wpxRow As Long, wpxPTRow As Long
    wpxName = ResolveWpxName(wsGTRoster, wsWpxRoster, wsWpxArchived, pID)
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

    wpxPTRow = FindNameRow(wsWpxPT, wpxName)
    Dim wpxOverallRank As Variant, wpxOverallScore As Variant
    Dim wpxDailyAvg As Variant, wpxMissedDaily As Variant, wpxMissedWeekly As Variant
    If wpxPTRow > 0 Then
        wpxOverallRank = Nz(wsWpxPT.Cells(wpxPTRow, 5).Value, "-")
        wpxOverallScore = Nz(wsWpxPT.Cells(wpxPTRow, 3).Value, 0)
        wpxDailyAvg = Nz(wsWpxPT.Cells(wpxPTRow, 7).Value, 0)
        wpxMissedDaily = Nz(wsWpxPT.Cells(wpxPTRow, 8).Value, 0)
        wpxMissedWeekly = Nz(wsWpxPT.Cells(wpxPTRow, 9).Value, 0)
    Else
        wpxOverallRank = "-"
        wpxOverallScore = 0
        wpxDailyAvg = 0
        wpxMissedDaily = 0
        wpxMissedWeekly = 0
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

    ' Write WPX overall row at A7:F7 (same order as GT top stats, marked with WPX tag)
    ' and WPX History section at rows 9-12. Existing GT top stats in rows 5-6 untouched.
    Call ClearWPXArea(wsDash)
    With wsDash
        .Range("A7").Value = "WPX"
        .Range("A7").Font.Bold = True
        .Range("A7").Font.Color = RGB(255, 200, 0)

        .Range("B7").Value = wpxOverallRank
        .Range("C7").Value = wpxOverallScore
        .Range("C7").NumberFormat = "#,##0"
        .Range("D7").Value = wpxDailyAvg
        .Range("D7").NumberFormat = "#,##0.0"
        .Range("E7").Value = wpxMissedDaily
        .Range("F7").Value = wpxMissedWeekly

        ' Apply amber font to the WPX overall row to mark it as legacy WPX data
        .Range("A7:F7").Font.Name = .Range("B6").Font.Name
        .Range("A7:F7").Font.Size = .Range("B6").Font.Size
        .Range("A7:F7").Font.Color = RGB(255, 200, 0)

        .Range("B9").Value = "WPX HISTORY"
        .Range("B9").Font.Bold = True
        .Range("B9").Font.Color = RGB(255, 200, 0)

        .Range("B10").Value = "Baseline"
        .Range("C10").Value = baseDate
        .Range("D10").Value = "Level " & baseLevel
        .Range("E10").Value = "Arena " & FormatPower(baseArena)
        .Range("F10").Value = "HQ " & FormatPower(baseHQ)

        .Range("B11").Value = "WPX Current"
        .Range("C11").Value = curDate
        .Range("D11").Value = "Level " & curLevel
        .Range("E11").Value = "Arena " & FormatPower(curArena)
        .Range("F11").Value = "HQ " & FormatPower(curHQ)

        .Range("B12").Value = "Level "
        .Range("C12").Value = FormatDeltaNum(SafeDiff(curLevel, baseLevel))
        .Range("D12").Value = "Arena "
        .Range("E12").Value = FormatDelta(CalcDelta(curArena, baseArena))
        .Range("F12").Value = "HQ "
        .Range("G12").Value = FormatDelta(CalcDelta(curHQ, baseHQ))
    End With

    If wpxWasOpen Then
        wbWpx.Close SaveChanges:=False
        Application.ScreenUpdating = True
    End If
End Sub

' Copies WPX Roster, Arena Power, and Player Tracking sheets into this workbook
' as hidden sheets so ShowWPXHistoryOnDashboard() can use them without prompting every time.
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
    ThisWorkbook.Sheets("WPX_PlayerTracking").Delete
    ThisWorkbook.Sheets("WPX_Archived").Delete
    On Error GoTo 0

    wbWpx.Sheets("Arena Power").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "WPX_ArenaPower"
    ThisWorkbook.Sheets("WPX_ArenaPower").Visible = xlSheetVeryHidden

    wbWpx.Sheets("Roster").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "WPX_Roster"
    ThisWorkbook.Sheets("WPX_Roster").Visible = xlSheetVeryHidden

    On Error Resume Next
    Err.Clear
    wbWpx.Sheets("Archived Players").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    If Err.Number = 0 Then
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "WPX_Archived"
        ThisWorkbook.Sheets("WPX_Archived").Visible = xlSheetVeryHidden
    End If
    Err.Clear
    On Error GoTo 0

    wbWpx.Sheets("Player Tracking").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "WPX_PlayerTracking"
    ThisWorkbook.Sheets("WPX_PlayerTracking").Visible = xlSheetVeryHidden

    wbWpx.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "WPX data imported. Run ShowWPXHistoryOnDashboard to view a player.", vbInformation
End Sub

' Clears the WPX History display area on the Player Dashboard.
Private Sub ClearWPXArea(ws As Worksheet)
    ws.Range("A7:G16").ClearContents
    On Error Resume Next
    ws.Range("A7:G16").Interior.ColorIndex = xlNone
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

Private Function GetRosterID(ws As Worksheet, player As String) As String
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If RosterNameMatches(ws.Cells(r, 2).Value, ws.Cells(r, 4).Value, player) Then _
            GetRosterID = NormalizeDashboardID(ws.Cells(r, 1).Value): Exit Function
        If RosterNameMatches(ws.Cells(r, 6).Value, ws.Cells(r, 8).Value, player) Then _
            GetRosterID = NormalizeDashboardID(ws.Cells(r, 5).Value): Exit Function
        If RosterNameMatches(ws.Cells(r, 10).Value, ws.Cells(r, 12).Value, player) Then _
            GetRosterID = NormalizeDashboardID(ws.Cells(r, 9).Value): Exit Function
        If RosterNameMatches(ws.Cells(r, 14).Value, ws.Cells(r, 16).Value, player) Then _
            GetRosterID = NormalizeDashboardID(ws.Cells(r, 13).Value): Exit Function
    Next r
    GetRosterID = ""
End Function

Private Function ResolveWpxName(wsGT As Worksheet, wsRoster As Worksheet, _
    wsArchived As Worksheet, pID As String) As String

    Dim mapped As String, mappedID As String
    mapped = WpxMapValue(pID)
    If mapped <> "" Then
        mappedID = NormalizeDashboardID(mapped)
        If mappedID <> "" Then
            ResolveWpxName = WpxNameFromID(wsRoster, mappedID)
            If ResolveWpxName = "" Then ResolveWpxName = WpxNameFromID(wsArchived, mappedID)
        Else
            ResolveWpxName = Trim$(mapped)
        End If
        If ResolveWpxName <> "" Then Exit Function
    End If

    ResolveWpxName = WpxNameFromID(wsRoster, pID)
    If ResolveWpxName = "" Then ResolveWpxName = WpxNameFromID(wsArchived, pID)
    If ResolveWpxName = "" Then ResolveWpxName = RosterAKAFromID(wsGT, pID)
End Function

Private Function WpxMapValue(pID As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("WPX ID Map")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If NormalizeDashboardID(ws.Cells(r, 1).Value) = pID Then
            WpxMapValue = Trim$(CStr(Nz(ws.Cells(r, 2).Value, "")))
            Exit Function
        End If
    Next r
End Function

Private Function WpxNameFromID(ws As Worksheet, pID As String) As String
    If ws Is Nothing Then Exit Function

    Dim blocks As Variant, b As Long, lastRow As Long, r As Long
    blocks = Array(Array(1, 2), Array(5, 6), Array(9, 10), Array(13, 14))
    For b = LBound(blocks) To UBound(blocks)
        lastRow = ws.Cells(ws.Rows.Count, blocks(b)(0)).End(xlUp).Row
        For r = 2 To lastRow
            If NormalizeDashboardID(ws.Cells(r, blocks(b)(0)).Value) = pID Then
                WpxNameFromID = Trim$(CStr(Nz(ws.Cells(r, blocks(b)(1)).Value, "")))
                Exit Function
            End If
        Next r
    Next b
End Function

Private Function RosterAKAFromID(ws As Worksheet, pID As String) As String
    If ws Is Nothing Then Exit Function

    Dim blocks As Variant, b As Long, lastRow As Long, r As Long
    blocks = Array(Array(1, 2, 4), Array(5, 6, 8), Array(9, 10, 12), Array(13, 14, 16))
    For b = LBound(blocks) To UBound(blocks)
        lastRow = ws.Cells(ws.Rows.Count, blocks(b)(0)).End(xlUp).Row
        For r = 2 To lastRow
            If NormalizeDashboardID(ws.Cells(r, blocks(b)(0)).Value) = pID Then
                RosterAKAFromID = Trim$(CStr(Nz(ws.Cells(r, blocks(b)(2)).Value, "")))
                Exit Function
            End If
        Next r
    Next b
End Function

Private Function RosterNameMatches(currentName As Variant, aka As Variant, player As String) As Boolean
    Dim needle As String
    needle = LCase$(Trim$(player))
    RosterNameMatches = (LCase$(Trim$(CStr(Nz(currentName, "")))) = needle Or _
                         LCase$(Trim$(CStr(Nz(aka, "")))) = needle)
End Function

Private Function NormalizeDashboardID(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim txt As String, digits As String, i As Long, ch As String
    txt = Trim$(CStr(v))
    txt = Replace(txt, Chr(160), "")
    txt = Replace(txt, " ", "")
    txt = Replace(txt, ",", "")
    If txt = "" Or Not IsNumeric(txt) Then Exit Function
    If InStr(1, txt, "E", vbTextCompare) > 0 Then txt = Format$(CDbl(txt), "0")

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i
    Do While Len(digits) > 1 And Left$(digits, 1) = "0"
        digits = Mid$(digits, 2)
    Loop
    If digits <> "0" Then NormalizeDashboardID = digits
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
