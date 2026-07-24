Attribute VB_Name = "PlayerDashboard"
' Player Dashboard refresh macro for GTStatsFINAL.xlsm
' Fills the existing "Player Dashboard" sheet (cell B3 has the player dropdown)
' with Overall Rank, Total Score, Weekly Avg, Weeks Played, Missed Daily/Weekly,
' 20M Weeks, current Arena/HQ power, growth baseline, WPX cross-team history,
' and a weekly breakdown.
'
' To make it auto-refresh when the player in B3 changes, add this event to the
' "Player Dashboard" sheet module:
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
'           Application.EnableEvents = False
'           RefreshPlayerDashboard
'           Application.EnableEvents = True
'       End If
'   End Sub
'
' Use ImportWPXData() once to copy WPX Stats into this workbook as hidden sheets
' so cross-team players show their legacy WPX Arena/HQ history.

Option Explicit

' Daily goal thresholds (matches gt_dashboard.html fallback)
Private Const DAILY_GOAL_MON As Long = 6000000
Private Const DAILY_GOAL_TUE As Long = 3000000
Private Const DAILY_GOAL_WED As Long = 4000000
Private Const DAILY_GOAL_THU As Long = 6000000
Private Const DAILY_GOAL_FRI As Long = 3000000

Private Const WEEKLY_GOAL As Long = 20000000

'============================================================================
' PUBLIC MACROS
'============================================================================

' Prompts for WPXStatsFinal.xlsm and copies the Arena Power + Roster sheets
' into this workbook as hidden sheets named WPX_ArenaPower and WPX_Roster.
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

    MsgBox "WPX data imported. Run RefreshPlayerDashboard to view a player.", vbInformation
End Sub

' Refreshes the Player Dashboard sheet for the player selected in B3.
Public Sub RefreshPlayerDashboard()
    Dim wsDash As Worksheet
    Dim wsPT As Worksheet
    Dim wsAP As Worksheet
    Dim wsRoster As Worksheet
    Dim player As String

    Set wsDash = ThisWorkbook.Sheets("Player Dashboard")
    player = Trim(CStr(Nz(wsDash.Range("B3").Value, "")))
    If player = "" Then
        MsgBox "Select a player from the dropdown in B3 first.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsPT = ThisWorkbook.Sheets("Player Tracking")
    Set wsAP = ThisWorkbook.Sheets("Arena Power")
    Set wsRoster = ThisWorkbook.Sheets("Roster")
    On Error GoTo 0

    If wsPT Is Nothing Then
        MsgBox "Player Tracking sheet not found.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Reset dashboard detail area
    wsDash.Range("B5:I99").ClearContents
    wsDash.Range("B5:I99").Interior.ColorIndex = xlNone

    ' Top stat labels
    wsDash.Range("B5").Value = "Overall Rank"
    wsDash.Range("C5").Value = "Total Score"
    wsDash.Range("D5").Value = "Weekly Avg"
    wsDash.Range("E5").Value = "Missed Daily"
    wsDash.Range("F5").Value = "Missed Weekly"
    wsDash.Range("G5").Value = "Weeks Played"
    wsDash.Range("H5").Value = "20M Weeks"

    Dim ptRow As Long
    ptRow = FindNameRow(wsPT, player, 1)
    If ptRow = 0 Then
        MsgBox "Player '" & player & "' not found in Player Tracking.", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Overall stats from Player Tracking
    Dim joinDate As String, overallRank As String
    Dim totalScore As Variant, weeklyAvg As Variant
    Dim missedDaily As Variant, missedWeekly As Variant
    Dim weeksPlayed As Long, weeks20M As Long

    joinDate = FormatDateNz(wsPT.Cells(ptRow, 2).Value)
    overallRank = CStr(Nz(wsPT.Cells(ptRow, 5).Value, "-"))
    totalScore = Nz(wsPT.Cells(ptRow, 3).Value, 0)
    weeklyAvg = Nz(wsPT.Cells(ptRow, 7).Value, 0)
    missedDaily = Nz(wsPT.Cells(ptRow, 8).Value, 0)
    missedWeekly = Nz(wsPT.Cells(ptRow, 9).Value, 0)

    CountWeeks wsPT, ptRow, weeksPlayed, weeks20M

    wsDash.Range("B6").Value = overallRank
    wsDash.Range("C6").Value = totalScore
    wsDash.Range("C6").NumberFormat = "#,##0"
    wsDash.Range("D6").Value = weeklyAvg
    wsDash.Range("D6").NumberFormat = "#,##0.0"
    wsDash.Range("E6").Value = missedDaily
    wsDash.Range("F6").Value = missedWeekly
    wsDash.Range("G6").Value = weeksPlayed
    wsDash.Range("H6").Value = weeks20M
    wsDash.Range("H3").Value = "Proud GT Member since: " & joinDate

    ' Arena Power + Growth
    If Not wsAP Is Nothing Then
        Dim apRow As Long
        apRow = FindNameRow(wsAP, player, 1)
        If apRow > 0 Then
            Dim curDate As String, curLevel As Variant
            Dim curArenaStr As String, curHQStr As String
            Dim curArena As Variant, curHQ As Variant
            Dim sessArena As String, sessHQ As String
            Dim lvlNote As String

            curDate = FormatDateNz(wsAP.Cells(apRow, 2).Value)
            curLevel = Nz(wsAP.Cells(apRow, 3).Value, "-")
            curArena = ParsePower(wsAP.Cells(apRow, 4).Value)
            curHQ = ParsePower(wsAP.Cells(apRow, 5).Value)
            curArenaStr = FormatPower(curArena)
            curHQStr = FormatPower(curHQ)
            sessArena = FormatDeltaFromValue(wsAP.Cells(apRow, 6).Value)
            sessHQ = FormatDeltaFromValue(wsAP.Cells(apRow, 7).Value)
            lvlNote = CStr(Nz(wsAP.Cells(apRow, 10).Value, ""))

            ' Cross-team WPX lookup
            Dim isCrossTeam As Boolean
            Dim gtBaseDate As String, gtBaseLevel As Variant
            Dim gtBaseArena As Variant, gtBaseHQ As Variant
            Dim wpxBaseDate As String, wpxBaseLevel As Variant
            Dim wpxBaseArena As Variant, wpxBaseHQ As Variant
            Dim wpxCurDate As String, wpxCurLevel As Variant
            Dim wpxCurArena As Variant, wpxCurHQ As Variant
            isCrossTeam = False

            If Not wsRoster Is Nothing Then
                Dim pID As Long
                pID = GetRosterID(wsRoster, player)
                If pID > 0 Then
                    Dim wsWpxAP As Worksheet, wsWpxRoster As Worksheet
                    On Error Resume Next
                    Set wsWpxAP = ThisWorkbook.Sheets("WPX_ArenaPower")
                    Set wsWpxRoster = ThisWorkbook.Sheets("WPX_Roster")
                    On Error GoTo 0

                    If Not wsWpxAP Is Nothing And Not wsWpxRoster Is Nothing Then
                        Dim wpxName As String, wpxRow As Long
                        wpxName = GetWpxNameFromID(wsWpxRoster, pID)
                        If wpxName <> "" Then
                            wpxRow = FindNameRow(wsWpxAP, wpxName, 1)
                            If wpxRow > 0 Then
                                isCrossTeam = True
                                Call GetArenaBaseline(wsWpxAP, wpxRow, wpxBaseDate, wpxBaseLevel, wpxBaseArena, wpxBaseHQ)
                                wpxCurDate = FormatDateNz(wsWpxAP.Cells(wpxRow, 2).Value)
                                wpxCurLevel = Nz(wsWpxAP.Cells(wpxRow, 3).Value, "-")
                                wpxCurArena = ParsePower(wsWpxAP.Cells(wpxRow, 4).Value)
                                wpxCurHQ = ParsePower(wsWpxAP.Cells(wpxRow, 5).Value)
                            End If
                        End If
                    End If
                End If
            End If

            ' Baseline from GT Arena Power history (rightmost group) only if not cross-team
            If Not isCrossTeam Then
                Call GetArenaBaseline(wsAP, apRow, gtBaseDate, gtBaseLevel, gtBaseArena, gtBaseHQ)
            End If

            ' Section header
            wsDash.Cells(8, 2).Value = "ARENA POWER"
            wsDash.Cells(8, 2).Font.Bold = True

            wsDash.Cells(9, 2).Value = "Current Level"
            wsDash.Cells(9, 3).Value = curLevel
            wsDash.Cells(9, 4).Value = "Current Arena"
            wsDash.Cells(9, 5).Value = curArenaStr
            wsDash.Cells(9, 6).Value = "Current HQ"
            wsDash.Cells(9, 7).Value = curHQStr

            wsDash.Cells(10, 2).Value = "Session Arena"
            wsDash.Cells(10, 3).Value = sessArena
            wsDash.Cells(10, 4).Value = "Session HQ"
            wsDash.Cells(10, 5).Value = sessHQ
            wsDash.Cells(10, 6).Value = "Overall Arena"
            wsDash.Cells(10, 7).Value = FormatDelta(CalcDelta(curArena, IIf(isCrossTeam, wpxBaseArena, gtBaseArena)))
            wsDash.Cells(10, 8).Value = "Overall HQ"
            wsDash.Cells(10, 9).Value = FormatDelta(CalcDelta(curHQ, IIf(isCrossTeam, wpxBaseHQ, gtBaseHQ)))

            If lvlNote <> "" And lvlNote <> "-" Then
                wsDash.Cells(11, 2).Value = "Level Note"
                wsDash.Cells(11, 3).Value = lvlNote
            End If

            ' Growth Since Joining
            Dim baseDate As String, baseLevel As Variant
            Dim baseArena As Variant, baseHQ As Variant
            If isCrossTeam Then
                baseDate = wpxBaseDate
                baseLevel = wpxBaseLevel
                baseArena = wpxBaseArena
                baseHQ = wpxBaseHQ
            Else
                baseDate = gtBaseDate
                baseLevel = gtBaseLevel
                baseArena = gtBaseArena
                baseHQ = gtBaseHQ
            End If

            wsDash.Cells(13, 2).Value = "GROWTH SINCE JOINING"
            wsDash.Cells(13, 2).Font.Bold = True

            wsDash.Cells(14, 2).Value = "Baseline"
            wsDash.Cells(14, 3).Value = baseDate
            wsDash.Cells(14, 4).Value = "Level " & baseLevel
            wsDash.Cells(14, 5).Value = "Arena " & FormatPower(baseArena)
            wsDash.Cells(14, 6).Value = "HQ " & FormatPower(baseHQ)

            wsDash.Cells(15, 2).Value = "Now"
            wsDash.Cells(15, 3).Value = curDate
            wsDash.Cells(15, 4).Value = "Level " & curLevel
            wsDash.Cells(15, 5).Value = "Arena " & curArenaStr
            wsDash.Cells(15, 6).Value = "HQ " & curHQStr

            wsDash.Cells(16, 2).Value = "Level "
            wsDash.Cells(16, 3).Value = FormatDeltaNum(SafeDiff(curLevel, baseLevel))
            wsDash.Cells(16, 4).Value = "Arena "
            wsDash.Cells(16, 5).Value = FormatDelta(CalcDelta(curArena, baseArena))
            wsDash.Cells(16, 6).Value = "HQ "
            wsDash.Cells(16, 7).Value = FormatDelta(CalcDelta(curHQ, baseHQ))

            ' WPX History (legacy, amber accent) for cross-team players
            If isCrossTeam Then
                wsDash.Cells(18, 2).Value = "WPX HISTORY"
                wsDash.Cells(18, 2).Font.Bold = True
                wsDash.Cells(18, 2).Font.Color = RGB(255, 200, 0)

                wsDash.Cells(19, 2).Value = "Baseline"
                wsDash.Cells(19, 3).Value = wpxBaseDate
                wsDash.Cells(19, 4).Value = "Level " & wpxBaseLevel
                wsDash.Cells(19, 5).Value = "Arena " & FormatPower(wpxBaseArena)
                wsDash.Cells(19, 6).Value = "HQ " & FormatPower(wpxBaseHQ)

                wsDash.Cells(20, 2).Value = "WPX Current"
                wsDash.Cells(20, 3).Value = wpxCurDate
                wsDash.Cells(20, 4).Value = "Level " & wpxCurLevel
                wsDash.Cells(20, 5).Value = "Arena " & FormatPower(wpxCurArena)
                wsDash.Cells(20, 6).Value = "HQ " & FormatPower(wpxCurHQ)

                wsDash.Cells(21, 2).Value = "Level "
                wsDash.Cells(21, 3).Value = FormatDeltaNum(SafeDiff(wpxCurLevel, wpxBaseLevel))
                wsDash.Cells(21, 4).Value = "Arena "
                wsDash.Cells(21, 5).Value = FormatDelta(CalcDelta(wpxCurArena, wpxBaseArena))
                wsDash.Cells(21, 6).Value = "HQ "
                wsDash.Cells(21, 7).Value = FormatDelta(CalcDelta(wpxCurHQ, wpxBaseHQ))
            End If
        End If
    End If

    ' Weekly Breakdown table
    Dim weeklyStartRow As Long
    weeklyStartRow = IIf(isCrossTeam, 23, 18)
    Call WriteWeeklyBreakdown(wsDash, wsPT, ptRow, player, weeklyStartRow)

    ' Tidy columns
    wsDash.Columns("A:I").AutoFit

    Application.ScreenUpdating = True
    MsgBox "Player Dashboard updated for " & player & ".", vbInformation
End Sub

'============================================================================
' HELPER SUBS/FUNCTIONS
'============================================================================

Private Function FindNameRow(ws As Worksheet, name As String, searchCol As Long) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, searchCol).End(xlUp).Row
    For r = 1 To lastRow
        If Trim(LCase(CStr(Nz(ws.Cells(r, searchCol).Value, "")))) = LCase(name) Then
            FindNameRow = r
            Exit Function
        End If
    Next r
    FindNameRow = 0
End Function

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

' Reads the rightmost non-empty history group from an Arena Power row.
Private Sub GetArenaBaseline(ws As Worksheet, r As Long, ByRef baseDate As String, _
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

Private Sub CountWeeks(ws As Worksheet, r As Long, ByRef weeksPlayed As Long, ByRef weeks20M As Long)
    Dim lastCol As Long, c As Long
    Dim header As String, score As Variant
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    weeksPlayed = 0
    weeks20M = 0
    For c = 1 To lastCol
        header = CStr(Nz(ws.Cells(1, c).Value, ""))
        If InStr(header, "Weekly Total") > 0 Then
            score = ws.Cells(r, c).Value
            If IsNumeric(score) Then
                If CDbl(score) > 0 Then weeksPlayed = weeksPlayed + 1
                If CDbl(score) >= WEEKLY_GOAL Then weeks20M = weeks20M + 1
            End If
        End If
    Next c
End Sub

Private Sub WriteWeeklyBreakdown(wsDash As Worksheet, wsPT As Worksheet, ptRow As Long, player As String, startRow As Long)
    Dim lastCol As Long, c As Long, i As Long
    Dim header As String, label As String, score As Variant, rank As Variant
    Dim weeks As Collection
    Set weeks = New Collection

    lastCol = wsPT.Cells(1, wsPT.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        header = CStr(Nz(wsPT.Cells(1, c).Value, ""))
        If InStr(header, "Weekly Total") > 0 Then
            label = ExtractWeekLabel(header)
            score = wsPT.Cells(ptRow, c).Value
            rank = wsPT.Cells(ptRow, c + 1).Value
            If IsNumeric(score) And label <> "" Then
                weeks.Add Array(label, score, rank)
            End If
        End If
    Next c

    Dim outRow As Long
    outRow = startRow
    wsDash.Cells(outRow, 2).Value = "WEEKLY BREAKDOWN"
    wsDash.Cells(outRow, 2).Font.Bold = True
    outRow = outRow + 1
    wsDash.Cells(outRow, 2).Value = "Week"
    wsDash.Cells(outRow, 3).Value = "Score"
    wsDash.Cells(outRow, 4).Value = "Rank"
    wsDash.Cells(outRow, 5).Value = "Daily Goals"
    wsDash.Cells(outRow, 6).Value = "20M Goal"

    ' Write oldest first (reverse collection)
    For i = weeks.Count To 1 Step -1
        Dim item As Variant
        item = weeks(i)
        label = item(0)
        score = item(1)
        rank = item(2)
        outRow = outRow + 1
        wsDash.Cells(outRow, 2).Value = label
        wsDash.Cells(outRow, 3).Value = score
        wsDash.Cells(outRow, 3).NumberFormat = "#,##0"
        wsDash.Cells(outRow, 4).Value = IIf(IsNumeric(rank), rank, "-")

        Dim dailyGoals As String, goalMet As String
        Call ComputeWeeklyGoals(label, player, score, dailyGoals, goalMet)
        wsDash.Cells(outRow, 5).Value = dailyGoals
        wsDash.Cells(outRow, 6).Value = goalMet
    Next i
End Sub

Private Sub ComputeWeeklyGoals(weekLabel As String, player As String, score As Variant, ByRef dailyGoals As String, ByRef goalMet As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(weekLabel)
    On Error GoTo 0

    If ws Is Nothing Then
        dailyGoals = "-"
        goalMet = IIf(IsNumeric(score) And CDbl(score) >= WEEKLY_GOAL, "YES", "NO")
        Exit Sub
    End If

    If Not IsWeekComplete(weekLabel) Then
        dailyGoals = "In Progress"
        goalMet = "In Progress"
        Exit Sub
    End If

    Dim r As Long
    r = FindNameRow(ws, player, 1)
    If r = 0 Then
        dailyGoals = "-"
    Else
        Dim hit As Long
        Dim monCol As Long, tueCol As Long, wedCol As Long, thuCol As Long, friCol As Long
        monCol = DayColumn(ws, "Monday")
        tueCol = DayColumn(ws, "Tuesday")
        wedCol = DayColumn(ws, "Wednesday")
        thuCol = DayColumn(ws, "Thursday")
        friCol = DayColumn(ws, "Friday")
        hit = 0
        If monCol > 0 And IsNumeric(ws.Cells(r, monCol).Value) And CDbl(ws.Cells(r, monCol).Value) >= DAILY_GOAL_MON Then hit = hit + 1
        If tueCol > 0 And IsNumeric(ws.Cells(r, tueCol).Value) And CDbl(ws.Cells(r, tueCol).Value) >= DAILY_GOAL_TUE Then hit = hit + 1
        If wedCol > 0 And IsNumeric(ws.Cells(r, wedCol).Value) And CDbl(ws.Cells(r, wedCol).Value) >= DAILY_GOAL_WED Then hit = hit + 1
        If thuCol > 0 And IsNumeric(ws.Cells(r, thuCol).Value) And CDbl(ws.Cells(r, thuCol).Value) >= DAILY_GOAL_THU Then hit = hit + 1
        If friCol > 0 And IsNumeric(ws.Cells(r, friCol).Value) And CDbl(ws.Cells(r, friCol).Value) >= DAILY_GOAL_FRI Then hit = hit + 1
        dailyGoals = hit & "/5"
    End If

    If IsNumeric(score) And CDbl(score) >= WEEKLY_GOAL Then
        goalMet = "YES"
    Else
        goalMet = "NO"
    End If
End Sub

Private Function IsWeekComplete(weekLabel As String) As Boolean
    Dim parts() As String, endPart As String
    Dim endDate As Date, completionDate As Date, nowDate As Date
    parts = Split(weekLabel, " - ")
    If UBound(parts) < 1 Then
        IsWeekComplete = True
        Exit Function
    End If

    nowDate = Now
    endPart = Trim(parts(1))

    On Error Resume Next
    endDate = CDate(endPart & ", " & Year(nowDate))
    On Error GoTo 0
    If endDate = 0 Then
        IsWeekComplete = True
        Exit Function
    End If

    ' If the parsed date is far in the future, assume previous year
    If endDate - nowDate > 180 Then
        endDate = DateAdd("yyyy", -1, endDate)
    End If

    completionDate = DateAdd("d", 1, endDate)
    IsWeekComplete = nowDate >= completionDate
End Function

Private Function DayColumn(ws As Worksheet, dayName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(LCase(CStr(Nz(ws.Cells(1, c).Value, "")))) = LCase(dayName) Then
            DayColumn = c
            Exit Function
        End If
    Next c
    DayColumn = 0
End Function

Private Function ExtractWeekLabel(header As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(header, "(")
    closePos = InStr(header, ")")
    If openPos > 0 And closePos > openPos Then
        ExtractWeekLabel = Trim(Mid(header, openPos + 1, closePos - openPos - 1))
    Else
        ExtractWeekLabel = ""
    End If
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

Private Function FormatDeltaFromValue(v As Variant) As String
    If IsEmpty(v) Or v = "" Or v = "-" Then
        FormatDeltaFromValue = "-"
        Exit Function
    End If
    Dim p As Variant
    p = ParsePower(v)
    If IsNumeric(p) Then
        FormatDeltaFromValue = FormatDelta(p)
    Else
        FormatDeltaFromValue = CStr(v)
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
