Attribute VB_Name = "WPXWeeksMerge"
' WPXWeeksMerge.bas
' Adds the legacy WPX weeks to the GT "Player Tracking" sheet, matched on the
' Player ID in column CL - the same ID the Player Tracking module already uses.
'
' It reads the weekly sheets out of WPXStatsFinal.xlsm exactly the way the GT
' module reads GT weekly sheets (player ID in column V, weekly total in column
' L, daily scores in C:G) and writes one "WPX Weekly Total/Rank" column pair
' per WPX week, plus WPX overall totals, onto each player's tracking row.
'
' The WPX block starts at column WPX_START_COL (default 100 = CV), which is
' past the Player ID (CL) and Possible Aliases (CM) columns, so
' UpdatePlayerTrackingFromRosterAndWeeks never clears it: that macro only
' clears C..(9 + weeks*2) and the alias column.
'
' To use:
'   1. Open GTStatsFINAL.xlsm.
'   2. Alt+F11 > File > Import File > WPXWeeksMerge.bas
'   3. Run AddWPXWeeksButton() once to put a "Pull WPX Weeks" button on the
'      Player Tracking sheet, then click it whenever you want a refresh.
'      Or run MergeWPXWeeksIntoPlayerTracking() from the macro list.
'   4. Pick WPXStatsFinal.xlsm when prompted (skipped when it sits in the same
'      folder as GTStatsFINAL.xlsm).
'
' To refresh GT and WPX in one click, add this line at the end of
' UpdatePlayerTrackingFromRosterAndWeeks, just before SafeExit:
'
'     Call MergeWPXWeeksIntoPlayerTracking

Option Explicit

Private Const PT_SHEET       As String = "Player Tracking"
Private Const PT_ID_COL      As String = "CL"
Private Const WPX_START_COL  As Long = 100          ' column CV
Private Const WPX_FILE       As String = "WPXStatsFinal.xlsm"

' Puts a "Pull WPX Weeks" button on the Player Tracking sheet.
Public Sub AddWPXWeeksButton()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ws.Shapes("btnPullWPXWeeks").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 150, 28)
    shp.name = "btnPullWPXWeeks"
    shp.Fill.ForeColor.RGB = RGB(0, 90, 160)
    shp.Line.ForeColor.RGB = RGB(0, 60, 110)
    With shp.TextFrame2.TextRange
        .Text = "Pull WPX Weeks"
        .Font.Size = 11
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.TextFrame2.HorizontalAnchor = msoAnchorCenter
    shp.OnAction = "MergeWPXWeeksIntoPlayerTracking"

    ws.Activate
    MsgBox "Button added to the " & PT_SHEET & " sheet.", vbInformation
End Sub

' Fills the WPX block on the Player Tracking sheet for every player whose ID
' in column CL also appears in the WPX workbook.
Public Sub MergeWPXWeeksIntoPlayerTracking()

    Dim wsPT As Worksheet
    Dim wbWpx As Workbook
    Dim oldCalc As XlCalculation
    Dim opened As Boolean

    On Error GoTo SafeExit

    On Error Resume Next
    Set wsPT = ThisWorkbook.Worksheets(PT_SHEET)
    On Error GoTo SafeExit
    If wsPT Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    Set wbWpx = OpenWpxWorkbook(opened)
    If wbWpx Is Nothing Then Exit Sub

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim weekNames() As String, weekDates() As Date, weekCount As Long
    GetWpxWeeklySheets wbWpx, weekNames, weekDates, weekCount

    If weekCount = 0 Then
        MsgBox "No weekly sheets found in " & wbWpx.name & ".", vbExclamation
        GoTo SafeExit
    End If

    ' Weekly sheets in the WPX workbook may predate its player-ID column, so
    ' rows without an ID are resolved through the WPX roster by name.
    Dim nameToID As Object
    Set nameToID = BuildWpxNameToID(wbWpx)

    Dim weekMaps() As Object, m As Long
    ReDim weekMaps(1 To weekCount)
    For m = 1 To weekCount
        Set weekMaps(m) = BuildWpxWeekMap(wbWpx.Worksheets(weekNames(m)), nameToID)
    Next m

    Dim firstRow As Long, lastRow As Long
    firstRow = 2
    lastRow = wsPT.Cells(wsPT.Rows.count, "A").End(xlUp).row
    If lastRow < firstRow Then GoTo SafeExit

    ' Clear and rewrite the whole WPX block so removed weeks do not linger.
    Dim lastWpxCol As Long
    lastWpxCol = WPX_START_COL + 2 + (weekCount * 2)
    wsPT.Range(wsPT.Cells(1, WPX_START_COL), wsPT.Cells(lastRow, lastWpxCol + 20)).ClearContents

    wsPT.Cells(1, WPX_START_COL).Value = "WPX Overall Total"
    wsPT.Cells(1, WPX_START_COL + 1).Value = "WPX Weeks Played"
    wsPT.Cells(1, WPX_START_COL + 2).Value = "WPX Weekly Average"

    Dim i As Long, displayIndex As Long, totalCol As Long, rankCol As Long
    For i = 1 To weekCount
        displayIndex = weekCount - i + 1
        totalCol = WPX_START_COL + 3 + (displayIndex - 1) * 2
        rankCol = totalCol + 1
        wsPT.Cells(1, totalCol).Value = "WPX Weekly Total (" & weekNames(i) & ")"
        wsPT.Cells(1, rankCol).Value = "WPX Weekly Rank (" & weekNames(i) & ")"
    Next i

    Dim r As Long, playerID As String
    Dim matched As Long, unmatched As Long
    Dim totalSum As Double, weeksPlayed As Long
    Dim v As Variant

    For r = firstRow To lastRow
        playerID = WpxNormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        totalSum = 0
        weeksPlayed = 0

        If playerID <> "" Then
            Dim seenInWpx As Boolean
            seenInWpx = False

            For i = 1 To weekCount
                displayIndex = weekCount - i + 1
                totalCol = WPX_START_COL + 3 + (displayIndex - 1) * 2

                If weekMaps(i).Exists(playerID) Then
                    seenInWpx = True
                    v = weekMaps(i)(playerID)
                    If IsNumeric(v) And CStr(v) <> "" Then
                        wsPT.Cells(r, totalCol).Value = CDbl(v)
                        totalSum = totalSum + CDbl(v)
                        weeksPlayed = weeksPlayed + 1
                    End If
                End If
            Next i

            If seenInWpx Then
                matched = matched + 1
                wsPT.Cells(r, WPX_START_COL).Value = totalSum
                wsPT.Cells(r, WPX_START_COL + 1).Value = weeksPlayed
                If weeksPlayed > 0 Then
                    wsPT.Cells(r, WPX_START_COL + 2).Value = totalSum / weeksPlayed
                End If
            Else
                unmatched = unmatched + 1
            End If
        End If
    Next r

    ' Ranks per WPX week, over the players who actually have a score that week.
    For i = 1 To weekCount
        displayIndex = weekCount - i + 1
        totalCol = WPX_START_COL + 3 + (displayIndex - 1) * 2
        RankColumn wsPT, totalCol, totalCol + 1, firstRow, lastRow
    Next i

SafeExit:
    If Not wbWpx Is Nothing Then
        If opened Then wbWpx.Close SaveChanges:=False
    End If

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "MergeWPXWeeksIntoPlayerTracking error " & Err.Number & ": " & Err.Description, vbExclamation
    ElseIf weekCount > 0 Then
        MsgBox "WPX weeks merged into " & PT_SHEET & "." & vbCrLf & _
               "WPX weeks read: " & weekCount & vbCrLf & _
               "Players with WPX history: " & matched & vbCrLf & _
               "Players with an ID but no WPX history: " & unmatched, vbInformation, "Done"
    End If

End Sub

' ID -> weekly total for one WPX weekly sheet, same layout as the GT weeklies:
' column A name, C:G daily scores, L weekly total, V player ID.
Private Function BuildWpxWeekMap(ws As Worksheet, nameToID As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.count, "A").End(xlUp).row, _
        ws.Cells(ws.Rows.count, "V").End(xlUp).row)
    If lastRow < 2 Then
        Set BuildWpxWeekMap = dict
        Exit Function
    End If

    Dim data As Variant
    data = ws.Range("A2:V" & lastRow).Value

    Dim r As Long, playerID As String, nameKey As String, weeklyTotal As Variant
    For r = 1 To UBound(data, 1)
        playerID = WpxNormalizeID(data(r, 22))
        If playerID = "" Then
            nameKey = WpxNameKey(data(r, 1))
            If nameKey <> "" Then
                If nameToID.Exists(nameKey) Then playerID = CStr(nameToID(nameKey))
            End If
        End If

        If playerID <> "" Then
            If IsNumeric(data(r, 12)) And CStr(data(r, 12)) <> "" Then
                weeklyTotal = CDbl(data(r, 12))
            Else
                weeklyTotal = ""
            End If

            If Not dict.Exists(playerID) Then
                dict.Add playerID, weeklyTotal
            ElseIf IsNumeric(weeklyTotal) And Not IsNumeric(dict(playerID)) Then
                dict(playerID) = weeklyTotal
            End If
        End If
    Next r

    Set BuildWpxWeekMap = dict
End Function

' Player name -> ID from the WPX Roster, using the same four ID/name blocks as
' the GT roster (A/B, E/F, I/J, M/N).
Private Function BuildWpxNameToID(wb As Workbook) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Roster")
    On Error GoTo 0
    If ws Is Nothing Then
        Set BuildWpxNameToID = dict
        Exit Function
    End If

    Dim blocks As Variant, b As Long, r As Long
    Dim id As String, nameKey As String
    blocks = Array(Array(1, 2), Array(5, 6), Array(9, 10), Array(13, 14))

    For b = LBound(blocks) To UBound(blocks)
        For r = 2 To ws.Cells(ws.Rows.count, blocks(b)(1)).End(xlUp).row
            id = WpxNormalizeID(ws.Cells(r, blocks(b)(0)).Value)
            nameKey = WpxNameKey(ws.Cells(r, blocks(b)(1)).Value)
            If id <> "" And nameKey <> "" Then
                If Not dict.Exists(nameKey) Then dict.Add nameKey, id
            End If
        Next r
    Next b

    Set BuildWpxNameToID = dict
End Function

' Case- and punctuation-insensitive key for player names.
Private Function WpxNameKey(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim s As String, out As String, i As Long, ch As String
    s = UCase$(Trim$(Replace(CStr(v), Chr(160), " ")))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    WpxNameKey = out
End Function

' Same ID normalisation the Player Tracking module uses, kept local so this
' module also works on its own.
Private Function WpxNormalizeID(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim txt As String
    txt = Trim$(CStr(v))
    txt = Replace(txt, Chr(160), "")
    txt = Replace(txt, " ", "")
    WpxNormalizeID = UCase$(txt)
End Function

' Weekly sheets in the WPX workbook, oldest first. Same naming rule as the GT
' module ("Jul 13 - Jul 19"), minus the GT season cutoff so the whole WPX
' history is picked up.
Private Sub GetWpxWeeklySheets(wb As Workbook, ByRef weekNames() As String, _
    ByRef weekDates() As Date, ByRef weekCount As Long)

    Dim ws As Worksheet, d As Date
    weekCount = 0

    For Each ws In wb.Worksheets
        If IsWpxWeeklySheet(ws.name) Then
            d = ParseWpxWeekStart(ws.name)
            If d > 0 And d <= Date + 3 Then
                weekCount = weekCount + 1
                ReDim Preserve weekNames(1 To weekCount)
                ReDim Preserve weekDates(1 To weekCount)
                weekNames(weekCount) = ws.name
                weekDates(weekCount) = d
            End If
        End If
    Next ws

    Dim i As Long, j As Long, tmpName As String, tmpDate As Date
    For i = 1 To weekCount - 1
        For j = i + 1 To weekCount
            If weekDates(j) < weekDates(i) Then
                tmpDate = weekDates(i): weekDates(i) = weekDates(j): weekDates(j) = tmpDate
                tmpName = weekNames(i): weekNames(i) = weekNames(j): weekNames(j) = tmpName
            End If
        Next j
    Next i
End Sub

Private Function IsWpxWeeklySheet(ByVal sheetName As String) As Boolean
    Select Case LCase$(Trim$(sheetName))
        Case "player tracking", "player tracking new", "master template", _
             "player dashboard", "archived players", "control", _
             "roster", "start", "end", "score import", "week settings", _
             "id match log", "arena power"
            Exit Function
    End Select
    If InStr(1, sheetName, "OLD", vbTextCompare) > 0 Then Exit Function
    IsWpxWeeklySheet = (InStr(1, sheetName, " - ") > 0)
End Function

' Week sheet names carry no year, so a parsed date in the future is rolled
' back a year - WPX weeks are all older than the current GT season.
Private Function ParseWpxWeekStart(ByVal sheetName As String) As Date
    Dim parts() As String, startText As String, d As Date

    If InStr(1, sheetName, " - ") = 0 Then Exit Function
    parts = Split(sheetName, " - ")
    If UBound(parts) < 1 Then Exit Function
    startText = Trim$(parts(0))

    On Error Resume Next
    d = DateValue(startText & ", " & Year(Date))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If d > Date + 3 Then d = DateSerial(Year(d) - 1, Month(d), Day(d))
    ParseWpxWeekStart = d
End Function

' Dense ranking over the numeric values in valueCol, highest score = rank 1.
Private Sub RankColumn(ws As Worksheet, valueCol As Long, rankCol As Long, _
    firstRow As Long, lastRow As Long)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstRow, valueCol), ws.Cells(lastRow, valueCol))

    Dim count As Long
    count = Application.WorksheetFunction.count(rng)

    Dim r As Long, v As Variant
    For r = firstRow To lastRow
        v = ws.Cells(r, valueCol).Value
        If count > 0 And IsNumeric(v) And CStr(v) <> "" Then
            ws.Cells(r, rankCol).Value = Application.WorksheetFunction.rank(v, rng, 0)
        Else
            ws.Cells(r, rankCol).ClearContents
        End If
    Next r
End Sub

' Opens the WPX workbook, or returns it if already open. Prompts only when the
' file is not sitting next to GTStatsFINAL.xlsm.
Private Function OpenWpxWorkbook(ByRef opened As Boolean) As Workbook
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(WPX_FILE)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set OpenWpxWorkbook = wb
        Exit Function
    End If

    Dim f As Variant, samePath As String
    samePath = ThisWorkbook.Path & Application.PathSeparator & WPX_FILE
    If Dir(samePath) <> "" Then
        f = samePath
    Else
        f = Application.GetOpenFilename( _
            "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , "Select " & WPX_FILE)
        If VarType(f) = vbBoolean Then Exit Function
        If CStr(f) = "False" Then Exit Function
    End If

    Set OpenWpxWorkbook = Workbooks.Open(CStr(f), ReadOnly:=True)
    opened = True
End Function
