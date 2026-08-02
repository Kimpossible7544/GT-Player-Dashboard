Attribute VB_Name = "WPXHistorySheet"
' WPXHistorySheet.bas
' Rebuilds the "WPX History" sheet in GTStatsFINAL.xlsm from WPXStatsFinal.xlsm,
' so the Player Dashboard formulas that read it, e.g.
'
'     =IFERROR(INDEX('WPX History'!$E$2:$E$100,
'              MATCH(TRIM($B$3),'WPX History'!$B$2:$B$100,0)),"")
'
' resolve for every cross-alliance player instead of only the ones typed in by
' hand. Players are matched on the ID in Player Tracking column CL, but the GT
' name goes in the name column, because that is what the dashboard matches on.
'
' The sheet's own header row drives where things land: each header is matched
' against the stats below, so an existing layout is kept as-is and only the
' recognised columns are written. Anything unrecognised is left untouched and
' listed in the summary. Where the sheet does not exist yet it is created with
' the default layout:
'
'     A Player ID | B Player | C First Week | D Last Week | E Overall Total
'     F Weeks Played | G Weekly Average | H Best Week | I Best Week Score
'     J Overall Rank | K Missed Daily Goals | L Missed Weekly Goals
'
' Column E is the overall total in that default because the dashboard formula
' above already points at E.
'
' To use:
'   1. Open GTStatsFINAL.xlsm.
'   2. Alt+F11 > File > Import File > WPXHistorySheet.bas
'   3. Alt+F8 > RefreshWPXHistorySheet.
'   4. Pick WPXStatsFinal.xlsm when prompted (skipped when it sits in the same
'      folder as GTStatsFINAL.xlsm).

Option Explicit

Private Const PT_SHEET      As String = "Player Tracking"
Private Const PT_ID_COL     As String = "CL"
Private Const HIST_SHEET    As String = "WPX History"
Private Const WPX_PT_SHEET  As String = "Player Tracking"
Private Const WPX_FILE      As String = "WPXStatsFinal.xlsm"

' Stat keys, also the default headers.
Private Const F_ID       As String = "Player ID"
Private Const F_NAME     As String = "Player"
Private Const F_FIRST    As String = "First Week"
Private Const F_LAST     As String = "Last Week"
Private Const F_TOTAL    As String = "Overall Total"
Private Const F_WEEKS    As String = "Weeks Played"
Private Const F_AVG      As String = "Weekly Average"
Private Const F_BESTWK   As String = "Best Week"
Private Const F_BEST     As String = "Best Week Score"
Private Const F_RANK     As String = "Overall Rank"
Private Const F_MISSD    As String = "Missed Daily Goals"
Private Const F_MISSW    As String = "Missed Weekly Goals"

' Rebuilds WPX History for every GT player who has an ID in CL and a row in the
' WPX workbook.
Public Sub RefreshWPXHistorySheet()

    Dim wsPT As Worksheet, wsHist As Worksheet, wsWpx As Worksheet
    Dim wbWpx As Workbook
    Dim oldCalc As XlCalculation
    Dim opened As Boolean
    Dim written As Long, skipped As Long
    Dim unknownHeaders As String
    Dim createdSheet As Boolean

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

    On Error Resume Next
    Set wsWpx = wbWpx.Worksheets(WPX_PT_SHEET)
    On Error GoTo SafeExit
    If wsWpx Is Nothing Then
        MsgBox "Sheet """ & WPX_PT_SHEET & """ not found in " & wbWpx.name & ".", vbExclamation
        GoTo SafeExit
    End If

    Set wsHist = GetHistorySheet(createdSheet)

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim cols As Object
    Set cols = MapHeaderColumns(wsHist, unknownHeaders)
    If cols.count = 0 Then
        MsgBox "None of the headers on " & HIST_SHEET & " were recognised." & vbCrLf & _
               "Expected names like ""Player"", ""Overall Total"", ""Weekly Average"".", _
               vbExclamation
        GoTo SafeExit
    End If

    Dim weekCols As Object, weekOrder As Variant
    Set weekCols = BuildWeekColumnMap(wsWpx)
    weekOrder = SortedWeekKeys(weekCols)

    Dim wpxRowByID As Object
    Set wpxRowByID = BuildWpxRowMap(wsWpx, wbWpx)

    ClearHistoryRows wsHist, cols

    Dim firstRow As Long, lastRow As Long
    firstRow = 2
    lastRow = wsPT.Cells(wsPT.Rows.count, "A").End(xlUp).row

    Dim r As Long, outRow As Long, playerID As String, wpxRow As Long
    Dim stats As Object
    outRow = 2

    For r = firstRow To lastRow
        playerID = WpxNormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        If playerID <> "" Then
            If wpxRowByID.Exists(playerID) Then
                wpxRow = wpxRowByID(playerID)
                Set stats = BuildPlayerStats(wsWpx, wpxRow, weekCols, weekOrder)

                If stats(F_WEEKS) > 0 Then
                    stats(F_ID) = playerID
                    stats(F_NAME) = Trim$(CStr(wsPT.Cells(r, "A").Value))
                    WriteStatsRow wsHist, outRow, cols, stats
                    outRow = outRow + 1
                    written = written + 1
                Else
                    skipped = skipped + 1
                End If
            End If
        End If
    Next r

    ' Overall rank over the players just written, highest total first.
    If cols.Exists(F_RANK) And cols.Exists(F_TOTAL) And written > 0 Then
        RankHistory wsHist, CLng(cols(F_TOTAL)), CLng(cols(F_RANK)), 2, outRow - 1
    End If

SafeExit:
    If Not wbWpx Is Nothing Then
        If opened Then wbWpx.Close SaveChanges:=False
    End If

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "RefreshWPXHistorySheet error " & Err.Number & ": " & Err.Description, vbExclamation
    ElseIf Not wsHist Is Nothing Then
        Dim msg As String
        msg = HIST_SHEET & " rebuilt from " & WPX_FILE & "." & vbCrLf & _
              "Players written: " & written & vbCrLf & _
              "Matched but with no WPX weeks: " & skipped
        If createdSheet Then msg = msg & vbCrLf & "(sheet created with the default layout)"
        If unknownHeaders <> "" Then _
            msg = msg & vbCrLf & vbCrLf & "Headers left untouched:" & unknownHeaders
        MsgBox msg, vbInformation, "Done"
    End If

End Sub

' The WPX History sheet, created with the default layout when missing.
Private Function GetHistorySheet(ByRef created As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HIST_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = HIST_SHEET
        ws.Range("A1:L1").Value = Array(F_ID, F_NAME, F_FIRST, F_LAST, F_TOTAL, _
            F_WEEKS, F_AVG, F_BESTWK, F_BEST, F_RANK, F_MISSD, F_MISSW)
        ws.Rows(1).Font.Bold = True
        created = True
    End If

    Set GetHistorySheet = ws
End Function

' Stat key -> column, read off the sheet's own header row.
Private Function MapHeaderColumns(ws As Worksheet, ByRef unknownHeaders As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long, header As String, key As String
    For c = 1 To lastCol
        header = Trim$(CStr(ws.Cells(1, c).Value))
        If header <> "" Then
            key = StatKeyForHeader(header)
            If key = "" Then
                unknownHeaders = unknownHeaders & vbCrLf & "  " & _
                    Split(ws.Cells(1, c).Address(True, False), "$")(0) & ": " & header
            ElseIf Not dict.Exists(key) Then
                dict.Add key, c
            End If
        End If
    Next c

    Set MapHeaderColumns = dict
End Function

' Header text -> stat key, tolerant of the usual wordings.
Private Function StatKeyForHeader(ByVal header As String) As String
    Dim h As String
    h = LCase$(Trim$(Replace(header, Chr(160), " ")))
    h = Replace(h, "wpx ", "")
    h = Replace(h, "  ", " ")

    Select Case h
        Case "id", "player id", "roster id"
            StatKeyForHeader = F_ID
        Case "player", "name", "player name"
            StatKeyForHeader = F_NAME
        Case "first week", "start week", "joined", "join week", "from"
            StatKeyForHeader = F_FIRST
        Case "last week", "end week", "left", "to"
            StatKeyForHeader = F_LAST
        Case "overall total", "total", "total score", "overall score", "score"
            StatKeyForHeader = F_TOTAL
        Case "weeks played", "weeks", "week count"
            StatKeyForHeader = F_WEEKS
        Case "weekly average", "average", "avg", "overall weekly average", "avg weekly"
            StatKeyForHeader = F_AVG
        Case "best week"
            StatKeyForHeader = F_BESTWK
        Case "best week score", "best score", "best week total", "highest week"
            StatKeyForHeader = F_BEST
        Case "overall rank", "rank"
            StatKeyForHeader = F_RANK
        Case "missed daily goals", "missed dailies", "missed daily"
            StatKeyForHeader = F_MISSD
        Case "missed weekly goals", "missed weeklies", "missed weekly"
            StatKeyForHeader = F_MISSW
    End Select
End Function

' Everything the sheet can hold for one player, from their WPX tracking row.
Private Function BuildPlayerStats(ws As Worksheet, wpxRow As Long, _
    weekCols As Object, weekOrder As Variant) As Object

    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")

    Dim total As Double, weeks As Long, best As Double
    Dim bestWeek As String, firstWeek As String, lastWeek As String
    Dim i As Long, key As String, label As String, v As Variant

    For i = LBound(weekOrder) To UBound(weekOrder)
        key = weekOrder(i)
        label = weekCols(key)(1)
        v = ws.Cells(wpxRow, weekCols(key)(0)).Value

        If HasScore(v) Then
            total = total + CDbl(v)
            weeks = weeks + 1
            If firstWeek = "" Then firstWeek = label
            lastWeek = label
            If CDbl(v) > best Then
                best = CDbl(v)
                bestWeek = label
            End If
        End If
    Next i

    s(F_ID) = ""
    s(F_NAME) = ""
    s(F_FIRST) = firstWeek
    s(F_LAST) = lastWeek
    s(F_TOTAL) = total
    s(F_WEEKS) = weeks
    If weeks > 0 Then s(F_AVG) = total / weeks Else s(F_AVG) = ""
    s(F_BESTWK) = bestWeek
    If weeks > 0 Then s(F_BEST) = best Else s(F_BEST) = ""
    s(F_RANK) = ""
    s(F_MISSD) = NumOrBlank(ws.Cells(wpxRow, "H").Value)
    s(F_MISSW) = NumOrBlank(ws.Cells(wpxRow, "I").Value)

    Set BuildPlayerStats = s
End Function

Private Sub WriteStatsRow(ws As Worksheet, row As Long, cols As Object, stats As Object)
    Dim k As Variant
    For Each k In cols.Keys
        If stats.Exists(k) Then ws.Cells(row, cols(k)).Value = stats(k)
    Next k
End Sub

' Wipes the columns this macro owns, leaving any hand-kept extra columns alone.
Private Sub ClearHistoryRows(ws As Worksheet, cols As Object)
    Dim lastRow As Long
    lastRow = ws.UsedRange.row + ws.UsedRange.Rows.count - 1
    If lastRow < 2 Then Exit Sub

    Dim k As Variant
    For Each k In cols.Keys
        ws.Range(ws.Cells(2, cols(k)), ws.Cells(lastRow, cols(k))).ClearContents
    Next k
End Sub

' Week start day/month -> Array(column, label), from the "Weekly Total (...)"
' headers on the WPX Player Tracking sheet.
Private Function BuildWeekColumnMap(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long, header As String, label As String, key As String
    For c = 1 To lastCol
        header = Trim$(CStr(ws.Cells(1, c).Value))
        If InStr(1, header, "Weekly Total", vbTextCompare) = 1 Then
            label = WeekLabelFromHeader(header)
            If label <> "" Then
                key = WeekKey(label)
                If key <> "" Then
                    If Not dict.Exists(key) Then dict.Add key, Array(c, label)
                End If
            End If
        End If
    Next c

    Set BuildWeekColumnMap = dict
End Function

' Week keys oldest first, so First/Last Week come out the right way round.
Private Function SortedWeekKeys(weekCols As Object) As Variant
    Dim keys As Variant
    keys = weekCols.Keys

    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    SortedWeekKeys = keys
End Function

' "Weekly Total (Jul 13 - Jul 19)" -> "Jul 13 - Jul 19"
Private Function WeekLabelFromHeader(ByVal header As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(header, "(")
    closePos = InStrRev(header, ")")
    If openPos = 0 Or closePos <= openPos + 1 Then Exit Function
    WeekLabelFromHeader = Trim$(Mid$(header, openPos + 1, closePos - openPos - 1))
End Function

' Weeks are keyed on their start date. Sorting is chronological within a season,
' so the key keeps month before day.
Private Function WeekKey(ByVal label As String) As String
    Dim startText As String, d As Date

    startText = label
    If InStr(1, label, " - ") > 0 Then startText = Trim$(Split(label, " - ")(0))

    On Error Resume Next
    d = DateValue(startText & ", " & Year(Date))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        WeekKey = UCase$(Trim$(label))
        Exit Function
    End If
    On Error GoTo 0

    If d > Date + 3 Then d = DateSerial(Year(d) - 1, Month(d), Day(d))
    WeekKey = Format$(d, "mm-dd")
End Function

' Player ID -> row on the WPX Player Tracking sheet.
Private Function BuildWpxRowMap(ws As Worksheet, wb As Workbook) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim idCol As Long
    idCol = FindWpxIDColumn(ws)

    Dim nameToID As Object
    Set nameToID = BuildWpxNameToID(wb)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If idCol > 0 Then
        If ws.Cells(ws.Rows.count, idCol).End(xlUp).row > lastRow Then
            lastRow = ws.Cells(ws.Rows.count, idCol).End(xlUp).row
        End If
    End If

    Dim r As Long, playerID As String, nameKey As String
    For r = 2 To lastRow
        playerID = ""
        If idCol > 0 Then playerID = WpxNormalizeID(ws.Cells(r, idCol).Value)

        If playerID = "" Then
            nameKey = WpxNameKey(ws.Cells(r, 1).Value)
            If nameKey <> "" Then
                If nameToID.Exists(nameKey) Then playerID = CStr(nameToID(nameKey))
            End If
        End If

        If playerID <> "" Then
            If Not dict.Exists(playerID) Then dict.Add playerID, r
        End If
    Next r

    Set BuildWpxRowMap = dict
End Function

' CL when it holds IDs, otherwise the first "Player ID" header.
Private Function FindWpxIDColumn(ws As Worksheet) As Long
    If WpxNormalizeID(ws.Cells(2, PT_ID_COL).Value) <> "" Then
        FindWpxIDColumn = ws.Range(PT_ID_COL & "1").Column
        Exit Function
    End If

    Dim lastCol As Long, c As Long, header As String
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        header = LCase$(Trim$(CStr(ws.Cells(1, c).Value)))
        If header = "player id" Or header = "id" Then
            FindWpxIDColumn = c
            Exit Function
        End If
    Next c

    FindWpxIDColumn = ws.Range(PT_ID_COL & "1").Column
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

' Highest total = rank 1.
Private Sub RankHistory(ws As Worksheet, valueCol As Long, rankCol As Long, _
    firstRow As Long, lastRow As Long)

    If lastRow < firstRow Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstRow, valueCol), ws.Cells(lastRow, valueCol))

    Dim r As Long, v As Variant
    For r = firstRow To lastRow
        v = ws.Cells(r, valueCol).Value
        If HasScore(v) Then
            ws.Cells(r, rankCol).Value = Application.WorksheetFunction.rank(v, rng, 0)
        End If
    Next r
End Sub

Private Function NumOrBlank(ByVal v As Variant) As Variant
    If HasScore(v) Then NumOrBlank = CDbl(v) Else NumOrBlank = ""
End Function

Private Function HasScore(ByVal v As Variant) As Boolean
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CStr(v) = "" Then Exit Function
    HasScore = True
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

' Same ID normalisation the Player Tracking module uses.
Private Function WpxNormalizeID(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim txt As String
    txt = Trim$(CStr(v))
    txt = Replace(txt, Chr(160), "")
    txt = Replace(txt, " ", "")
    WpxNormalizeID = UCase$(txt)
End Function

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
