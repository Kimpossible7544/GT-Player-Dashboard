Attribute VB_Name = "WPXHistorySheet"
' WPXHistorySheet.bas
' Fills the "WPX History" sheet in GTStatsFINAL.xlsm from WPXStatsFinal.xlsm,
' so the Player Dashboard formulas that read it, e.g.
'
'     =IFERROR(INDEX('WPX History'!$E$2:$E$100,
'              MATCH(TRIM($B$3),'WPX History'!$B$2:$B$100,0)),"")
'
' resolve for every cross-alliance player instead of only the rows filled in by
' hand. Players are matched on the ID in the sheet's own ID column, falling back
' to Player Tracking column CL by name, and existing rows are updated in place
' so the sheet's row order, notes and any hand-kept columns survive.
'
' The header row drives placement: each header is matched against the stats
' below, so the layout can be anything. Per-week columns named
' "WPX Weekly Total (Mar 23 - Mar 29)" / "WPX Weekly Rank (...)" are filled from
' the matching week on the WPX Player Tracking sheet. Headers that match nothing
' (Arena Power, HQ, Level, ...) are left untouched and listed in the summary, so
' run ListWPXHistoryHeaders if you want to see exactly what the macro made of
' each one.
'
' Recognised headers, in any wording containing these words:
'   ID | Player/Name | Alliance | Join Date | WPX Overall Total | GT Overall
'   Total | Combined Total | Weeks Played | Weekly Avg | Best Week | Best Week
'   Score | Overall Rank | Missed Daily Goals | Missed Weekly Goals |
'   First Week | Last Week | WPX Weekly Total (week) | WPX Weekly Rank (week)
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
Private Const ALLIANCE_TAG  As String = "WPX"

' Stat keys.
Private Const F_ID       As String = "ID"
Private Const F_NAME     As String = "Player Name"
Private Const F_ALLIANCE As String = "Alliance"
Private Const F_JOIN     As String = "Join Date"
Private Const F_FIRST    As String = "First Week"
Private Const F_LAST     As String = "Last Week"
Private Const F_TOTAL    As String = "WPX Overall Total"
Private Const F_GTTOTAL  As String = "GT Overall Total"
Private Const F_COMBINED As String = "Combined Total"
Private Const F_WEEKS    As String = "Weeks Played"
Private Const F_AVG      As String = "WPX Weekly Avg"
Private Const F_BESTWK   As String = "Best Week"
Private Const F_BEST     As String = "Best Week Score"
Private Const F_RANK     As String = "Overall Rank"
Private Const F_MISSD    As String = "Missed Daily Goals"
Private Const F_MISSW    As String = "Missed Weekly Goals"

' Descriptive columns that are only written when the cell is empty, so numbers
' typed in by hand are never replaced by a computed guess.
Private Const FILL_IF_BLANK As String = "|" & F_ALLIANCE & "|" & F_JOIN & "|"

' Writes every header on the WPX History sheet next to the stat it maps to,
' onto a "WPX History Headers" sheet. Use it when a column stays blank.
Public Sub ListWPXHistoryHeaders()
    Dim wsHist As Worksheet, wsOut As Worksheet
    On Error Resume Next
    Set wsHist = ThisWorkbook.Worksheets(HIST_SHEET)
    On Error GoTo 0
    If wsHist Is Nothing Then
        MsgBox "Sheet """ & HIST_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("WPX History Headers")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        wsOut.name = "WPX History Headers"
    End If
    wsOut.Cells.ClearContents
    wsOut.Range("A1:D1").Value = Array("Column", "Header", "Filled as", "Week")

    Dim lastCol As Long, c As Long, out As Long
    lastCol = wsHist.Cells(1, wsHist.Columns.count).End(xlToLeft).Column
    out = 2

    Dim header As String, key As String, weekKey As String, isRank As Boolean
    For c = 1 To lastCol
        header = Trim$(CStr(wsHist.Cells(1, c).Value))
        If header <> "" Then
            key = StatKeyForHeader(header)
            weekKey = WeekKeyFromHeader(header, isRank)

            wsOut.Cells(out, 1).Value = ColumnLetter(c)
            wsOut.Cells(out, 2).Value = header
            If weekKey <> "" Then
                wsOut.Cells(out, 3).Value = IIf(isRank, "weekly rank", "weekly total")
                wsOut.Cells(out, 4).Value = WeekLabelFromHeader(header)
            ElseIf key <> "" Then
                wsOut.Cells(out, 3).Value = key
            Else
                wsOut.Cells(out, 3).Value = "(not recognised - left alone)"
            End If
            out = out + 1
        End If
    Next c

    wsOut.Columns("A:D").AutoFit
    wsOut.Activate
    MsgBox "Header mapping written to " & wsOut.name & ".", vbInformation
End Sub

' Fills WPX History for every row on it, and for every GT player with a WPX row
' that is missing from it.
Public Sub RefreshWPXHistorySheet()

    Dim wsPT As Worksheet, wsHist As Worksheet, wsWpx As Worksheet
    Dim wbWpx As Workbook
    Dim oldCalc As XlCalculation
    Dim opened As Boolean
    Dim updated As Long, added As Long, noData As Long
    Dim unknownHeaders As String

    On Error GoTo SafeExit

    On Error Resume Next
    Set wsPT = ThisWorkbook.Worksheets(PT_SHEET)
    Set wsHist = ThisWorkbook.Worksheets(HIST_SHEET)
    On Error GoTo SafeExit
    If wsPT Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If
    If wsHist Is Nothing Then
        MsgBox "Sheet """ & HIST_SHEET & """ not found.", vbExclamation
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

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim cols As Object, weekTotalCols As Object, weekRankCols As Object
    Set cols = CreateObject("Scripting.Dictionary")
    Set weekTotalCols = CreateObject("Scripting.Dictionary")
    Set weekRankCols = CreateObject("Scripting.Dictionary")
    MapHeaderColumns wsHist, cols, weekTotalCols, weekRankCols, unknownHeaders

    If Not cols.Exists(F_NAME) Then
        MsgBox "No player-name column found on " & HIST_SHEET & _
               ". Run ListWPXHistoryHeaders to see what the headers mapped to.", vbExclamation
        GoTo SafeExit
    End If

    ' GT side: name and overall total per ID, and the ID for a name typed into
    ' WPX History without one.
    Dim gtByID As Object, gtIDByName As Object
    BuildGTMaps wsPT, gtByID, gtIDByName

    Dim wpxWeekCols As Object, weekOrder As Variant
    Set wpxWeekCols = BuildWpxWeekColumnMap(wsWpx)
    weekOrder = SortedKeys(wpxWeekCols)

    Dim wpxRowByID As Object
    Set wpxRowByID = BuildWpxRowMap(wsWpx, wbWpx)

    ' Existing rows, top of the sheet down to the first blank name - anything
    ' below that (blank spacer, footnotes) is left alone.
    Dim lastDataRow As Long, rowByID As Object
    Set rowByID = CreateObject("Scripting.Dictionary")
    lastDataRow = ScanExistingRows(wsHist, cols, gtIDByName, rowByID)

    Dim k As Variant, playerID As String, wpxRow As Long, r As Long
    Dim stats As Object

    ' 1) Every row already on the sheet.
    For Each k In rowByID.Keys
        playerID = CStr(k)
        r = rowByID(playerID)

        If wpxRowByID.Exists(playerID) Then
            Set stats = BuildPlayerStats(wsWpx, wpxRowByID(playerID), wpxWeekCols, weekOrder)
            FinishStats stats, playerID, gtByID, wsHist, r, cols
            WriteRow wsHist, r, cols, weekTotalCols, weekRankCols, stats
            If stats(F_WEEKS) > 0 Then updated = updated + 1 Else noData = noData + 1
        Else
            noData = noData + 1
        End If
    Next k

    ' 2) GT players with WPX history that the sheet has no row for yet.
    Dim gtLast As Long
    gtLast = wsPT.Cells(wsPT.Rows.count, "A").End(xlUp).row

    For r = 2 To gtLast
        playerID = WpxNormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        If playerID <> "" Then
            If Not rowByID.Exists(playerID) Then
                If wpxRowByID.Exists(playerID) Then
                    Set stats = BuildPlayerStats(wsWpx, wpxRowByID(playerID), wpxWeekCols, weekOrder)

                    If stats(F_WEEKS) > 0 Then
                        lastDataRow = lastDataRow + 1
                        wsHist.Rows(lastDataRow).Insert Shift:=xlDown
                        FinishStats stats, playerID, gtByID, wsHist, lastDataRow, cols
                        WriteRow wsHist, lastDataRow, cols, weekTotalCols, weekRankCols, stats
                        rowByID.Add playerID, lastDataRow
                        added = added + 1
                    End If
                End If
            End If
        End If
    Next r


SafeExit:
    If Not wbWpx Is Nothing Then
        If opened Then wbWpx.Close SaveChanges:=False
    End If

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "RefreshWPXHistorySheet error " & Err.Number & ": " & Err.Description, vbExclamation
    ElseIf Not wsWpx Is Nothing Then
        Dim msg As String
        msg = HIST_SHEET & " filled from " & WPX_FILE & "." & vbCrLf & _
              "Rows updated: " & updated & vbCrLf & _
              "Rows added: " & added & vbCrLf & _
              "Rows with no WPX weeks found: " & noData
        If unknownHeaders <> "" Then _
            msg = msg & vbCrLf & vbCrLf & "Columns left untouched:" & unknownHeaders
        MsgBox msg, vbInformation, "Done"
    End If

End Sub

' Splits the header row into plain stat columns and per-week columns.
Private Sub MapHeaderColumns(ws As Worksheet, cols As Object, _
    weekTotalCols As Object, weekRankCols As Object, ByRef unknownHeaders As String)

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long, header As String, key As String, weekKey As String
    Dim isRank As Boolean
    For c = 1 To lastCol
        header = Trim$(CStr(ws.Cells(1, c).Value))
        If header <> "" Then
            weekKey = WeekKeyFromHeader(header, isRank)

            If weekKey <> "" Then
                If isRank Then
                    If Not weekRankCols.Exists(weekKey) Then weekRankCols.Add weekKey, c
                Else
                    If Not weekTotalCols.Exists(weekKey) Then weekTotalCols.Add weekKey, c
                End If
            Else
                key = StatKeyForHeader(header)
                If key = "" Then
                    unknownHeaders = unknownHeaders & vbCrLf & "  " & ColumnLetter(c) & ": " & header
                ElseIf Not cols.Exists(key) Then
                    cols.Add key, c
                End If
            End If
        End If
    Next c
End Sub

' Header text -> stat key. Matching is on words contained in the header, so
' "WPX Overall Total" and "Overall Total" both land on the same stat.
Private Function StatKeyForHeader(ByVal header As String) As String
    Dim h As String
    h = " " & LCase$(Trim$(Replace(header, Chr(160), " "))) & " "

    If Has(h, "combined") Then StatKeyForHeader = F_COMBINED: Exit Function
    If Has(h, "gt") And Has(h, "total") Then StatKeyForHeader = F_GTTOTAL: Exit Function
    If Has(h, "missed") And Has(h, "daily") Then StatKeyForHeader = F_MISSD: Exit Function
    If Has(h, "missed") And Has(h, "weekly") Then StatKeyForHeader = F_MISSW: Exit Function
    If Has(h, "avg") Or Has(h, "average") Then StatKeyForHeader = F_AVG: Exit Function
    If Has(h, "best") And (Has(h, "score") Or Has(h, "total")) Then StatKeyForHeader = F_BEST: Exit Function
    If Has(h, "best") Then StatKeyForHeader = F_BESTWK: Exit Function
    If Has(h, "weeks") Then StatKeyForHeader = F_WEEKS: Exit Function
    If Has(h, "rank") Then StatKeyForHeader = F_RANK: Exit Function
    If Has(h, "total") Then StatKeyForHeader = F_TOTAL: Exit Function
    If Has(h, "join") Then StatKeyForHeader = F_JOIN: Exit Function
    If Has(h, "first") Or Has(h, "start") Then StatKeyForHeader = F_FIRST: Exit Function
    If Has(h, "last") Or Has(h, "end") Then StatKeyForHeader = F_LAST: Exit Function
    If Has(h, "alliance") Then StatKeyForHeader = F_ALLIANCE: Exit Function
    If Has(h, "name") Or Has(h, "player") Then StatKeyForHeader = F_NAME: Exit Function
    If Has(h, "id") Then StatKeyForHeader = F_ID: Exit Function
End Function

Private Function Has(ByVal paddedHeader As String, ByVal word As String) As Boolean
    Has = InStr(1, paddedHeader, " " & word, vbTextCompare) > 0
End Function

' "WPX Weekly Total (Mar 23 - Mar 29)" -> week key, with isRank set for the
' matching "Weekly Rank" column. Non-weekly headers return "".
Private Function WeekKeyFromHeader(ByVal header As String, ByRef isRank As Boolean) As String
    isRank = False
    If InStr(1, header, "weekly", vbTextCompare) = 0 Then Exit Function

    Dim label As String
    label = WeekLabelFromHeader(header)
    If label = "" Then Exit Function
    If InStr(1, label, " - ") = 0 Then Exit Function

    isRank = InStr(1, header, "rank", vbTextCompare) > 0
    WeekKeyFromHeader = WeekKey(label)
End Function

' Existing player rows, from row 2 down to the first blank name. Returns that
' last data row and fills rowByID.
Private Function ScanExistingRows(ws As Worksheet, cols As Object, _
    gtIDByName As Object, rowByID As Object) As Long

    Dim nameCol As Long, idCol As Long
    nameCol = cols(F_NAME)
    If cols.Exists(F_ID) Then idCol = cols(F_ID)

    Dim r As Long, nm As String, playerID As String, nameKey As String
    r = 2
    Do While Trim$(CStr(ws.Cells(r, nameCol).Value)) <> ""
        nm = Trim$(CStr(ws.Cells(r, nameCol).Value))

        playerID = ""
        If idCol > 0 Then playerID = WpxNormalizeID(ws.Cells(r, idCol).Value)
        If playerID = "" Then
            nameKey = WpxNameKey(nm)
            If gtIDByName.Exists(nameKey) Then playerID = CStr(gtIDByName(nameKey))
        End If

        If playerID <> "" Then
            If Not rowByID.Exists(playerID) Then rowByID.Add playerID, r
        End If

        r = r + 1
    Loop

    ScanExistingRows = r - 1
End Function

' GT name and overall total per ID, plus name key -> ID.
Private Sub BuildGTMaps(wsPT As Worksheet, ByRef gtByID As Object, ByRef gtIDByName As Object)
    Set gtByID = CreateObject("Scripting.Dictionary")
    Set gtIDByName = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long, playerID As String, nm As String, nameKey As String
    lastRow = wsPT.Cells(wsPT.Rows.count, "A").End(xlUp).row

    For r = 2 To lastRow
        playerID = WpxNormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        nm = Trim$(CStr(wsPT.Cells(r, "A").Value))
        If playerID <> "" Then
            If Not gtByID.Exists(playerID) Then
                gtByID.Add playerID, Array(nm, wsPT.Cells(r, "C").Value)
            End If
            nameKey = WpxNameKey(nm)
            If nameKey <> "" Then
                If Not gtIDByName.Exists(nameKey) Then gtIDByName.Add nameKey, playerID
            End If
        End If
    Next r
End Sub

' WPX numbers for one player, from their row on the WPX Player Tracking sheet.
Private Function BuildPlayerStats(ws As Worksheet, wpxRow As Long, _
    wpxWeekCols As Object, weekOrder As Variant) As Object

    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")

    Dim weeklies As Object, weeklyRanks As Object
    Set weeklies = CreateObject("Scripting.Dictionary")
    Set weeklyRanks = CreateObject("Scripting.Dictionary")

    Dim total As Double, weeks As Long, best As Double
    Dim bestWeek As String, firstWeek As String, lastWeek As String
    Dim i As Long, key As String, label As String, v As Variant

    For i = LBound(weekOrder) To UBound(weekOrder)
        key = weekOrder(i)
        label = wpxWeekCols(key)(1)
        v = ws.Cells(wpxRow, wpxWeekCols(key)(0)).Value

        If HasScore(v) Then
            weeklies.Add key, CDbl(v)
            If wpxWeekCols(key)(2) > 0 Then
                If HasScore(ws.Cells(wpxRow, wpxWeekCols(key)(2)).Value) Then
                    weeklyRanks.Add key, CDbl(ws.Cells(wpxRow, wpxWeekCols(key)(2)).Value)
                End If
            End If
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

    Set s("weeklies") = weeklies
    Set s("weeklyRanks") = weeklyRanks
    s(F_FIRST) = firstWeek
    s(F_LAST) = lastWeek
    s(F_JOIN) = WeekStartDate(firstWeek)
    s(F_TOTAL) = total
    s(F_WEEKS) = weeks

    ' WPX's own average and overall rank are the official numbers; the computed
    ' average is only a fallback, and a rank cannot be recomputed here because
    ' WPX ranked every player in that alliance, not just the ones who moved.
    If HasScore(ws.Cells(wpxRow, "G").Value) Then
        s(F_AVG) = CDbl(ws.Cells(wpxRow, "G").Value)
    ElseIf weeks > 0 Then
        s(F_AVG) = total / weeks
    Else
        s(F_AVG) = ""
    End If
    s(F_RANK) = NumOrBlank(ws.Cells(wpxRow, "E").Value)
    s(F_BESTWK) = bestWeek
    If weeks > 0 Then s(F_BEST) = best Else s(F_BEST) = ""
    s(F_ALLIANCE) = ALLIANCE_TAG
    s(F_MISSD) = NumOrBlank(ws.Cells(wpxRow, "H").Value)
    s(F_MISSW) = NumOrBlank(ws.Cells(wpxRow, "I").Value)

    Set BuildPlayerStats = s
End Function

' Adds the ID, the GT name and the GT/combined totals, keeping whatever name is
' already on the row when GT has none.
Private Sub FinishStats(stats As Object, playerID As String, gtByID As Object, _
    ws As Worksheet, row As Long, cols As Object)

    stats(F_ID) = IDValue(playerID)

    Dim gtTotal As Variant
    gtTotal = ""

    If gtByID.Exists(playerID) Then
        stats(F_NAME) = gtByID(playerID)(0)
        gtTotal = gtByID(playerID)(1)
    Else
        stats(F_NAME) = Trim$(CStr(ws.Cells(row, cols(F_NAME)).Value))
    End If

    If HasScore(gtTotal) Then
        stats(F_GTTOTAL) = CDbl(gtTotal)
        stats(F_COMBINED) = CDbl(gtTotal) + stats(F_TOTAL)
    Else
        stats(F_GTTOTAL) = ""
        stats(F_COMBINED) = stats(F_TOTAL)
    End If
End Sub

Private Sub WriteRow(ws As Worksheet, row As Long, cols As Object, _
    weekTotalCols As Object, weekRankCols As Object, stats As Object)

    Dim k As Variant
    For Each k In cols.Keys
        If stats.Exists(k) Then
            If InStr(1, FILL_IF_BLANK, "|" & k & "|", vbTextCompare) > 0 Then
                If Trim$(CStr(ws.Cells(row, cols(k)).Value)) = "" Then
                    ws.Cells(row, cols(k)).Value = stats(k)
                End If
            Else
                ws.Cells(row, cols(k)).Value = stats(k)
            End If
        End If
    Next k

    Dim weeklies As Object, weeklyRanks As Object
    Set weeklies = stats("weeklies")
    Set weeklyRanks = stats("weeklyRanks")

    For Each k In weekTotalCols.Keys
        If weeklies.Exists(k) Then
            ws.Cells(row, weekTotalCols(k)).Value = weeklies(k)
        Else
            ws.Cells(row, weekTotalCols(k)).ClearContents
        End If
    Next k

    For Each k In weekRankCols.Keys
        If weeklyRanks.Exists(k) Then
            ws.Cells(row, weekRankCols(k)).Value = weeklyRanks(k)
        Else
            ws.Cells(row, weekRankCols(k)).ClearContents
        End If
    Next k
End Sub

' Week key -> Array(totalColumn, label, rankColumn) from the WPX Player
' Tracking header row.
Private Function BuildWpxWeekColumnMap(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long, header As String, label As String, key As String
    Dim isTotal As Boolean, isRank As Boolean, entry As Variant
    For c = 1 To lastCol
        header = Trim$(CStr(ws.Cells(1, c).Value))
        isTotal = InStr(1, header, "weekly total", vbTextCompare) > 0
        isRank = InStr(1, header, "weekly rank", vbTextCompare) > 0

        If isTotal Or isRank Then
            label = WeekLabelFromHeader(header)
            If label <> "" Then
                key = WeekKey(label)
                If key <> "" Then
                    If Not dict.Exists(key) Then dict.Add key, Array(0, label, 0)
                    entry = dict(key)
                    If isTotal Then entry(0) = c Else entry(2) = c
                    dict(key) = entry
                End If
            End If
        End If
    Next c

    ' Weeks with no total column carry no score to copy.
    Dim k As Variant
    For Each k In dict.Keys
        If dict(k)(0) = 0 Then dict.Remove k
    Next k

    Set BuildWpxWeekColumnMap = dict
End Function

Private Function SortedKeys(dict As Object) As Variant
    Dim keys As Variant
    keys = dict.Keys

    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(j) < keys(i) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i

    SortedKeys = keys
End Function

' "WPX Weekly Total (Mar 23 - Mar 29)" -> "Mar 23 - Mar 29"
Private Function WeekLabelFromHeader(ByVal header As String) As String
    Dim openPos As Long, closePos As Long
    openPos = InStr(header, "(")
    closePos = InStrRev(header, ")")
    If openPos = 0 Or closePos <= openPos + 1 Then Exit Function
    WeekLabelFromHeader = Trim$(Mid$(header, openPos + 1, closePos - openPos - 1))
End Function

Private Function WeekKey(ByVal label As String) As String
    Dim d As Date
    d = WeekStartDate(label)
    If d = 0 Then
        WeekKey = UCase$(Trim$(label))
    Else
        WeekKey = Format$(d, "mm-dd")
    End If
End Function

' Week labels carry no year, so a date in the future is rolled back one.
Private Function WeekStartDate(ByVal label As String) As Date
    If label = "" Then Exit Function

    Dim startText As String, d As Date
    startText = label
    If InStr(1, label, " - ") > 0 Then startText = Trim$(Split(label, " - ")(0))

    On Error Resume Next
    d = DateValue(startText & ", " & Year(Date))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If d > Date + 3 Then d = DateSerial(Year(d) - 1, Month(d), Day(d))
    WeekStartDate = d
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

' IDs are written back as numbers where they are all digits, so they sort and
' match the way the hand-entered ones already on the sheet do.
Private Function IDValue(ByVal playerID As String) As Variant
    If playerID Like String(Len(playerID), "#") And Len(playerID) > 0 And Len(playerID) < 15 Then
        IDValue = CDbl(playerID)
    Else
        IDValue = playerID
    End If
End Function

Private Function ColumnLetter(ByVal col As Long) As String
    ColumnLetter = Split(Cells(1, col).Address(True, False), "$")(0)
End Function

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
