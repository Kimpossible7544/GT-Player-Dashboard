Attribute VB_Name = "WPXTrackingMerge"
' WPXTrackingMerge.bas
' Fills the GT "Player Tracking" sheet with a player's WPX weeks, read straight
' off the "Player Tracking" sheet in WPXStatsFinal.xlsm and matched on the
' Player ID in column CL - the same ID UpdatePlayerTrackingFromRosterAndWeeks
' uses.
'
' Each WPX week lands in the GT column for that same week: both sheets label
' their weekly columns "Weekly Total (Jul 13 - Jul 19)", so the two sets of
' headers are paired by week, not by column position. Only cells the GT refresh
' left empty are filled, so a week the player actually played in GT is never
' overwritten by their WPX number. WPX weeks with no matching GT column are
' reported in the summary and left alone.
'
' After filling, the weekly ranks of the touched columns plus Overall Total (C),
' Overall Rank (E) and Overall Weekly Average (G) are recalculated so they
' include the WPX weeks. Pushing Total (D), Pushing Rank (F) and the missed-goal
' counts (H, I) are GT-season figures and are left as the GT macro wrote them.
'
' To use:
'   1. Open GTStatsFINAL.xlsm.
'   2. Alt+F11 > File > Import File > WPXTrackingMerge.bas
'   3. Run AddWPXTrackingButton() once for a "Pull WPX Weeks" button on the
'      Player Tracking sheet, then click it after each GT refresh.
'   4. Pick WPXStatsFinal.xlsm when prompted (skipped when it sits in the same
'      folder as GTStatsFINAL.xlsm).
'
' To refresh GT and WPX in one go, add this line at the end of
' UpdatePlayerTrackingFromRosterAndWeeks, just before SafeExit:
'
'     Call MergeWPXTrackingIntoPlayerTracking

Option Explicit

Private Const PT_SHEET      As String = "Player Tracking"
Private Const PT_ID_COL     As String = "CL"
Private Const WPX_PT_SHEET  As String = "Player Tracking"
Private Const WPX_FILE      As String = "WPXStatsFinal.xlsm"

' A GT cell holding a real score is left alone; set this to False to let WPX
' numbers overwrite whatever the GT refresh wrote.
Private Const FILL_ONLY_BLANKS As Boolean = True

' Puts a "Pull WPX Weeks" button on the Player Tracking sheet.
Public Sub AddWPXTrackingButton()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ws.Shapes("btnPullWPXTracking").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 8, 8, 150, 28)
    shp.name = "btnPullWPXTracking"
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
    shp.OnAction = "MergeWPXTrackingIntoPlayerTracking"

    ws.Activate
    MsgBox "Button added to the " & PT_SHEET & " sheet.", vbInformation
End Sub

' Copies each player's WPX weekly totals into the matching GT week columns.
Public Sub MergeWPXTrackingIntoPlayerTracking()

    Dim wsPT As Worksheet, wsWpx As Worksheet
    Dim wbWpx As Workbook
    Dim oldCalc As XlCalculation
    Dim opened As Boolean
    Dim filledCells As Long, matchedPlayers As Long
    Dim unmatchedPlayers As Long, unmatchedWeeks As Long
    Dim missingWeekList As String

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

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Week key -> column, for both sheets. The key is the week's start date, so
    ' "Jul 13 - Jul 19" and "Jul 13 - Jul 17" still pair up.
    Dim gtWeekCols As Object, wpxWeekCols As Object
    Set gtWeekCols = BuildWeekColumnMap(wsPT)
    Set wpxWeekCols = BuildWeekColumnMap(wsWpx)

    If gtWeekCols.count = 0 Then
        MsgBox "No ""Weekly Total (...)"" headers found on " & PT_SHEET & _
               ". Run the GT Player Tracking refresh first.", vbExclamation
        GoTo SafeExit
    End If
    If wpxWeekCols.count = 0 Then
        MsgBox "No ""Weekly Total (...)"" headers found on the WPX " & _
               WPX_PT_SHEET & " sheet.", vbExclamation
        GoTo SafeExit
    End If

    ' Weeks the player was in WPX that the GT sheet has no column for.
    Dim k As Variant
    For Each k In wpxWeekCols.Keys
        If Not gtWeekCols.Exists(k) Then
            unmatchedWeeks = unmatchedWeeks + 1
            If Len(missingWeekList) < 200 Then
                missingWeekList = missingWeekList & vbCrLf & "  " & wpxWeekCols(k)(1)
            End If
        End If
    Next k

    Dim wpxRowByID As Object
    Set wpxRowByID = BuildWpxRowMap(wsWpx, wbWpx)

    Dim firstRow As Long, lastRow As Long
    firstRow = 2
    lastRow = wsPT.Cells(wsPT.Rows.count, "A").End(xlUp).row
    If lastRow < firstRow Then GoTo SafeExit

    Dim touched As Object
    Set touched = CreateObject("Scripting.Dictionary")

    Dim r As Long, playerID As String, wpxRow As Long
    Dim gtCol As Long, wpxCol As Long, v As Variant, gtVal As Variant
    Dim playerFilled As Long

    For r = firstRow To lastRow
        playerID = WpxNormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        If playerID <> "" Then
            If wpxRowByID.Exists(playerID) Then
                wpxRow = wpxRowByID(playerID)
                playerFilled = 0

                For Each k In gtWeekCols.Keys
                    If wpxWeekCols.Exists(k) Then
                        gtCol = gtWeekCols(k)(0)
                        wpxCol = wpxWeekCols(k)(0)
                        gtVal = wsPT.Cells(r, gtCol).Value

                        If (Not FILL_ONLY_BLANKS) Or Not HasScore(gtVal) Then
                            v = wsWpx.Cells(wpxRow, wpxCol).Value
                            If HasScore(v) Then
                                wsPT.Cells(r, gtCol).Value = CDbl(v)
                                touched(gtCol) = gtWeekCols(k)(2)
                                playerFilled = playerFilled + 1
                            End If
                        End If
                    End If
                Next k

                If playerFilled > 0 Then
                    matchedPlayers = matchedPlayers + 1
                    filledCells = filledCells + playerFilled
                End If
            Else
                unmatchedPlayers = unmatchedPlayers + 1
            End If
        End If
    Next r

    If filledCells > 0 Then
        For Each k In touched.Keys
            FillRankColumn wsPT, CLng(k), CLng(touched(k)), firstRow, lastRow
        Next k
        RecalcOverall wsPT, gtWeekCols, firstRow, lastRow
    End If

SafeExit:
    If Not wbWpx Is Nothing Then
        If opened Then wbWpx.Close SaveChanges:=False
    End If

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "MergeWPXTrackingIntoPlayerTracking error " & Err.Number & ": " & _
               Err.Description, vbExclamation
    ElseIf Not wsWpx Is Nothing Then
        MsgBox "WPX weeks pulled into " & PT_SHEET & "." & vbCrLf & _
               "Players filled: " & matchedPlayers & vbCrLf & _
               "Week cells filled: " & filledCells & vbCrLf & _
               "Players with an ID but no WPX row: " & unmatchedPlayers & vbCrLf & _
               "WPX weeks with no GT column: " & unmatchedWeeks & missingWeekList, _
               vbInformation, "Done"
    End If

End Sub

' Week start date -> Array(total column, week label, rank column), read from the
' "Weekly Total (...)" headers in row 1.
Private Function BuildWeekColumnMap(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long, header As String, label As String, key As String
    For c = 1 To lastCol
        header = Trim$(CStr(ws.Cells(1, c).Value))
        label = WeekLabelFromHeader(header)
        If label <> "" Then
            key = WeekKey(label)
            If key <> "" Then
                If Not dict.Exists(key) Then
                    dict.Add key, Array(c, label, RankColumnFor(ws, c, label, lastCol))
                End If
            End If
        End If
    Next c

    Set BuildWeekColumnMap = dict
End Function

' The rank column that goes with a weekly total column: the next column when it
' is that week's "Weekly Rank (...)" header, otherwise none (0).
Private Function RankColumnFor(ws As Worksheet, totalCol As Long, _
    label As String, lastCol As Long) As Long

    If totalCol >= lastCol Then Exit Function

    Dim header As String
    header = Trim$(CStr(ws.Cells(1, totalCol + 1).Value))
    If InStr(1, header, "Weekly Rank", vbTextCompare) = 1 Then
        If StrComp(WeekLabelFromHeader(header), label, vbTextCompare) = 0 Then
            RankColumnFor = totalCol + 1
        End If
    End If
End Function

' "Weekly Total (Jul 13 - Jul 19)" -> "Jul 13 - Jul 19"
Private Function WeekLabelFromHeader(ByVal header As String) As String
    Dim openPos As Long, closePos As Long

    If InStr(1, header, "Weekly Total", vbTextCompare) <> 1 Then
        If InStr(1, header, "Weekly Rank", vbTextCompare) <> 1 Then Exit Function
    End If

    openPos = InStr(header, "(")
    closePos = InStrRev(header, ")")
    If openPos = 0 Or closePos <= openPos + 1 Then Exit Function

    WeekLabelFromHeader = Trim$(Mid$(header, openPos + 1, closePos - openPos - 1))
End Function

' Week labels carry no year and the two workbooks may end their weeks on
' different days, so weeks are keyed on their start day/month where it parses.
' The year is dropped deliberately: the sheet names never carry one.
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

' Player ID -> row on the WPX Player Tracking sheet. IDs live in CL there too;
' when that column is empty the ID column is found by header, and rows with no
' ID at all are resolved through the WPX roster by name.
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

' Overall Total (C), Overall Rank (E) and Overall Weekly Average (G), now that
' the WPX weeks are part of the row. Pushing figures stay as the GT macro left
' them - those weeks are GT-season only.
Private Sub RecalcOverall(ws As Worksheet, weekCols As Object, _
    firstRow As Long, lastRow As Long)

    Dim r As Long, k As Variant, v As Variant
    Dim totalSum As Double, weekCount As Long

    For r = firstRow To lastRow
        totalSum = 0
        weekCount = 0

        For Each k In weekCols.Keys
            v = ws.Cells(r, weekCols(k)(0)).Value
            If HasScore(v) Then
                totalSum = totalSum + CDbl(v)
                weekCount = weekCount + 1
            End If
        Next k

        If weekCount > 0 Then
            ws.Cells(r, "C").Value = totalSum
            ws.Cells(r, "G").Value = totalSum / weekCount
        End If
    Next r

    FillRankColumn ws, 3, 5, firstRow, lastRow
End Sub

' Highest score = rank 1, over the players who have a score in valueCol.
Private Sub FillRankColumn(ws As Worksheet, valueCol As Long, rankCol As Long, _
    firstRow As Long, lastRow As Long)

    If rankCol = 0 Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(firstRow, valueCol), ws.Cells(lastRow, valueCol))

    Dim count As Long
    count = Application.WorksheetFunction.count(rng)

    Dim r As Long, v As Variant
    For r = firstRow To lastRow
        v = ws.Cells(r, valueCol).Value
        If count > 0 And HasScore(v) Then
            ws.Cells(r, rankCol).Value = Application.WorksheetFunction.rank(v, rng, 0)
        Else
            ws.Cells(r, rankCol).ClearContents
        End If
    Next r
End Sub

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
