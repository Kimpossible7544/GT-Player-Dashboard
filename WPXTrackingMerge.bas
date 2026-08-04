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
'
' To have it run by itself every time the workbook opens, run
' InstallWPXAutoRun() once - it writes a Workbook_Open handler into the
' ThisWorkbook module (RemoveWPXAutoRun() takes it back out). Excel only allows
' that with "Trust access to the VBA project object model" ticked under File >
' Options > Trust Center > Trust Center Settings > Macro Settings; without it,
' the installer just shows you the four lines to paste into ThisWorkbook.
'
' The auto-run is silent: no popups, and if WPXStatsFinal.xlsm is not open and
' not sitting next to GTStatsFINAL.xlsm it is skipped rather than prompting for
' the file and stalling the open.
'
' IDs are compared as digits only, so an ID typed with commas or spaces, stored
' in scientific notation, or padded with leading zeros still matches. Players
' whose ID is missing or mistyped in one of the workbooks fall back to matching
' on player name (set MATCH_BY_NAME to False for ID-only matching).
'
' If players are missing from the result, run DiagnoseWPXTrackingMerge(). It
' builds a "WPX Merge Diagnostics" sheet listing every GT Player Tracking row
' with its ID, the WPX row it matched, how it matched, and the reason nothing
' was filled (no WPX row, no shared week column, or GT already has a score).

Option Explicit

Private Const PT_SHEET      As String = "Player Tracking"
Private Const PT_ID_COL     As String = "CL"
Private Const WPX_PT_SHEET  As String = "Player Tracking"
Private Const WPX_FILE      As String = "WPXStatsFinal.xlsm"

' A GT cell holding a real score is left alone; set this to False to let WPX
' numbers overwrite whatever the GT refresh wrote.
Private Const FILL_ONLY_BLANKS As Boolean = True

' Players whose ID is missing or mistyped in one of the workbooks are matched on
' player name instead. Set to False for ID-only matching.
Private Const MATCH_BY_NAME As Boolean = True

Private Const AUTORUN_TAG As String = "' --- WPX auto-run (WPXTrackingMerge) ---"

' Set while running from Workbook_Open: no message boxes, no file prompt.
Private m_silent As Boolean

' Entry point for the Workbook_Open handler.
Public Sub RunWPXTrackingOnOpen()
    m_silent = True
    On Error Resume Next
    MergeWPXTrackingIntoPlayerTracking
    On Error GoTo 0
    m_silent = False
End Sub

' Writes the Workbook_Open handler into ThisWorkbook so the merge runs on every
' open. Needs "Trust access to the VBA project object model"; without it the
' handler is shown for pasting in by hand.
Public Sub InstallWPXAutoRun()
    Dim code As String
    code = AUTORUN_TAG & vbCrLf & _
           "Private Sub Workbook_Open()" & vbCrLf & _
           "    RunWPXTrackingOnOpen" & vbCrLf & _
           "End Sub"

    Dim cm As Object
    On Error Resume Next
    Set cm = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    On Error GoTo 0

    If cm Is Nothing Then
        MsgBox "Excel is blocking macro access to the VBA project." & vbCrLf & vbCrLf & _
               "Either tick File > Options > Trust Center > Trust Center Settings > " & _
               "Macro Settings > ""Trust access to the VBA project object model"" and " & _
               "run this again, or paste these lines into the ThisWorkbook module " & _
               "yourself (Alt+F11 > double-click ThisWorkbook):" & vbCrLf & vbCrLf & code, _
               vbExclamation, "WPX auto-run"
        Exit Sub
    End If

    If FindAutoRunLine(cm) > 0 Then
        MsgBox "The WPX auto-run is already installed.", vbInformation, "WPX auto-run"
        Exit Sub
    End If

    If HasWorkbookOpen(cm) Then
        MsgBox "ThisWorkbook already has its own Workbook_Open." & vbCrLf & vbCrLf & _
               "Add this line inside it instead:" & vbCrLf & vbCrLf & _
               "    RunWPXTrackingOnOpen", vbExclamation, "WPX auto-run"
        Exit Sub
    End If

    cm.AddFromString code
    MsgBox "WPX auto-run installed - the merge now runs when " & ThisWorkbook.name & _
           " opens. Save the workbook to keep it.", vbInformation, "WPX auto-run"
End Sub

' Takes the generated Workbook_Open handler back out.
Public Sub RemoveWPXAutoRun()
    Dim cm As Object
    On Error Resume Next
    Set cm = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
    On Error GoTo 0

    If cm Is Nothing Then
        MsgBox "Excel is blocking macro access to the VBA project - delete the " & _
               "Workbook_Open handler in ThisWorkbook by hand.", vbExclamation, "WPX auto-run"
        Exit Sub
    End If

    Dim tagLine As Long
    tagLine = FindAutoRunLine(cm)
    If tagLine = 0 Then
        MsgBox "No WPX auto-run handler found in ThisWorkbook.", vbInformation, "WPX auto-run"
        Exit Sub
    End If

    ' The tag line plus the three lines of the handler it precedes.
    cm.DeleteLines tagLine, 4
    MsgBox "WPX auto-run removed. Save the workbook to keep it.", _
           vbInformation, "WPX auto-run"
End Sub

Private Function FindAutoRunLine(cm As Object) As Long
    Dim i As Long
    For i = 1 To cm.CountOfLines
        If InStr(1, cm.Lines(i, 1), AUTORUN_TAG, vbTextCompare) > 0 Then
            FindAutoRunLine = i
            Exit Function
        End If
    Next i
End Function

Private Function HasWorkbookOpen(cm As Object) As Boolean
    Dim i As Long, txt As String
    For i = 1 To cm.CountOfLines
        txt = LCase$(Trim$(cm.Lines(i, 1)))
        If InStr(txt, "sub workbook_open(") > 0 Then
            HasWorkbookOpen = True
            Exit Function
        End If
    Next i
End Function

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
        Notify "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    Set wbWpx = OpenWpxWorkbook(opened)
    If wbWpx Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsWpx = wbWpx.Worksheets(WPX_PT_SHEET)
    On Error GoTo SafeExit
    If wsWpx Is Nothing Then
        Notify "Sheet """ & WPX_PT_SHEET & """ not found in " & wbWpx.name & ".", vbExclamation
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
        Notify "No ""Weekly Total (...)"" headers found on " & PT_SHEET & _
               ". Run the GT Player Tracking refresh first.", vbExclamation
        GoTo SafeExit
    End If
    If wpxWeekCols.count = 0 Then
        Notify "No ""Weekly Total (...)"" headers found on the WPX " & _
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

    Dim wpxRowByID As Object, wpxRowByName As Object
    Set wpxRowByID = BuildWpxRowMap(wsWpx, wbWpx)
    Set wpxRowByName = BuildWpxRowByName(wsWpx)

    Dim gtIDCol As Long
    gtIDCol = FindIDColumn(wsPT)

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
        playerID = ""
        If gtIDCol > 0 Then playerID = WpxNormalizeID(wsPT.Cells(r, gtIDCol).Value)
        If playerID <> "" Or MATCH_BY_NAME Then
            wpxRow = ResolveWpxRow(playerID, wsPT.Cells(r, 1).Value, wpxRowByID, wpxRowByName)
            If wpxRow > 0 Then
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
            ElseIf playerID <> "" Then
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
        Notify "MergeWPXTrackingIntoPlayerTracking error " & Err.Number & ": " & _
               Err.Description, vbExclamation
    ElseIf Not wsWpx Is Nothing Then
        Notify "WPX weeks pulled into " & PT_SHEET & "." & vbCrLf & _
               "Players filled: " & matchedPlayers & vbCrLf & _
               "Week cells filled: " & filledCells & vbCrLf & _
               "Players with an ID but no WPX row: " & unmatchedPlayers & vbCrLf & _
               "WPX weeks with no GT column: " & unmatchedWeeks & missingWeekList & vbCrLf & vbCrLf & _
               "Players missing? Run DiagnoseWPXTrackingMerge for a per-player reason.", _
               vbInformation, "Done"
    End If

End Sub

' Reports, per GT Player Tracking row, whether a WPX row was found, how it was
' matched, and how many WPX weeks were usable. Writes the "WPX Merge
' Diagnostics" sheet and changes no data.
Public Sub DiagnoseWPXTrackingMerge()
    Dim wsPT As Worksheet, wsWpx As Worksheet
    Dim wbWpx As Workbook
    Dim opened As Boolean

    On Error Resume Next
    Set wsPT = ThisWorkbook.Worksheets(PT_SHEET)
    On Error GoTo 0
    If wsPT Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    Set wbWpx = OpenWpxWorkbook(opened)
    If wbWpx Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsWpx = wbWpx.Worksheets(WPX_PT_SHEET)
    On Error GoTo 0
    If wsWpx Is Nothing Then
        MsgBox "Sheet """ & WPX_PT_SHEET & """ not found in " & wbWpx.name & ".", vbExclamation
        If opened Then wbWpx.Close SaveChanges:=False
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim gtWeekCols As Object, wpxWeekCols As Object
    Set gtWeekCols = BuildWeekColumnMap(wsPT)
    Set wpxWeekCols = BuildWeekColumnMap(wsWpx)

    Dim wpxRowByID As Object, wpxRowByName As Object
    Set wpxRowByID = BuildWpxRowMap(wsWpx, wbWpx)
    Set wpxRowByName = BuildWpxRowByName(wsWpx)

    Dim gtIDCol As Long
    gtIDCol = FindIDColumn(wsPT)

    Dim ws As Worksheet
    Set ws = GetOrCreateDiagSheet("WPX Merge Diagnostics")
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("GT Player", "GT Row", "GT Player ID", "WPX Row", _
        "Matched By", "WPX Weeks With A Shared GT Column", "Result")
    ws.Rows(1).Font.Bold = True

    Dim lastRow As Long, r As Long, outRow As Long
    lastRow = wsPT.Cells(wsPT.Rows.count, "A").End(xlUp).row
    outRow = 1

    Dim gtName As String, playerID As String, wpxRow As Long
    Dim k As Variant, usable As Long, blocked As Long

    For r = 2 To lastRow
        gtName = Trim$(CStr(wsPT.Cells(r, 1).Value))
        playerID = ""
        If gtIDCol > 0 Then playerID = WpxNormalizeID(wsPT.Cells(r, gtIDCol).Value)
        If gtName <> "" Or playerID <> "" Then
            outRow = outRow + 1
            ws.Cells(outRow, 1).Value = gtName
            ws.Cells(outRow, 2).Value = r
            ws.Cells(outRow, 3).Value = playerID

            wpxRow = 0
            If playerID <> "" Then
                If wpxRowByID.Exists(playerID) Then
                    wpxRow = wpxRowByID(playerID)
                    ws.Cells(outRow, 5).Value = "Player ID"
                End If
            End If
            If wpxRow = 0 And MATCH_BY_NAME Then
                Dim nk As String
                nk = WpxNameKey(gtName)
                If nk <> "" Then
                    If wpxRowByName.Exists(nk) Then
                        wpxRow = wpxRowByName(nk)
                        ws.Cells(outRow, 5).Value = "Name"
                    End If
                End If
            End If

            If wpxRow = 0 Then
                If playerID = "" Then
                    ws.Cells(outRow, 7).Value = "Skipped - no Player ID on the GT row and no WPX row with that name"
                Else
                    ws.Cells(outRow, 7).Value = "Skipped - ID not found on the WPX Player Tracking sheet"
                End If
            Else
                ws.Cells(outRow, 4).Value = wpxRow
                usable = 0
                blocked = 0
                For Each k In gtWeekCols.Keys
                    If wpxWeekCols.Exists(k) Then
                        If HasScore(wsWpx.Cells(wpxRow, wpxWeekCols(k)(0)).Value) Then
                            If FILL_ONLY_BLANKS And HasScore(wsPT.Cells(r, gtWeekCols(k)(0)).Value) Then
                                blocked = blocked + 1
                            Else
                                usable = usable + 1
                            End If
                        End If
                    End If
                Next k
                ws.Cells(outRow, 6).Value = usable
                If usable > 0 Then
                    ws.Cells(outRow, 7).Value = "Matched - " & usable & " week(s) filled"
                ElseIf blocked > 0 Then
                    ws.Cells(outRow, 7).Value = "Matched - nothing filled, GT already has a score in all " & _
                                                blocked & " shared week(s)"
                Else
                    ws.Cells(outRow, 7).Value = "Matched - no WPX week lines up with a GT week column"
                End If
            End If
        End If
    Next r

    ws.Columns("A:G").AutoFit
    ws.Activate

    If opened Then wbWpx.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "Diagnostics written to the ""WPX Merge Diagnostics"" sheet.", vbInformation
End Sub

Private Function GetOrCreateDiagSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = sheetName
    End If
    Set GetOrCreateDiagSheet = ws
End Function

' The WPX row for a player: by ID first, then by name when MATCH_BY_NAME is on,
' so a missing or mistyped ID in either workbook still matches.
Private Function ResolveWpxRow(playerID As String, gtName As Variant, _
    wpxRowByID As Object, wpxRowByName As Object) As Long

    If playerID <> "" Then
        If wpxRowByID.Exists(playerID) Then
            ResolveWpxRow = wpxRowByID(playerID)
            Exit Function
        End If
    End If

    If Not MATCH_BY_NAME Then Exit Function

    Dim key As String
    key = WpxNameKey(gtName)
    If key = "" Then Exit Function
    If wpxRowByName.Exists(key) Then ResolveWpxRow = wpxRowByName(key)
End Function

' Player name -> row on a Player Tracking sheet, names in column A.
Private Function BuildWpxRowByName(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long, key As String
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    For r = 2 To lastRow
        key = WpxNameKey(ws.Cells(r, 1).Value)
        If key <> "" Then
            If Not dict.Exists(key) Then dict.Add key, r
        End If
    Next r

    Set BuildWpxRowByName = dict
End Function

' MsgBox, unless we are running from Workbook_Open.
Private Sub Notify(ByVal msg As String, ByVal style As VbMsgBoxStyle, _
    Optional ByVal title As String = "WPX Merge")

    If m_silent Then Exit Sub
    MsgBox msg, style, title
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
    idCol = FindIDColumn(ws)

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

' CL when any row of it holds an ID, otherwise the first "Player ID" header.
Private Function FindIDColumn(ws As Worksheet) As Long
    Dim clCol As Long
    clCol = ws.Range(PT_ID_COL & "1").Column
    If ColumnHasID(ws, clCol) Then
        FindIDColumn = clCol
        Exit Function
    End If

    Dim lastCol As Long, c As Long, header As String
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        header = LCase$(Trim$(CStr(ws.Cells(1, c).Value)))
        If header = "player id" Or header = "id" Then
            FindIDColumn = c
            Exit Function
        End If
    Next c

    FindIDColumn = clCol
End Function

' True when the column holds an ID somewhere in its first rows, so a blank
' leading row does not make the column look empty.
Private Function ColumnHasID(ws As Worksheet, col As Long) As Boolean
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).row
    For r = 2 To lastRow
        If WpxNormalizeID(ws.Cells(r, col).Value) <> "" Then
            ColumnHasID = True
            Exit Function
        End If
    Next r
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

' Roster IDs arrive as numbers, as text with commas or spaces, and sometimes in
' scientific notation, so they are reduced to a digits-only key. IDs are kept as
' text so values longer than a Long cannot overflow. Non-numeric IDs fall back
' to their uppercased text.
Private Function WpxNormalizeID(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim txt As String
    txt = Trim$(CStr(v))
    txt = Replace(txt, Chr(160), "")
    txt = Replace(txt, ChrW(8203), "")
    txt = Replace(txt, ChrW(65279), "")
    txt = Replace(txt, " ", "")
    txt = Replace(txt, ",", "")

    If IsNumeric(txt) Then
        If InStr(1, txt, "E", vbTextCompare) > 0 Then txt = Format$(CDbl(txt), "0")
    End If

    Dim digits As String, i As Long, ch As String
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    If digits = "" Then
        WpxNormalizeID = UCase$(txt)
        Exit Function
    End If

    Do While Len(digits) > 1 And Left$(digits, 1) = "0"
        digits = Mid$(digits, 2)
    Loop

    If digits = "0" Then Exit Function
    WpxNormalizeID = digits
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
    ElseIf m_silent Then
        ' Opening the workbook must not stall on a file picker.
        Exit Function
    Else
        f = Application.GetOpenFilename( _
            "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , "Select " & WPX_FILE)
        If VarType(f) = vbBoolean Then Exit Function
        If CStr(f) = "False" Then Exit Function
    End If

    Set OpenWpxWorkbook = Workbooks.Open(CStr(f), ReadOnly:=True)
    opened = True
End Function
