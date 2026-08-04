Attribute VB_Name = "WPXTrackingMerge"
' WPXTrackingMerge.bas
' Puts each player's legacy WPX weeks onto the GT "Player Tracking" sheet, in the
' column for the week they were scored, matched on the Player ID in column CL -
' the same ID UpdatePlayerTrackingFromRosterAndWeeks writes.
'
' The WPX numbers are written as ID-keyed formulas, not as pasted values. That
' matters because UpdatePlayerTrackingFromRosterAndWeeks clears A2:A500, B2:B500
' and CL2:CL500 and rewrites names and IDs in roster order, while leaving the
' WPX columns alone: pasted values stay glued to their old row numbers and end
' up beside whoever now sits in that row. A formula keyed on CL follows its
' player through any roster reshuffle instead.
'
' Three pieces are built, all rebuilt from scratch on each run:
'   "WPX Snapshot"  hidden sheet - one row per WPX player (ID, name, one column
'                   per WPX week), copied out of WPXStatsFinal.xlsm so the GT
'                   workbook keeps working when the WPX file is not open.
'   "WPX ID Map"    crosswalk for players whose GT roster ID differs from their
'                   WPX one - the transfers that were re-added to the roster and
'                   given a new ID. Put the GT ID in column A and their WPX
'                   player ID *or* their WPX player name in column B. Every
'                   player with a note in the "WPX Transfer" column (CN) that no
'                   WPX row can be found for is pre-listed here for you to fill
'                   in, so flag transfers there as you re-add them.
'   Column CO       hidden helper on Player Tracking: the WPX key for the row,
'                   i.e. the mapped ID when there is one, else the GT ID.
'
' Only weeks GT itself never played are filled: a GT week column is treated as
' WPX-owned when the workbook has no weekly sheet of that name, so a week the
' player actually played in GT is never overwritten. A WPX week with no GT column
' at all gets one, appended after the last existing week column so nothing
' shifts, and every WPX week is also mirrored into the "WPX Weekly ..." block
' from CV onwards, which is past everything the GT refresh clears.
'
' This module owns both the pre-GT week columns and the CV block. Do not also run
' WPXWeeksMerge or PullWPXPlayerTracking against the same sheet: two macros
' pasting into the same columns is how the numbers drifted onto the wrong players
' in the first place.
'
' The overall figures in C-I are left exactly as the GT refresh wrote them, so
' every ranking on the sheet and on the dashboard stays GT-season only. The WPX
' side gets its own summary in the CV block - "WPX Overall Total", "WPX Weeks
' Played", "WPX Weekly Average" and "WPX Overall Rank", the last two ranking WPX
' players against each other.
'
' To use:
'   1. Open GTStatsFINAL.xlsm.
'   2. Alt+F11 > File > Import File > WPXTrackingMerge.bas
'   3. Run AddWPXTrackingButton() once for a "Pull WPX Weeks" button on the
'      Player Tracking sheet, then click it after each GT refresh.
'   4. Pick WPXStatsFinal.xlsm when prompted (skipped when it sits in the same
'      folder as GTStatsFINAL.xlsm).
'   5. Fill in any rows the macro added to "WPX ID Map" and click the button
'      again.
'
' To refresh GT and WPX in one go, add this line at the end of
' UpdatePlayerTrackingFromRosterAndWeeks, just before SafeExit:
'
'     Call RefreshWPXTracking
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
' in scientific notation, or padded with leading zeros still matches.
'
' ClearWPXTracking() strips every WPX column back out, which is also the way to
' clear stale pasted values left behind by an older version of this macro.
' DiagnoseWPXTrackingMerge() writes a "WPX Merge Diagnostics" sheet listing each
' player, the WPX row they resolved to, how they matched, and how many weeks
' were filled.

Option Explicit

Private Const PT_SHEET      As String = "Player Tracking"
Private Const PT_ID_COL     As String = "CL"
Private Const TRANSFER_COL  As String = "CN"    ' "WPX Transfer" note column
Private Const HELPER_COL    As String = "CO"
Private Const WPX_BLOCK_COL As String = "CV"
Private Const SNAP_SHEET    As String = "WPX Snapshot"
Private Const MAP_SHEET     As String = "WPX ID Map"
Private Const DIAG_SHEET    As String = "WPX Merge Diagnostics"
Private Const WPX_PT_SHEET  As String = "Player Tracking"
Private Const WPX_FILE      As String = "WPXStatsFinal.xlsm"

' Rows the GT refresh itself writes and clears, so the formulas cover the same
' span it does.
Private Const LAST_PT_ROW   As Long = 500

Private Const AUTORUN_TAG As String = "' --- WPX auto-run (WPXTrackingMerge) ---"

' Set while running from Workbook_Open: no message boxes, no file prompt.
Private m_silent As Boolean

'==============================================================================
' Entry points
'==============================================================================

' Rebuilds the snapshot and every WPX column on Player Tracking.
Public Sub RefreshWPXTracking()
    Dim wsPT As Worksheet, wsWpx As Worksheet
    Dim wbWpx As Workbook
    Dim opened As Boolean
    Dim oldCalc As XlCalculation
    Dim calcSaved As Boolean

    On Error GoTo SafeExit

    Set wsPT = SheetOrNothing(ThisWorkbook, PT_SHEET)
    If wsPT Is Nothing Then
        Notify "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    Set wbWpx = OpenWpxWorkbook(opened)
    If wbWpx Is Nothing Then Exit Sub

    Set wsWpx = SheetOrNothing(wbWpx, WPX_PT_SHEET)
    If wsWpx Is Nothing Then
        Notify "Sheet """ & WPX_PT_SHEET & """ not found in " & wbWpx.name & ".", vbExclamation
        GoTo SafeExit
    End If

    oldCalc = Application.Calculation
    calcSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Week start date -> Array(total column, label, rank column) on each sheet.
    Dim gtWeekCols As Object, wpxWeekCols As Object
    Set gtWeekCols = BuildWeekColumnMap(wsPT)
    Set wpxWeekCols = BuildWeekColumnMap(wsWpx)

    If wpxWeekCols.count = 0 Then
        Notify "No ""Weekly Total (...)"" headers found on the WPX " & _
               WPX_PT_SHEET & " sheet.", vbExclamation
        GoTo SafeExit
    End If

    ' Week key -> snapshot column letter.
    Dim snapCols As Object
    Set snapCols = BuildSnapshot(wsWpx, wpxWeekCols)

    Dim lastRow As Long
    lastRow = LastPlayerRow(wsPT)
    If lastRow < 2 Then GoTo SafeExit

    ' A WPX week with no GT column of its own gets one, so every WPX week ends up
    ' in a week column rather than only in the CV block.
    Dim addedCols As Long
    addedCols = EnsureWeekColumns(wsPT, gtWeekCols, wpxWeekCols)
    If addedCols > 0 Then Set gtWeekCols = BuildWeekColumnMap(wsPT)

    Dim unmatched As Collection
    Set unmatched = ListUnmatchedPlayers(wsPT, lastRow)
    Dim mapAdded As Long
    mapAdded = EnsureIDMap(unmatched)

    WriteHelperColumn wsPT, lastRow

    ' GT week columns for weeks GT never played: those are the WPX-owned ones.
    Dim ownedWeeks As Long, missingWeeks As Long
    ownedWeeks = WriteGTWeekFormulas(wsPT, gtWeekCols, snapCols, lastRow)
    missingWeeks = WriteWpxBlock(wsPT, wsWpx, wpxWeekCols, snapCols, lastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.Calculation = xlCalculationManual

SafeExit:
    Dim errNum As Long, errText As String
    errNum = Err.Number
    errText = Err.Description
    On Error Resume Next

    If Not wbWpx Is Nothing Then
        If opened Then wbWpx.Close SaveChanges:=False
    End If

    If calcSaved Then Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If errNum <> 0 Then
        Notify "RefreshWPXTracking error " & errNum & ": " & errText, vbExclamation
    ElseIf Not unmatched Is Nothing Then
        Dim msg As String
        msg = "WPX weeks linked into " & PT_SHEET & " by Player ID." & vbCrLf & _
              "GT week columns filled from WPX: " & ownedWeeks & vbCrLf & _
              "Week columns added for WPX weeks GT had none for: " & addedCols & vbCrLf & _
              "WPX weeks still without a GT column (in the " & WPX_BLOCK_COL & _
              " block only): " & missingWeeks & vbCrLf & _
              "Flagged transfers with no WPX match: " & unmatched.count
        If mapAdded > 0 Then
            msg = msg & vbCrLf & vbCrLf & mapAdded & " player(s) were added to the """ & _
                  MAP_SHEET & """ sheet. Fill in their WPX player ID or name in " & _
                  "column B and run this again."
        End If
        msg = msg & vbCrLf & vbCrLf & "Run DiagnoseWPXTrackingMerge for a per-player reason."
        Notify msg, vbInformation, "Done"
    End If
End Sub

' Kept so the old macro name and any existing button still work.
Public Sub MergeWPXTrackingIntoPlayerTracking()
    RefreshWPXTracking
End Sub

' Entry point for the Workbook_Open handler.
Public Sub RunWPXTrackingOnOpen()
    m_silent = True
    On Error Resume Next
    RefreshWPXTracking
    On Error GoTo 0
    m_silent = False
End Sub

' Strips every WPX column back off Player Tracking, including stale pasted
' values written by an older version of this macro.
Public Sub ClearWPXTracking()
    Dim wsPT As Worksheet
    Set wsPT = SheetOrNothing(ThisWorkbook, PT_SHEET)
    If wsPT Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Clear every WPX column on " & PT_SHEET & "?" & vbCrLf & vbCrLf & _
              "Weeks GT played are left untouched.", vbYesNo + vbQuestion, _
              "Clear WPX data") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    Dim gtWeekCols As Object
    Set gtWeekCols = BuildWeekColumnMap(wsPT)

    Dim k As Variant
    For Each k In gtWeekCols.Keys
        If Not GTPlayedWeek(CStr(gtWeekCols(k)(1))) Then
            ClearColumnBody wsPT, CLng(gtWeekCols(k)(0))
            If gtWeekCols(k)(2) > 0 Then ClearColumnBody wsPT, CLng(gtWeekCols(k)(2))
        End If
    Next k

    Dim firstBlock As Long, lastCol As Long
    firstBlock = wsPT.Range(WPX_BLOCK_COL & "1").Column
    lastCol = wsPT.Cells(1, wsPT.Columns.count).End(xlToLeft).Column
    If lastCol >= firstBlock Then
        wsPT.Range(wsPT.Cells(1, firstBlock), wsPT.Cells(LAST_PT_ROW, lastCol)).ClearContents
    End If
    ClearColumnBody wsPT, wsPT.Range(HELPER_COL & "1").Column

    Application.ScreenUpdating = True
    MsgBox "WPX columns cleared.", vbInformation
End Sub

' Reports, per Player Tracking row, the WPX row it resolved to, how it matched,
' and how many weeks it filled. Writes a sheet and changes no data.
Public Sub DiagnoseWPXTrackingMerge()
    Dim wsPT As Worksheet
    Set wsPT = SheetOrNothing(ThisWorkbook, PT_SHEET)
    If wsPT Is Nothing Then
        MsgBox "Sheet """ & PT_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If

    Dim wsSnap As Worksheet
    Set wsSnap = SheetOrNothing(ThisWorkbook, SNAP_SHEET)
    If wsSnap Is Nothing Then
        MsgBox "No """ & SNAP_SHEET & """ sheet yet - run RefreshWPXTracking first.", _
               vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim snapByID As Object, snapByName As Object
    BuildSnapshotIndex wsSnap, snapByID, snapByName

    Dim idMap As Object
    Set idMap = BuildIDMap()

    Dim gtWeekCols As Object
    Set gtWeekCols = BuildWeekColumnMap(wsPT)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(DIAG_SHEET)
    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("GT Player", "GT Row", "GT Player ID", _
        "WPX Key Used", "Matched By", "Result")
    ws.Rows(1).Font.Bold = True

    Dim lastRow As Long, r As Long, outRow As Long
    lastRow = LastPlayerRow(wsPT)
    outRow = 1

    Dim gtName As String, gtID As String, key As String, snapRow As Long
    Dim matchedBy As String, filled As Long, k As Variant

    For r = 2 To lastRow
        gtName = CleanText(wsPT.Cells(r, 1).Value)
        gtID = NormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        If gtName <> "" Or gtID <> "" Then
            outRow = outRow + 1
            ws.Cells(outRow, 1).Value = gtName
            ws.Cells(outRow, 2).Value = r
            ws.Cells(outRow, 3).Value = gtID

            key = WpxKeyFor(gtID, idMap)
            ws.Cells(outRow, 4).Value = key

            snapRow = SnapshotRow(key, snapByID, snapByName, matchedBy)
            ws.Cells(outRow, 5).Value = matchedBy

            If snapRow = 0 Then
                If gtID = "" Then
                    ws.Cells(outRow, 6).Value = "No Player ID on the GT row"
                Else
                    ws.Cells(outRow, 6).Value = "Not in WPX - add their WPX ID or name to the """ & _
                                                MAP_SHEET & """ sheet"
                End If
            Else
                filled = 0
                For Each k In gtWeekCols.Keys
                    If Not GTPlayedWeek(CStr(gtWeekCols(k)(1))) Then
                        If HasScore(wsPT.Cells(r, gtWeekCols(k)(0)).Value) Then filled = filled + 1
                    End If
                Next k
                ws.Cells(outRow, 6).Value = "Matched - " & filled & " GT week column(s) filled"
            End If
        End If
    Next r

    ws.Columns("A:F").AutoFit
    ws.Activate

    Application.ScreenUpdating = True
    MsgBox "Diagnostics written to the """ & DIAG_SHEET & """ sheet.", vbInformation
End Sub

' Puts a "Pull WPX Weeks" button on the Player Tracking sheet.
Public Sub AddWPXTrackingButton()
    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, PT_SHEET)
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
    shp.OnAction = "RefreshWPXTracking"

    ws.Activate
    MsgBox "Button added to the " & PT_SHEET & " sheet.", vbInformation
End Sub

'==============================================================================
' Snapshot of the WPX Player Tracking sheet
'==============================================================================

' Rebuilds the hidden snapshot: A = WPX player ID, B = WPX player name, then one
' column per WPX week, headed with that week's label. Returns week key ->
' snapshot column letter.
Private Function BuildSnapshot(wsWpx As Worksheet, wpxWeekCols As Object) As Object
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SNAP_SHEET)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("WPX Player ID", "WPX Player")

    ' Weeks in the order the WPX sheet lists them, so the snapshot reads the
    ' same way as the source.
    Dim keys As Collection
    Set keys = WeekKeysByColumn(wpxWeekCols)

    Dim cols As Object
    Set cols = CreateObject("Scripting.Dictionary")

    Dim c As Long, k As Variant
    c = 2
    For Each k In keys
        c = c + 1
        ws.Cells(1, c).Value = wpxWeekCols(k)(1)
        cols(k) = ColumnLetter(c)
    Next k

    Dim idCol As Long
    idCol = FindIDColumn(wsWpx)

    Dim lastRow As Long, r As Long, outRow As Long
    lastRow = LastPlayerRow(wsWpx)
    outRow = 1

    Dim id As String, nameVal As String, v As Variant
    For r = 2 To lastRow
        id = ""
        If idCol > 0 Then id = NormalizeID(wsWpx.Cells(r, idCol).Value)
        nameVal = CleanText(wsWpx.Cells(r, 1).Value)
        If id <> "" Or nameVal <> "" Then
            outRow = outRow + 1
            ws.Cells(outRow, 1).Value = id
            ws.Cells(outRow, 2).Value = nameVal

            c = 2
            For Each k In keys
                c = c + 1
                v = wsWpx.Cells(r, wpxWeekCols(k)(0)).Value
                If HasScore(v) Then ws.Cells(outRow, c).Value = CDbl(v)
            Next k
        End If
    Next r

    AddWeeklySheetScores wsWpx.Parent, ws, keys, wpxWeekCols, cols, outRow

    ws.Rows(1).Font.Bold = True
    ws.Visible = xlSheetHidden

    Set BuildSnapshot = cols
End Function

' Fills the snapshot from the WPX weekly sheets: the per-week sum of Monday to
' Friday, which is exactly what the WPX Player Tracking week column holds.
'
' Player Tracking only carries the players still on the WPX roster, so a player
' who left WPX before transferring is not on it at all and would otherwise come
' back empty - their scores only survive on the weekly sheets. A blank week on a
' player who *is* on Player Tracking is filled from here too, which covers a week
' WPX never rolled up.
Private Sub AddWeeklySheetScores(wbWpx As Workbook, wsSnap As Worksheet, _
    keys As Collection, wpxWeekCols As Object, cols As Object, ByRef outRow As Long)

    Dim byID As Object, byName As Object
    BuildSnapshotIndex wsSnap, byID, byName

    Dim rosterIDs As Object
    Set rosterIDs = BuildWpxNameToID(wbWpx)

    Dim k As Variant, wsWeek As Worksheet, label As String
    For Each k In keys
        label = CStr(wpxWeekCols(k)(1))
        Set wsWeek = SheetOrNothing(wbWpx, label)
        If Not wsWeek Is Nothing Then
            AddOneWeeklySheet wsWeek, wsSnap, CStr(cols(k)), byName, rosterIDs, outRow
        End If
    Next k
End Sub

Private Sub AddOneWeeklySheet(wsWeek As Worksheet, wsSnap As Worksheet, _
    snapCol As String, byName As Object, rosterIDs As Object, ByRef outRow As Long)

    Dim dayCols As Collection, headerRow As Long
    Set dayCols = FindDayColumns(wsWeek, headerRow)
    If dayCols.count = 0 Then Exit Sub

    Dim col As Long
    col = wsSnap.Range(snapCol & "1").Column

    Dim lastRow As Long, r As Long, nameVal As String, key As String
    Dim total As Double, scored As Boolean, dc As Variant, v As Variant
    Dim row As Long

    lastRow = wsWeek.Cells(wsWeek.Rows.count, "A").End(xlUp).row

    For r = headerRow + 1 To lastRow
        nameVal = CleanText(wsWeek.Cells(r, 1).Value)
        key = NameKey(nameVal)
        If key <> "" Then
            total = 0
            scored = False
            For Each dc In dayCols
                v = wsWeek.Cells(r, CLng(dc)).Value
                If HasScore(v) Then
                    total = total + CDbl(v)
                    scored = True
                End If
            Next dc

            If scored Then
                If byName.Exists(key) Then
                    row = byName(key)
                Else
                    outRow = outRow + 1
                    row = outRow
                    wsSnap.Cells(row, 2).Value = nameVal
                    If rosterIDs.Exists(key) Then wsSnap.Cells(row, 1).Value = rosterIDs(key)
                    byName.Add key, row
                End If

                ' Player Tracking wins where it has a figure; this only fills gaps.
                If Not HasScore(wsSnap.Cells(row, col).Value) Then
                    wsSnap.Cells(row, col).Value = total
                End If
            End If
        End If
    Next r
End Sub

' Columns of the Monday..Friday headers on a WPX weekly sheet, plus the row they
' sit on.
Private Function FindDayColumns(ws As Worksheet, ByRef headerRow As Long) As Collection
    Dim found As New Collection
    Dim days As Variant
    days = Array("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY")

    Dim r As Long, c As Long, h As String, d As Variant
    For r = 1 To 5
        Set found = New Collection
        For c = 1 To 30
            h = UCase$(CleanText(ws.Cells(r, c).Value))
            For Each d In days
                If h = d Then found.Add c
            Next d
        Next c
        If found.count >= 3 Then
            headerRow = r
            Set FindDayColumns = found
            Exit Function
        End If
    Next r

    Set FindDayColumns = New Collection
End Function

' WPX player name key -> player ID, from the Roster and Archived Players sheets,
' so a player harvested off a weekly sheet still gets an ID where WPX kept one.
Private Function BuildWpxNameToID(wbWpx As Workbook) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim names As Variant, n As Variant
    names = Array("Roster", "Archived Players")

    Dim ws As Worksheet, idCol As Long, lastRow As Long, r As Long
    Dim key As String, id As String
    For Each n In names
        Set ws = SheetOrNothing(wbWpx, CStr(n))
        If Not ws Is Nothing Then
            idCol = FindIDColumn(ws)
            If idCol > 0 Then
                lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
                For r = 2 To lastRow
                    key = NameKey(ws.Cells(r, 1).Value)
                    id = NormalizeID(ws.Cells(r, idCol).Value)
                    If key <> "" And id <> "" Then
                        If Not map.Exists(key) Then map.Add key, id
                    End If
                Next r
            End If
        End If
    Next n

    Set BuildWpxNameToID = map
End Function

' Snapshot ID -> row and snapshot name key -> row.
Private Sub BuildSnapshotIndex(ws As Worksheet, ByRef byID As Object, ByRef byName As Object)
    Set byID = CreateObject("Scripting.Dictionary")
    Set byName = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long, id As String, key As String
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If ws.Cells(ws.Rows.count, 2).End(xlUp).row > lastRow Then
        lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    End If

    For r = 2 To lastRow
        id = NormalizeID(ws.Cells(r, 1).Value)
        If id <> "" Then
            If Not byID.Exists(id) Then byID.Add id, r
        End If
        key = NameKey(ws.Cells(r, 2).Value)
        If key <> "" Then
            If Not byName.Exists(key) Then byName.Add key, r
        End If
    Next r
End Sub

' The row a WPX key resolves to, by ID first then by player name.
Private Function SnapshotRow(key As String, byID As Object, byName As Object, _
    ByRef matchedBy As String) As Long

    matchedBy = ""
    If key = "" Then Exit Function

    Dim id As String
    id = NormalizeID(key)
    If id <> "" Then
        If byID.Exists(id) Then
            matchedBy = "Player ID"
            SnapshotRow = byID(id)
            Exit Function
        End If
    End If

    Dim nk As String
    nk = NameKey(key)
    If nk <> "" Then
        If byName.Exists(nk) Then
            matchedBy = "WPX player name"
            SnapshotRow = byName(nk)
        End If
    End If
End Function

'==============================================================================
' GT ID -> WPX ID crosswalk
'==============================================================================

' GT ID -> WPX ID or WPX player name, from the crosswalk sheet.
Private Function BuildIDMap() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, MAP_SHEET)
    If ws Is Nothing Then
        Set BuildIDMap = dict
        Exit Function
    End If

    Dim lastRow As Long, r As Long, gtID As String, wpxKey As String
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        gtID = NormalizeID(ws.Cells(r, 1).Value)
        wpxKey = CleanText(ws.Cells(r, 2).Value)
        If gtID <> "" And wpxKey <> "" Then dict(gtID) = wpxKey
    Next r

    Set BuildIDMap = dict
End Function

' The WPX key for a GT row: the crosswalk entry when there is one, else the GT ID.
Private Function WpxKeyFor(gtID As String, idMap As Object) As String
    If gtID = "" Then Exit Function
    If idMap.Exists(gtID) Then
        WpxKeyFor = CStr(idMap(gtID))
    Else
        WpxKeyFor = gtID
    End If
End Function

' Players marked as WPX transfers (a note in the "WPX Transfer" column, or an
' entry already on the crosswalk) that no WPX row can be found for. Each entry is
' Array(gtName, gtID). Players who were simply never in WPX are not listed:
' nothing needs mapping for them.
Private Function ListUnmatchedPlayers(wsPT As Worksheet, lastRow As Long) As Collection
    Dim coll As New Collection

    Dim wsSnap As Worksheet
    Set wsSnap = SheetOrNothing(ThisWorkbook, SNAP_SHEET)
    If wsSnap Is Nothing Then
        Set ListUnmatchedPlayers = coll
        Exit Function
    End If

    Dim byID As Object, byName As Object
    BuildSnapshotIndex wsSnap, byID, byName

    Dim idMap As Object
    Set idMap = BuildIDMap()

    Dim r As Long, gtName As String, gtID As String, matchedBy As String
    For r = 2 To lastRow
        gtName = CleanText(wsPT.Cells(r, 1).Value)
        gtID = NormalizeID(wsPT.Cells(r, PT_ID_COL).Value)
        If gtName <> "" And IsTransfer(wsPT, r, gtID, idMap) Then
            If SnapshotRow(WpxKeyFor(gtID, idMap), byID, byName, matchedBy) = 0 Then
                coll.Add Array(gtName, gtID)
            End If
        End If
    Next r

    Set ListUnmatchedPlayers = coll
End Function

' True when the row is flagged as a WPX transfer or is already on the crosswalk.
Private Function IsTransfer(wsPT As Worksheet, r As Long, gtID As String, _
    idMap As Object) As Boolean

    If CleanText(wsPT.Cells(r, TRANSFER_COL).Value) <> "" Then
        IsTransfer = True
        Exit Function
    End If

    If gtID <> "" Then IsTransfer = idMap.Exists(gtID)
End Function

' Creates the crosswalk sheet when missing and lists every unmatched player on
' it, so the only manual step is typing their WPX ID or name in column B.
' Returns how many rows were added.
Private Function EnsureIDMap(unmatched As Collection) As Long
    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, MAP_SHEET)

    If ws Is Nothing Then
        Set ws = GetOrCreateSheet(MAP_SHEET)
        ws.Range("A1:D1").Value = Array("GT Player ID", "WPX Player ID or Name", _
            "GT Player", "Note")
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:D").ColumnWidth = 22
    End If

    Dim listed As Object
    Set listed = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long, id As String
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        id = NormalizeID(ws.Cells(r, 1).Value)
        If id <> "" Then listed(id) = True
    Next r

    Dim added As Long, e As Variant
    For Each e In unmatched
        id = CStr(e(1))
        If id <> "" Then
            If Not listed.Exists(id) Then
                listed(id) = True
                lastRow = lastRow + 1
                ws.Cells(lastRow, 1).Value = id
                ws.Cells(lastRow, 3).Value = e(0)
                ws.Cells(lastRow, 4).Value = "Not found in WPX - put their WPX player ID or name in column B"
                added = added + 1
            End If
        End If
    Next e

    EnsureIDMap = added
End Function

'==============================================================================
' Formulas on Player Tracking
'==============================================================================

' The hidden helper column: the WPX key for the row, mapped through the
' crosswalk when it has an entry.
Private Sub WriteHelperColumn(wsPT As Worksheet, lastRow As Long)
    Dim col As Long
    col = wsPT.Range(HELPER_COL & "1").Column
    wsPT.Cells(1, col).Value = "WPX Key (auto)"

    Dim lookup As String
    lookup = "VLOOKUP($" & PT_ID_COL & "2,'" & MAP_SHEET & "'!$A:$B,2,FALSE)"

    Dim f As String
    f = "=IF($" & PT_ID_COL & "2=" & Q2 & "," & Q2 & _
        ",IF(IFERROR(" & lookup & ",0)=0,$" & PT_ID_COL & "2," & lookup & "))"

    With wsPT.Range(wsPT.Cells(2, col), wsPT.Cells(lastRow, col))
        .ClearContents
        .Formula = f
    End With
    wsPT.Columns(col).Hidden = True
End Sub

' Lays the WPX-owned week columns out newest to oldest, directly after the GT
' weeks, so the sheet reads in one unbroken run of weeks: a WPX week the GT sheet
' had no column for gets one in date order rather than tacked on at the end.
' Returns how many week pairs were written.
'
' The GT weeks themselves are left exactly where they are. The GT refresh writes
' them positionally - newest at column J, two columns per week, clearing
' C..(9 + weeks*2) - so moving them would put its output in the wrong columns.
Private Function EnsureWeekColumns(wsPT As Worksheet, gtWeekCols As Object, _
    wpxWeekCols As Object) As Long

    Dim idCol As Long
    idCol = wsPT.Range(PT_ID_COL & "1").Column

    Dim firstCol As Long
    firstCol = LastGTWeekColumn(gtWeekCols) + 1
    If firstCol < 10 Then firstCol = 10
    If firstCol + 1 >= idCol Then Exit Function

    ' Every week that is not a GT week: the WPX weeks, plus any week already on
    ' the sheet that GT never played, keyed so a week is only laid out once.
    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In gtWeekCols.Keys
        If Not GTPlayedWeek(CStr(gtWeekCols(k)(1))) Then labels(k) = gtWeekCols(k)(1)
    Next k
    For Each k In wpxWeekCols.Keys
        If Not labels.Exists(k) Then
            If Not GTPlayedWeek(CStr(wpxWeekCols(k)(1))) Then labels(k) = wpxWeekCols(k)(1)
        End If
    Next k

    ' Newest first, matching the direction the GT refresh writes its own weeks.
    Dim ordered As Collection
    Set ordered = KeysNewestFirst(labels)

    ' The old layout goes before the new one is written, so a week that moves
    ' does not leave a copy of itself behind.
    wsPT.Range(wsPT.Cells(1, firstCol), wsPT.Cells(LAST_PT_ROW, idCol - 1)).ClearContents

    Dim c As Long, written As Long
    c = firstCol
    For Each k In ordered
        If c + 1 >= idCol Then Exit For
        wsPT.Cells(1, c).Value = "Weekly Total (" & labels(k) & ")"
        wsPT.Cells(1, c + 1).Value = "Weekly Rank (" & labels(k) & ")"
        c = c + 2
        written = written + 1
    Next k

    EnsureWeekColumns = written
End Function

' The last column the GT weeks occupy: the weeks GT played itself, which the GT
' refresh owns.
Private Function LastGTWeekColumn(gtWeekCols As Object) As Long
    Dim k As Variant, c As Long
    For Each k In gtWeekCols.Keys
        If GTPlayedWeek(CStr(gtWeekCols(k)(1))) Then
            c = gtWeekCols(k)(0)
            If gtWeekCols(k)(2) > c Then c = gtWeekCols(k)(2)
            If c > LastGTWeekColumn Then LastGTWeekColumn = c
        End If
    Next k
End Function

' Week keys sorted newest week first. The keys are month-day, which sorts
' correctly for a season that stays inside one calendar year.
Private Function KeysNewestFirst(weekLabels As Object) As Collection
    Dim ordered As New Collection

    Dim k As Variant, i As Long, inserted As Boolean
    For Each k In weekLabels.Keys
        inserted = False
        For i = 1 To ordered.count
            If CStr(k) > CStr(ordered(i)) Then
                ordered.Add k, , i
                inserted = True
                Exit For
            End If
        Next i
        If Not inserted Then ordered.Add k
    Next k

    Set KeysNewestFirst = ordered
End Function

' Fills the GT weekly columns for weeks GT never played, plus their rank
' columns. Returns how many columns were written.
Private Function WriteGTWeekFormulas(wsPT As Worksheet, gtWeekCols As Object, _
    snapCols As Object, lastRow As Long) As Long

    Dim k As Variant, totalCol As Long, rankCol As Long, written As Long
    For Each k In gtWeekCols.Keys
        If snapCols.Exists(k) Then
            If Not GTPlayedWeek(CStr(gtWeekCols(k)(1))) Then
                totalCol = gtWeekCols(k)(0)
                rankCol = gtWeekCols(k)(2)
                WriteLookupColumn wsPT, totalCol, CStr(snapCols(k)), lastRow
                If rankCol > 0 Then WriteRankColumn wsPT, totalCol, rankCol, lastRow
                written = written + 1
            End If
        End If
    Next k

    WriteGTWeekFormulas = written
End Function

' Rebuilds the "WPX Weekly ..." block from CV: overall total, weeks played,
' weekly average, then a total/rank pair per WPX week. Returns how many WPX
' weeks have no GT column of their own.
Private Function WriteWpxBlock(wsPT As Worksheet, wsWpx As Worksheet, _
    wpxWeekCols As Object, snapCols As Object, lastRow As Long) As Long

    Dim firstCol As Long, lastCol As Long
    firstCol = wsPT.Range(WPX_BLOCK_COL & "1").Column
    lastCol = wsPT.Cells(1, wsPT.Columns.count).End(xlToLeft).Column
    If lastCol < firstCol Then lastCol = firstCol
    wsPT.Range(wsPT.Cells(1, firstCol), wsPT.Cells(LAST_PT_ROW, lastCol)).ClearContents

    Dim gtWeekCols As Object
    Set gtWeekCols = BuildWeekColumnMap(wsPT)

    Dim keys As Collection
    Set keys = WeekKeysByColumn(wpxWeekCols)

    wsPT.Cells(1, firstCol).Value = "WPX Overall Total"
    wsPT.Cells(1, firstCol + 1).Value = "WPX Weeks Played"
    wsPT.Cells(1, firstCol + 2).Value = "WPX Weekly Average"
    wsPT.Cells(1, firstCol + 3).Value = "WPX Overall Rank"

    Dim totalList As String, c As Long, k As Variant, missing As Long
    c = firstCol + 3
    For Each k In keys
        c = c + 1
        wsPT.Cells(1, c).Value = "WPX Weekly Total (" & wpxWeekCols(k)(1) & ")"
        wsPT.Cells(1, c + 1).Value = "WPX Weekly Rank (" & wpxWeekCols(k)(1) & ")"
        WriteLookupColumn wsPT, c, CStr(snapCols(k)), lastRow
        WriteRankColumn wsPT, c, c + 1, lastRow

        If totalList <> "" Then totalList = totalList & ","
        totalList = totalList & ColumnLetter(c) & "2"

        If Not gtWeekCols.Exists(k) Then missing = missing + 1
        c = c + 1
    Next k

    If totalList <> "" Then
        Dim totalLetter As String, countLetter As String
        totalLetter = ColumnLetter(firstCol)
        countLetter = ColumnLetter(firstCol + 1)

        FillFormula wsPT, firstCol, lastRow, _
            "=IF(COUNT(" & totalList & ")=0," & Q2 & ",SUM(" & totalList & "))"
        FillFormula wsPT, firstCol + 1, lastRow, "=COUNT(" & totalList & ")"
        FillFormula wsPT, firstCol + 2, lastRow, _
            "=IF(" & countLetter & "2=0," & Q2 & "," & totalLetter & "2/" & countLetter & "2)"
        WriteRankColumn wsPT, firstCol, firstCol + 3, lastRow
    End If

    wsPT.Rows(1).WrapText = True
    WriteWpxBlock = missing
End Function

' One snapshot column, looked up by the helper key: its ID first, then its
' player name, so a crosswalk entry can name the player instead of their ID.
Private Sub WriteLookupColumn(wsPT As Worksheet, col As Long, snapCol As String, _
    lastRow As Long)

    Dim byID As String, byName As String, v As String
    byID = "INDEX('" & SNAP_SHEET & "'!$" & snapCol & ":$" & snapCol & _
           ",MATCH($" & HELPER_COL & "2,'" & SNAP_SHEET & "'!$A:$A,0))"
    byName = "INDEX('" & SNAP_SHEET & "'!$" & snapCol & ":$" & snapCol & _
             ",MATCH($" & HELPER_COL & "2,'" & SNAP_SHEET & "'!$B:$B,0))"
    v = "IFERROR(" & byID & "," & byName & ")"

    FillFormula wsPT, col, lastRow, _
        "=IF($" & HELPER_COL & "2=" & Q2 & "," & Q2 & _
        ",IFERROR(IF(" & v & "=" & Q2 & "," & Q2 & "," & v & ")," & Q2 & "))"
End Sub

' Highest score = rank 1, over the rows that have a score. Text and blanks are
' ignored by RANK, so the "" a lookup formula leaves behind does not count.
Private Sub WriteRankColumn(wsPT As Worksheet, totalCol As Long, rankCol As Long, _
    lastRow As Long)

    Dim letter As String
    letter = ColumnLetter(totalCol)

    FillFormula wsPT, rankCol, lastRow, _
        "=IF(" & letter & "2=" & Q2 & "," & Q2 & ",RANK(" & letter & "2," & _
        letter & "$2:" & letter & "$" & LAST_PT_ROW & ",0))"
End Sub

Private Sub FillFormula(ws As Worksheet, col As Long, lastRow As Long, f As String)
    With ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
        .ClearContents
        .Formula = f
    End With
End Sub

Private Sub ClearColumnBody(ws As Worksheet, col As Long)
    If col <= 0 Then Exit Sub
    ws.Range(ws.Cells(2, col), ws.Cells(LAST_PT_ROW, col)).ClearContents
End Sub

'==============================================================================
' Week headers
'==============================================================================

' Week start date -> Array(total column, week label, rank column), read from the
' "Weekly Total (...)" headers in row 1. The "WPX Weekly ..." block is skipped
' so the block this macro writes is never mistaken for a GT week.
Private Function BuildWeekColumnMap(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    Dim c As Long, header As String, label As String, key As String
    For c = 1 To lastCol
        header = CleanText(ws.Cells(1, c).Value)
        If InStr(1, header, "WPX", vbTextCompare) <> 1 Then
            label = WeekLabelFromHeader(header)
            If label <> "" Then
                key = WeekKey(label)
                If key <> "" Then
                    If Not dict.Exists(key) Then
                        dict.Add key, Array(c, label, RankColumnFor(ws, c, label, lastCol))
                    End If
                End If
            End If
        End If
    Next c

    Set BuildWeekColumnMap = dict
End Function

' Week keys in the order their columns appear on the sheet.
Private Function WeekKeysByColumn(weekCols As Object) As Collection
    Dim keys As New Collection

    Dim k As Variant, i As Long, inserted As Boolean
    For Each k In weekCols.Keys
        inserted = False
        For i = 1 To keys.count
            If weekCols(k)(0) < weekCols(keys(i))(0) Then
                keys.Add k, , i
                inserted = True
                Exit For
            End If
        Next i
        If Not inserted Then keys.Add k
    Next k

    Set WeekKeysByColumn = keys
End Function

' True when GT played the week itself, i.e. the workbook has a weekly sheet of
' that name. Those columns belong to the GT refresh and are never touched.
Private Function GTPlayedWeek(label As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(WeekKey(ws.name), WeekKey(label), vbTextCompare) = 0 Then
            If WeekKey(ws.name) <> "" And InStr(ws.name, " - ") > 0 Then
                GTPlayedWeek = True
                Exit Function
            End If
        End If
    Next ws
End Function

' The rank column that goes with a weekly total column: the next column when it
' is that week's "Weekly Rank (...)" header, otherwise none (0).
Private Function RankColumnFor(ws As Worksheet, totalCol As Long, _
    label As String, lastCol As Long) As Long

    If totalCol >= lastCol Then Exit Function

    Dim header As String
    header = CleanText(ws.Cells(1, totalCol + 1).Value)
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

'==============================================================================
' Small helpers
'==============================================================================

' Two double-quote characters, i.e. an empty string inside a formula.
Private Function Q2() As String
    Q2 = Chr(34) & Chr(34)
End Function

Private Function ColumnLetter(ByVal col As Long) As String
    Dim n As Long
    n = col
    Do While n > 0
        ColumnLetter = Chr(65 + ((n - 1) Mod 26)) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function LastPlayerRow(ws As Worksheet) As Long
    LastPlayerRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
End Function

Private Function SheetOrNothing(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = SheetOrNothing(ThisWorkbook, sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.name = sheetName
    End If
    ws.Visible = xlSheetVisible
    Set GetOrCreateSheet = ws
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
        header = LCase$(CleanText(ws.Cells(1, c).Value))
        If header = "player id" Or header = "id" Then
            FindIDColumn = c
            Exit Function
        End If
    Next c

    FindIDColumn = clCol
End Function

' True when the column holds an ID somewhere, so a blank leading row does not
' make the column look empty.
Private Function ColumnHasID(ws As Worksheet, col As Long) As Boolean
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).row
    For r = 2 To lastRow
        If NormalizeID(ws.Cells(r, col).Value) <> "" Then
            ColumnHasID = True
            Exit Function
        End If
    Next r
End Function

Private Function HasScore(ByVal v As Variant) As Boolean
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If CStr(v) = "" Then Exit Function
    HasScore = True
End Function

' Trims a cell to plain text, dropping the non-breaking and zero-width spaces
' that come along when names are pasted in from the game.
Private Function CleanText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, ChrW(8203), "")
    s = Replace(s, ChrW(65279), "")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    CleanText = Trim$(s)
End Function

' Case- and punctuation-insensitive key for player names.
Private Function NameKey(ByVal v As Variant) As String
    Dim s As String, out As String, i As Long, ch As String
    s = UCase$(CleanText(v))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    NameKey = out
End Function

' Roster IDs arrive as numbers, as text with commas or spaces, and sometimes in
' scientific notation, so they are reduced to a digits-only key. They are kept
' as text so IDs longer than a Long cannot overflow. A non-numeric value is not
' an ID and comes back empty, so it can be treated as a player name instead.
Private Function NormalizeID(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function

    Dim txt As String
    txt = CleanText(v)
    txt = Replace(txt, " ", "")
    txt = Replace(txt, ",", "")
    If txt = "" Then Exit Function

    If IsNumeric(txt) Then
        If InStr(1, txt, "E", vbTextCompare) > 0 Then txt = Format$(CDbl(txt), "0")
    Else
        Exit Function
    End If

    Dim digits As String, i As Long, ch As String
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    Do While Len(digits) > 1 And Left$(digits, 1) = "0"
        digits = Mid$(digits, 2)
    Loop

    If digits = "0" Then Exit Function
    NormalizeID = digits
End Function

'==============================================================================
' WPX workbook, and the optional Workbook_Open hook
'==============================================================================

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

' MsgBox, unless we are running from Workbook_Open.
Private Sub Notify(ByVal msg As String, ByVal style As VbMsgBoxStyle, _
    Optional ByVal title As String = "WPX Merge")

    If m_silent Then Exit Sub
    MsgBox msg, style, title
End Sub

' Writes the Workbook_Open handler into ThisWorkbook so the refresh runs on every
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
    MsgBox "WPX auto-run installed - the refresh now runs when " & ThisWorkbook.name & _
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
