Attribute VB_Name = "WPXPlayerTrackingPull"
' WPXPlayerTrackingPull.bas
' Pulls Player Tracking data out of the legacy WPX workbook and into
' GTStatsFINAL.xlsm for players whose roster ID exists in BOTH rosters.
'
' For each matched player it copies every Player Tracking column that has
' data - summary columns plus each "Weekly Total (...)" / rank pair for the
' weeks the player actually scored in WPX. Weeks before they joined WPX or
' after they left are left blank instead of being written as zeros.
'
' The data lands on a dedicated "WPX Player Tracking" sheet (created if it
' does not exist, fully rebuilt on each run). It is NOT merged into the GT
' "Player Tracking" sheet on purpose: the WPX week labels are older weeks and
' gt_data.js would read them as GT weeks and skew ranks, averages and streaks.
'
' To use:
'   1. Open GTStatsFINAL.xlsm.
'   2. Alt+F11 > File > Import File > WPXPlayerTrackingPull.bas
'   3. Run PullWPXPlayerTracking() from the macro list.
'   4. Pick WPXStatsFinal.xlsm if prompted (auto-found if it sits in the
'      same folder, or if WPX_PlayerTracking / WPX_Roster sheets are
'      already embedded by WPXDashboard.ImportWPXData).
'
' To limit the pull to certain players, list them in PLAYER_FILTER below,
' by GT name or by roster ID, comma separated. Leave it empty for every
' cross-team player.
'   Private Const PLAYER_FILTER As String = "Kimpossible7544, 12345678"

Option Explicit

Private Const PLAYER_FILTER   As String = ""
Private Const DEST_SHEET      As String = "WPX Player Tracking"
Private Const WPX_PT_SHEET    As String = "Player Tracking"
Private Const WPX_ROSTER      As String = "Roster"
Private Const GT_ROSTER       As String = "Roster"
Private Const WPX_FILE_NAME   As String = "WPXStatsFinal.xlsm"

Public Sub PullWPXPlayerTracking()
    Dim wsGTRoster As Worksheet
    Dim wsWpxRoster As Worksheet, wsWpxPT As Worksheet
    Dim wbWpx As Workbook
    Dim wpxWasOpened As Boolean

    On Error Resume Next
    Set wsGTRoster = ThisWorkbook.Sheets(GT_ROSTER)
    On Error GoTo 0
    If wsGTRoster Is Nothing Then
        MsgBox "Sheet """ & GT_ROSTER & """ not found in this workbook.", vbExclamation
        Exit Sub
    End If

    If Not GetWpxSheets(wsWpxRoster, wsWpxPT, wbWpx, wpxWasOpened) Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim matches As Collection
    Set matches = New Collection
    Call BuildMatchedPlayers(wsGTRoster, wsWpxRoster, matches)

    Dim wsDest As Worksheet
    Set wsDest = GetOrCreateSheet(DEST_SHEET)
    wsDest.Cells.Clear

    Dim lastCol As Long
    lastCol = wsWpxPT.Cells(1, wsWpxPT.Columns.Count).End(xlToLeft).Column

    ' Header: GT identity columns, then the WPX Player Tracking header row as-is.
    wsDest.Cells(1, 1).Value = "GT Player"
    wsDest.Cells(1, 2).Value = "Player ID"
    wsDest.Cells(1, 3).Value = "WPX Player"
    Dim c As Long
    For c = 1 To lastCol
        wsDest.Cells(1, c + 3).Value = wsWpxPT.Cells(1, c).Value
    Next c
    wsDest.Rows(1).Font.Bold = True

    Dim i As Long, p As Variant
    Dim destRow As Long, written As Long, missing As Long
    destRow = 1
    For i = 1 To matches.Count
        p = matches(i)
        Dim wpxRow As Long
        wpxRow = FindNameRow(wsWpxPT, CStr(p(1)))
        If wpxRow = 0 Then
            missing = missing + 1
        Else
            destRow = destRow + 1
            wsDest.Cells(destRow, 1).Value = p(0)
            wsDest.Cells(destRow, 2).Value = p(2)
            wsDest.Cells(destRow, 3).Value = p(1)
            Call CopyTrackingRow(wsWpxPT, wpxRow, lastCol, wsDest, destRow)
            written = written + 1
        End If
    Next i

    wsDest.Columns(1).ColumnWidth = 22
    wsDest.Columns(2).ColumnWidth = 12
    wsDest.Columns(3).ColumnWidth = 22
    wsDest.Rows(1).WrapText = True
    wsDest.Activate
    ActiveWindow.FreezePanes = False
    wsDest.Range("D2").Select
    ActiveWindow.FreezePanes = True

    If wpxWasOpened Then wbWpx.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "WPX Player Tracking pulled for " & written & " player(s)." & vbCrLf & _
           "Cross-team IDs matched: " & matches.Count & vbCrLf & _
           "Matched IDs with no WPX Player Tracking row: " & missing, _
           vbInformation, "Done"
End Sub

' Copies one WPX Player Tracking row, skipping cells with no real data so
' weeks outside the player's WPX tenure stay blank rather than zero.
Private Sub CopyTrackingRow(wsSrc As Worksheet, rSrc As Long, lastCol As Long, _
    wsDst As Worksheet, rDst As Long)

    Dim c As Long, v As Variant
    For c = 1 To lastCol
        v = wsSrc.Cells(rSrc, c).Value
        If HasData(v) Then
            wsDst.Cells(rDst, c + 3).Value = v
            wsDst.Cells(rDst, c + 3).NumberFormat = wsSrc.Cells(rSrc, c).NumberFormat
        End If
    Next c
End Sub

Private Function HasData(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    Dim s As String
    s = Trim(CStr(v))
    If s = "" Or s = "-" Then Exit Function
    HasData = True
End Function

' Builds Array(gtName, wpxName, id) for every ID present in both rosters
' and allowed by PLAYER_FILTER.
Private Sub BuildMatchedPlayers(wsGTRoster As Worksheet, wsWpxRoster As Worksheet, coll As Collection)
    Dim wpxMap As Object
    Set wpxMap = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long
    lastRow = wsWpxRoster.Cells(wsWpxRoster.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 1, 2)
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 5, 6)
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 9, 10)
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 13, 14)
    Next r

    lastRow = wsGTRoster.Cells(wsGTRoster.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        Call CheckAndAdd(coll, wsGTRoster, wpxMap, r, 1, 2)
        Call CheckAndAdd(coll, wsGTRoster, wpxMap, r, 5, 6)
        Call CheckAndAdd(coll, wsGTRoster, wpxMap, r, 9, 10)
        Call CheckAndAdd(coll, wsGTRoster, wpxMap, r, 13, 14)
    Next r
End Sub

Private Sub AddToIdMap(map As Object, ws As Worksheet, r As Long, idCol As Long, nameCol As Long)
    Dim idVal As Variant, nameVal As Variant
    idVal = ws.Cells(r, idCol).Value
    nameVal = ws.Cells(r, nameCol).Value
    If IsNumeric(idVal) And Not IsEmpty(nameVal) Then
        If Trim(CStr(nameVal)) <> "" Then map(CLng(idVal)) = Trim(CStr(nameVal))
    End If
End Sub

Private Sub CheckAndAdd(coll As Collection, ws As Worksheet, map As Object, r As Long, idCol As Long, nameCol As Long)
    Dim idVal As Variant, id As Long
    idVal = ws.Cells(r, idCol).Value
    If Not IsNumeric(idVal) Then Exit Sub

    id = CLng(idVal)
    If Not map.Exists(id) Then Exit Sub

    Dim gtName As String, wpxName As String
    gtName = Trim(CStr(ws.Cells(r, nameCol).Value))
    wpxName = Trim(CStr(map(id)))
    If gtName = "" Or wpxName = "" Then Exit Sub
    If Not IsWanted(gtName, id) Then Exit Sub

    coll.Add Array(gtName, wpxName, id)
End Sub

' True when PLAYER_FILTER is empty, or the player's GT name or ID is listed in it.
Private Function IsWanted(gtName As String, id As Long) As Boolean
    Dim filter As String
    filter = Trim(PLAYER_FILTER)
    If filter = "" Then
        IsWanted = True
        Exit Function
    End If

    Dim parts() As String, i As Long, token As String
    parts = Split(filter, ",")
    For i = LBound(parts) To UBound(parts)
        token = Trim(parts(i))
        If token <> "" Then
            If StrComp(token, gtName, vbTextCompare) = 0 Then
                IsWanted = True
                Exit Function
            End If
            If IsNumeric(token) Then
                If CLng(token) = id Then
                    IsWanted = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' Resolves the WPX Roster and Player Tracking sheets, preferring the hidden
' WPX_* copies embedded by WPXDashboard.ImportWPXData, then WPXStatsFinal.xlsm
' next to this workbook, then a file picker.
Private Function GetWpxSheets(ByRef wsWpxRoster As Worksheet, ByRef wsWpxPT As Worksheet, _
    ByRef wbWpx As Workbook, ByRef wpxWasOpened As Boolean) As Boolean

    On Error Resume Next
    Set wsWpxRoster = ThisWorkbook.Sheets("WPX_Roster")
    Set wsWpxPT = ThisWorkbook.Sheets("WPX_PlayerTracking")
    On Error GoTo 0
    If Not wsWpxRoster Is Nothing And Not wsWpxPT Is Nothing Then
        GetWpxSheets = True
        Exit Function
    End If

    Dim f As Variant, samePath As String
    samePath = ThisWorkbook.Path & Application.PathSeparator & WPX_FILE_NAME
    If Dir(samePath) <> "" Then
        f = samePath
    Else
        f = Application.GetOpenFilename( _
            "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , _
            "Select " & WPX_FILE_NAME)
        If VarType(f) = vbBoolean Then Exit Function
        If CStr(f) = "False" Then Exit Function
    End If

    Set wbWpx = Workbooks.Open(CStr(f), ReadOnly:=True)
    wpxWasOpened = True

    On Error Resume Next
    Set wsWpxRoster = wbWpx.Sheets(WPX_ROSTER)
    Set wsWpxPT = wbWpx.Sheets(WPX_PT_SHEET)
    On Error GoTo 0

    If wsWpxRoster Is Nothing Or wsWpxPT Is Nothing Then
        MsgBox "The selected workbook is missing a """ & WPX_ROSTER & """ or """ & _
               WPX_PT_SHEET & """ sheet.", vbExclamation
        wbWpx.Close SaveChanges:=False
        Exit Function
    End If

    GetWpxSheets = True
End Function

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function FindNameRow(ws As Worksheet, name As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(Trim(CStr(ws.Cells(r, 1).Value)), name, vbTextCompare) = 0 Then
            FindNameRow = r
            Exit Function
        End If
    Next r
End Function
