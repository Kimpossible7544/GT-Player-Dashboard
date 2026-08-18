Attribute VB_Name = "WPXArenaBlock"
' ============================================================
' WPXArenaBlock - VBA Macro Module
' Rebuilds the WPX Arena Power comparison block on the GT Arena Power sheet.
'
' The block starts at column 220 so the GT-start growth columns and weekly
' history remain unchanged. Matching uses the GT name first, then the roster's
' current-name/AKA resolution through LookupRosterName.
'
' Requires the ArenaPowerMacros module (LookupRosterName, NormalizeArenaName).
'
' USAGE:
'   1. Run ImportWPXData in WPXDashboard.bas first.
'   2. Run RefreshWPXArenaBlock after the GT Arena Power sheet is refreshed.
' ============================================================

Option Explicit

Private Const WPXA_ARENA_SHEET As String = "Arena Power"
Private Const WPXA_ROSTER_SHEET As String = "Roster"
Private Const WPXA_WPX_SHEET As String = "WPX_ArenaPower"
Private Const WPXA_START_COL As Long = 220
Private Const WPXA_HEADER_ROW As Long = 1
Private Const WPXA_HISTORY_START_COL As Long = 11
Private Const WPXA_HISTORY_GROUP_SIZE As Long = 4
Private Const WPXA_FIXED_COLS As Long = 15
Private Const WPXA_MAX_COL As Long = 16384

Public Sub RefreshWPXArenaBlock()

    Dim wsArena As Worksheet, wsRoster As Worksheet, wsWpx As Worksheet
    On Error Resume Next
    Set wsArena = ThisWorkbook.Worksheets(WPXA_ARENA_SHEET)
    Set wsRoster = ThisWorkbook.Worksheets(WPXA_ROSTER_SHEET)
    Set wsWpx = ThisWorkbook.Worksheets(WPXA_WPX_SHEET)
    On Error GoTo 0

    If wsArena Is Nothing Then
        MsgBox "Sheet """ & WPXA_ARENA_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If
    If wsRoster Is Nothing Then
        MsgBox "Sheet """ & WPXA_ROSTER_SHEET & """ not found.", vbExclamation
        Exit Sub
    End If
    If wsWpx Is Nothing Then
        MsgBox "Sheet """ & WPXA_WPX_SHEET & """ not found. Run ImportWPXData first.", _
               vbExclamation
        Exit Sub
    End If

    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo WPXA_Failed
    Application.ScreenUpdating = False

    Dim gtLastRow As Long, wpxLastCol As Long
    gtLastRow = wsArena.Cells(wsArena.Rows.Count, 1).End(xlUp).Row
    wpxLastCol = wsWpx.Cells(WPXA_HEADER_ROW, wsWpx.Columns.Count).End(xlToLeft).Column

    WPXA_ClearBlock wsArena, gtLastRow
    WPXA_WriteHeaders wsArena, wpxLastCol

    Dim gtRow As Long, wpxRow As Long
    Dim gtName As String, wpxName As String
    Dim startDate As Variant, startArena As Variant, startHQ As Variant
    Dim matched As Long, missing As Long

    For gtRow = 2 To gtLastRow
        gtName = Trim$(CStr(wsArena.Cells(gtRow, 1).Value))
        If gtName <> "" Then
            wpxRow = WPXA_FindWpxRow(wsWpx, wsRoster, gtName)
            If wpxRow = 0 Then
                missing = missing + 1
            Else
                wpxName = Trim$(CStr(wsWpx.Cells(wpxRow, 1).Value))
                WPXA_FindBaseline wsWpx, wpxRow, startDate, startArena, startHQ
                WPXA_WritePlayerBlock wsArena, wsWpx, gtRow, wpxRow, wpxName, _
                    startDate, startArena, startHQ
                matched = matched + 1
            End If
        End If
    Next gtRow

    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "WPX Arena block refreshed." & vbCrLf & _
           "Players written: " & matched & vbCrLf & _
           "GT rows with no WPX match: " & missing, vbInformation
    Exit Sub

WPXA_Failed:
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "WPX Arena block refresh failed: " & Err.Description, vbExclamation
End Sub

Private Sub WPXA_ClearBlock(wsArena As Worksheet, gtLastRow As Long)
    Dim clearLastCol As Long, clearLastRow As Long
    clearLastCol = wsArena.Cells(WPXA_HEADER_ROW, wsArena.Columns.Count).End(xlToLeft).Column
    If clearLastCol < WPXA_START_COL Then clearLastCol = WPXA_START_COL

    clearLastRow = gtLastRow
    Dim oldBlockLastRow As Long, c As Long
    For c = WPXA_START_COL To clearLastCol
        oldBlockLastRow = wsArena.Cells(wsArena.Rows.Count, c).End(xlUp).Row
        If oldBlockLastRow > clearLastRow Then clearLastRow = oldBlockLastRow
    Next c

    wsArena.Range(wsArena.Cells(WPXA_HEADER_ROW, WPXA_START_COL), _
                  wsArena.Cells(clearLastRow, clearLastCol)).ClearContents
End Sub

Private Sub WPXA_WriteHeaders(wsArena As Worksheet, wpxLastCol As Long)
    Dim fixedHeaders As Variant
    fixedHeaders = Array("WPX Player", "Lifetime Arena Chg", "Lifetime HQ Chg", _
                         "WPX Start Date", "WPX Start Arena", "WPX Start HQ", _
                         "WPX Current Date", "WPX Current Level", _
                         "WPX Current Arena", "WPX Current HQ", _
                         "WPX Arena Chg", "WPX HQ Chg", _
                         "WPX Overall Arena Chg", "WPX Overall HQ Chg", _
                         "WPX Level Note")

    Dim i As Long, groupStart As Long, groupCount As Long
    For i = LBound(fixedHeaders) To UBound(fixedHeaders)
        wsArena.Cells(WPXA_HEADER_ROW, WPXA_START_COL + i).Value = fixedHeaders(i)
    Next i

    If wpxLastCol < WPXA_HISTORY_START_COL Then Exit Sub
    groupCount = ((wpxLastCol - WPXA_HISTORY_START_COL) \ WPXA_HISTORY_GROUP_SIZE) + 1
    For i = 0 To groupCount - 1
        groupStart = WPXA_START_COL + WPXA_FIXED_COLS + (i * WPXA_HISTORY_GROUP_SIZE)
        If groupStart + WPXA_HISTORY_GROUP_SIZE - 1 > WPXA_MAX_COL Then Exit For
        wsArena.Cells(WPXA_HEADER_ROW, groupStart).Value = "Date"
        wsArena.Cells(WPXA_HEADER_ROW, groupStart + 1).Value = "Level"
        wsArena.Cells(WPXA_HEADER_ROW, groupStart + 2).Value = "Arena Power"
        wsArena.Cells(WPXA_HEADER_ROW, groupStart + 3).Value = "HQ Power"
    Next i
End Sub

Private Function WPXA_FindWpxRow(wsWpx As Worksheet, wsRoster As Worksheet, _
    gtName As String) As Long

    Dim targetKey As String, sourceName As String, rosterName As String
    Dim lastRow As Long, r As Long
    targetKey = NormalizeArenaName(gtName)
    lastRow = wsWpx.Cells(wsWpx.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        sourceName = Trim$(CStr(wsWpx.Cells(r, 1).Value))
        If sourceName <> "" Then
            If NormalizeArenaName(sourceName) = targetKey Then
                WPXA_FindWpxRow = r
                Exit Function
            End If
        End If
    Next r

    For r = 2 To lastRow
        sourceName = Trim$(CStr(wsWpx.Cells(r, 1).Value))
        If sourceName <> "" Then
            rosterName = LookupRosterName(wsRoster, sourceName)
            If rosterName <> "" Then
                If NormalizeArenaName(rosterName) = targetKey Then
                    WPXA_FindWpxRow = r
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Sub WPXA_FindBaseline(wsWpx As Worksheet, wpxRow As Long, _
    ByRef startDate As Variant, ByRef startArena As Variant, ByRef startHQ As Variant)

    Dim lastCol As Long, c As Long
    startDate = Empty
    startArena = Empty
    startHQ = Empty
    lastCol = wsWpx.Cells(wpxRow, wsWpx.Columns.Count).End(xlToLeft).Column

    If lastCol < WPXA_HISTORY_START_COL Then Exit Sub
    For c = WPXA_HISTORY_START_COL To lastCol Step WPXA_HISTORY_GROUP_SIZE
        If c + WPXA_HISTORY_GROUP_SIZE - 1 > wsWpx.Columns.Count Then Exit For
        If WPXA_HistoryGroupHasData(wsWpx, wpxRow, c) Then
            startDate = wsWpx.Cells(wpxRow, c).Value
            startArena = WPXA_ParsePower(wsWpx.Cells(wpxRow, c + 2).Value)
            startHQ = WPXA_ParsePower(wsWpx.Cells(wpxRow, c + 3).Value)
        End If
    Next c
End Sub

Private Function WPXA_HistoryGroupHasData(wsWpx As Worksheet, wpxRow As Long, _
    groupStart As Long) As Boolean
    If groupStart + WPXA_HISTORY_GROUP_SIZE - 1 > wsWpx.Columns.Count Then Exit Function

    Dim dateValue As Variant, levelValue As Variant
    dateValue = wsWpx.Cells(wpxRow, groupStart).Value
    levelValue = wsWpx.Cells(wpxRow, groupStart + 1).Value
    WPXA_HistoryGroupHasData = WPXA_ValuePresent(dateValue) Or _
                               WPXA_ValuePresent(levelValue) Or _
                               IsNumeric(WPXA_ParsePower(wsWpx.Cells(wpxRow, groupStart + 2).Value)) Or _
                               IsNumeric(WPXA_ParsePower(wsWpx.Cells(wpxRow, groupStart + 3).Value))
End Function

Private Sub WPXA_WritePlayerBlock(wsArena As Worksheet, wsWpx As Worksheet, _
    gtRow As Long, wpxRow As Long, wpxName As String, _
    startDate As Variant, startArena As Variant, startHQ As Variant)

    Dim gtArena As Variant, gtHQ As Variant
    gtArena = WPXA_ParsePower(wsArena.Cells(gtRow, 4).Value)
    gtHQ = WPXA_ParsePower(wsArena.Cells(gtRow, 5).Value)

    wsArena.Cells(gtRow, WPXA_START_COL).Value = wpxName
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 1), _
        WPXA_SafeDelta(gtArena, startArena)
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 2), _
        WPXA_SafeDelta(gtHQ, startHQ)
    WPXA_WriteDateOrDash wsArena.Cells(gtRow, WPXA_START_COL + 3), startDate
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 4), startArena
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 5), startHQ

    WPXA_WriteDateOrDash wsArena.Cells(gtRow, WPXA_START_COL + 6), _
        wsWpx.Cells(wpxRow, 2).Value
    WPXA_WriteValueOrDash wsArena.Cells(gtRow, WPXA_START_COL + 7), _
        wsWpx.Cells(wpxRow, 3).Value
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 8), _
        wsWpx.Cells(wpxRow, 4).Value
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 9), _
        wsWpx.Cells(wpxRow, 5).Value
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 10), _
        wsWpx.Cells(wpxRow, 6).Value
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 11), _
        wsWpx.Cells(wpxRow, 7).Value
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 12), _
        wsWpx.Cells(wpxRow, 8).Value
    WPXA_WritePower wsArena.Cells(gtRow, WPXA_START_COL + 13), _
        wsWpx.Cells(wpxRow, 9).Value
    WPXA_WriteValueOrDash wsArena.Cells(gtRow, WPXA_START_COL + 14), _
        wsWpx.Cells(wpxRow, 10).Value

    WPXA_CopyHistory wsArena, wsWpx, gtRow, wpxRow
End Sub

Private Sub WPXA_CopyHistory(wsArena As Worksheet, wsWpx As Worksheet, _
    gtRow As Long, wpxRow As Long)

    Dim sourceLastCol As Long, oldestStart As Long, c As Long
    sourceLastCol = wsWpx.Cells(wpxRow, wsWpx.Columns.Count).End(xlToLeft).Column
    If sourceLastCol < WPXA_HISTORY_START_COL Then Exit Sub

    For c = WPXA_HISTORY_START_COL To sourceLastCol Step WPXA_HISTORY_GROUP_SIZE
        If c + WPXA_HISTORY_GROUP_SIZE - 1 > wsWpx.Columns.Count Then Exit For
        If WPXA_HistoryGroupHasData(wsWpx, wpxRow, c) Then oldestStart = c
    Next c
    If oldestStart = 0 Then Exit Sub

    Dim targetCol As Long
    targetCol = WPXA_START_COL + WPXA_FIXED_COLS
    For c = oldestStart To WPXA_HISTORY_START_COL Step -WPXA_HISTORY_GROUP_SIZE
        If targetCol + WPXA_HISTORY_GROUP_SIZE - 1 > WPXA_MAX_COL Then Exit For
        WPXA_WriteHistoryGroup wsArena, wsWpx, gtRow, wpxRow, targetCol, c
        targetCol = targetCol + WPXA_HISTORY_GROUP_SIZE
    Next c
End Sub

Private Sub WPXA_WriteHistoryGroup(wsArena As Worksheet, wsWpx As Worksheet, _
    gtRow As Long, wpxRow As Long, targetCol As Long, sourceCol As Long)

    WPXA_WriteDateOrDash wsArena.Cells(gtRow, targetCol), _
        wsWpx.Cells(wpxRow, sourceCol).Value
    WPXA_WriteValueOrDash wsArena.Cells(gtRow, targetCol + 1), _
        wsWpx.Cells(wpxRow, sourceCol + 1).Value
    WPXA_WritePower wsArena.Cells(gtRow, targetCol + 2), _
        wsWpx.Cells(wpxRow, sourceCol + 2).Value
    WPXA_WritePower wsArena.Cells(gtRow, targetCol + 3), _
        wsWpx.Cells(wpxRow, sourceCol + 3).Value
End Sub

Private Sub WPXA_WritePower(targetCell As Range, value As Variant)
    targetCell.Value = WPXA_FormatPower(WPXA_ParsePower(value))
End Sub

Private Sub WPXA_WriteValueOrDash(targetCell As Range, value As Variant)
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        targetCell.Value = "-"
    ElseIf Trim$(CStr(value)) = "" Or Trim$(CStr(value)) = "-" Then
        targetCell.Value = "-"
    Else
        targetCell.Value = value
    End If
End Sub

Private Sub WPXA_WriteDateOrDash(targetCell As Range, value As Variant)
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        targetCell.Value = "-"
    ElseIf Trim$(CStr(value)) = "" Or Trim$(CStr(value)) = "-" Then
        targetCell.Value = "-"
    ElseIf IsDate(value) Then
        targetCell.Value = CDate(value)
    Else
        targetCell.Value = value
    End If
End Sub

Private Function WPXA_ValuePresent(value As Variant) As Boolean
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    WPXA_ValuePresent = Trim$(CStr(value)) <> "" And Trim$(CStr(value)) <> "-"
End Function

Private Function WPXA_ParsePower(value As Variant) As Variant
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        WPXA_ParsePower = Empty
        Exit Function
    End If
    If Trim$(CStr(value)) = "" Or Trim$(CStr(value)) = "-" Then
        WPXA_ParsePower = Empty
        Exit Function
    End If
    If IsNumeric(value) Then
        WPXA_ParsePower = CDbl(value)
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(CStr(value))
    textValue = Replace(textValue, "Mil", "", , , vbTextCompare)
    textValue = Replace(textValue, "M", "", , , vbTextCompare)
    textValue = Replace(textValue, ",", "")
    textValue = Replace(textValue, " ", "")
    If IsNumeric(textValue) Then WPXA_ParsePower = CDbl(textValue)
End Function

Private Function WPXA_FormatPower(value As Variant) As String
    If IsNumeric(value) Then
        WPXA_FormatPower = Format$(CDbl(value), "0.00") & " M"
    Else
        WPXA_FormatPower = "-"
    End If
End Function

Private Function WPXA_SafeDelta(currentValue As Variant, baselineValue As Variant) As Variant
    If IsNumeric(currentValue) And IsNumeric(baselineValue) Then
        WPXA_SafeDelta = CDbl(currentValue) - CDbl(baselineValue)
    Else
        WPXA_SafeDelta = Empty
    End If
End Function
