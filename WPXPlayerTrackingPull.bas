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
'
' Set MATCH_BY_NAME to True to also pull players whose roster ID is missing or
' mistyped in one of the workbooks, by matching on player name instead.
'
' If players are missing from the result, run DiagnoseWPXPlayerTracking().
' It builds a "WPX Pull Diagnostics" sheet listing every GT roster entry with
' the reason it was or was not pulled (no ID, ID not in WPX, no WPX Player
' Tracking row, excluded by the filter, pulled).

Option Explicit

Private Const PLAYER_FILTER   As String = ""
Private Const MATCH_BY_NAME   As Boolean = False  ' also match players whose ID is missing/mistyped, by name
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
    Dim missingNames As String
    destRow = 1
    For i = 1 To matches.Count
        p = matches(i)
        Dim wpxRow As Long
        wpxRow = FindPlayerRow(wsWpxPT, CStr(p(1)), CStr(p(0)))
        If wpxRow = 0 Then
            missing = missing + 1
            If missing <= 15 Then missingNames = missingNames & vbCrLf & "   " & p(1)
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
           "Matched IDs with no WPX Player Tracking row: " & missing & missingNames & vbCrLf & vbCrLf & _
           "Players missing? Run DiagnoseWPXPlayerTracking for a per-player reason.", _
           vbInformation, "Done"
End Sub

' Reports, for every GT roster entry, whether it was pulled and why not.
' Writes the "WPX Pull Diagnostics" sheet; leaves the data sheet untouched.
Public Sub DiagnoseWPXPlayerTracking()
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

    Dim wpxMap As Object
    Set wpxMap = BuildWpxIdMap(wsWpxRoster)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("WPX Pull Diagnostics")
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("GT Player", "GT Roster ID", "WPX Player (by ID)", "WPX Tracking Row", "Result")
    ws.Rows(1).Font.Bold = True

    Dim pairs As Collection
    Set pairs = RosterColumnPairs(wsGTRoster)

    Dim outRow As Long, lastRow As Long, r As Long, pr As Variant
    outRow = 1
    lastRow = RosterLastRow(wsGTRoster, pairs)

    For r = 2 To lastRow
        For Each pr In pairs
            Dim gtName As String, idText As String, id As String
            gtName = CleanText(wsGTRoster.Cells(r, pr(0) + 1).Value)
            idText = CleanText(wsGTRoster.Cells(r, pr(0)).Value)
            If gtName <> "" Or idText <> "" Then
                outRow = outRow + 1
                ws.Cells(outRow, 1).Value = gtName
                ws.Cells(outRow, 2).Value = idText

                id = ParseId(idText)
                If gtName = "" Then
                    ws.Cells(outRow, 5).Value = "Skipped - roster ID with no player name"
                ElseIf id = "" Then
                    ws.Cells(outRow, 5).Value = "Skipped - no usable roster ID on the GT Roster"
                ElseIf Not wpxMap.Exists(id) And Not MATCH_BY_NAME Then
                    ws.Cells(outRow, 5).Value = "Skipped - ID not found on the WPX Roster"
                Else
                    Dim wpxName As String, wpxRow As Long
                    If wpxMap.Exists(id) Then
                        wpxName = CStr(wpxMap(id))
                    Else
                        wpxName = gtName
                    End If
                    ws.Cells(outRow, 3).Value = wpxName
                    wpxRow = FindPlayerRow(wsWpxPT, wpxName, gtName)
                    If wpxRow = 0 Then
                        ws.Cells(outRow, 5).Value = "Skipped - no row for that name on WPX Player Tracking"
                    ElseIf Not IsWanted(gtName, id) Then
                        ws.Cells(outRow, 4).Value = wpxRow
                        ws.Cells(outRow, 5).Value = "Skipped - excluded by PLAYER_FILTER"
                    Else
                        ws.Cells(outRow, 4).Value = wpxRow
                        ws.Cells(outRow, 5).Value = "Pulled"
                    End If
                End If
            End If
        Next pr
    Next r

    ws.Columns("A:E").AutoFit
    ws.Activate

    If wpxWasOpened Then wbWpx.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Diagnostics written to the ""WPX Pull Diagnostics"" sheet.", vbInformation
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
    Set wpxMap = BuildWpxIdMap(wsWpxRoster)

    Dim pairs As Collection
    Set pairs = RosterColumnPairs(wsGTRoster)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long, pr As Variant
    lastRow = RosterLastRow(wsGTRoster, pairs)
    For r = 2 To lastRow
        For Each pr In pairs
            Call CheckAndAdd(coll, seen, wsGTRoster, wpxMap, r, pr(0), pr(0) + 1)
        Next pr
    Next r
End Sub

' Maps roster ID -> WPX player name for every ID/name block on the WPX Roster.
Private Function BuildWpxIdMap(wsWpxRoster As Worksheet) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim pairs As Collection
    Set pairs = RosterColumnPairs(wsWpxRoster)

    Dim lastRow As Long, r As Long, pr As Variant
    lastRow = RosterLastRow(wsWpxRoster, pairs)
    For r = 2 To lastRow
        For Each pr In pairs
            Call AddToIdMap(map, wsWpxRoster, r, pr(0), pr(0) + 1)
        Next pr
    Next r

    Set BuildWpxIdMap = map
End Function

' Roster sheets repeat blocks of ID, Name, (two spare columns) across the
' sheet. Returns Array(idCol) for every block that exists, so rosters with
' more than four blocks are picked up instead of being cut off at column N.
Private Function RosterColumnPairs(ws As Worksheet) As Collection
    Dim pairs As Collection
    Set pairs = New Collection

    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 14 Then lastCol = 14

    For c = 1 To lastCol Step 4
        pairs.Add Array(c)
    Next c

    Set RosterColumnPairs = pairs
End Function

' Last used row across every roster ID and name column, so a block that runs
' longer than column B is not cut short.
Private Function RosterLastRow(ws As Worksheet, pairs As Collection) As Long
    Dim maxRow As Long, pr As Variant, r As Long, i As Long
    For Each pr In pairs
        For i = 0 To 1
            r = ws.Cells(ws.Rows.Count, pr(0) + i).End(xlUp).Row
            If r > maxRow Then maxRow = r
        Next i
    Next pr
    RosterLastRow = maxRow
End Function

Private Sub AddToIdMap(map As Object, ws As Worksheet, r As Long, idCol As Long, nameCol As Long)
    Dim id As String, nameVal As String
    id = ParseId(ws.Cells(r, idCol).Value)
    nameVal = CleanText(ws.Cells(r, nameCol).Value)
    If id <> "" And nameVal <> "" Then map(id) = nameVal
End Sub

Private Sub CheckAndAdd(coll As Collection, seen As Object, ws As Worksheet, map As Object, r As Long, idCol As Long, nameCol As Long)
    Dim id As String, gtName As String, wpxName As String
    id = ParseId(ws.Cells(r, idCol).Value)
    gtName = CleanText(ws.Cells(r, nameCol).Value)
    If gtName = "" Then Exit Sub

    If id <> "" And map.Exists(id) Then
        wpxName = CStr(map(id))
    ElseIf MATCH_BY_NAME Then
        wpxName = gtName
    Else
        Exit Sub
    End If

    If wpxName = "" Then Exit Sub
    If Not IsWanted(gtName, id) Then Exit Sub

    ' A player listed in more than one roster block is only pulled once.
    Dim key As String
    key = id & "|" & NameKey(gtName)
    If seen.Exists(key) Then Exit Sub
    seen(key) = True

    coll.Add Array(gtName, wpxName, id)
End Sub

' Reads a roster ID stored as a number or as text with spaces, commas or a
' prefix, and returns it as a digits-only key ("" when there is no usable ID).
' Kept as text so IDs longer than a Long cannot overflow.
Private Function ParseId(v As Variant) As String
    If IsError(v) Then Exit Function

    Dim s As String, digits As String, i As Long, ch As String
    s = CleanText(v)

    ' Numbers can arrive in scientific notation; render them in full first.
    If IsNumeric(s) Then
        If InStr(1, s, "E", vbTextCompare) > 0 Then s = Format$(CDbl(s), "0")
    End If

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next i

    ' Drop leading zeros so "0123" and "123" are the same player.
    Do While Len(digits) > 1 And Left$(digits, 1) = "0"
        digits = Mid$(digits, 2)
    Loop

    If digits = "0" Then Exit Function
    ParseId = digits
End Function

' Trims a cell to plain text, dropping non-breaking and zero-width spaces
' that come along when names are pasted in from the game.
Private Function CleanText(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsError(v) Then Exit Function

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

' Comparison key for names: case, spaces and punctuation that differ between
' the two workbooks are ignored.
Private Function NameKey(v As Variant) As String
    Dim s As String, out As String, i As Long, ch As String
    s = LCase$(CleanText(v))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    NameKey = out
End Function

' True when PLAYER_FILTER is empty, or the player's GT name or ID is listed in it.
Private Function IsWanted(gtName As String, id As String) As Boolean
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
            If NameKey(token) = NameKey(gtName) Then
                IsWanted = True
                Exit Function
            End If
            If id <> "" Then
                If ParseId(token) = id Then
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

' Finds a player's row in column A by WPX name, falling back to the GT name so
' players renamed between the two workbooks still match.
Private Function FindPlayerRow(ws As Worksheet, wpxName As String, gtName As String) As Long
    FindPlayerRow = FindNameRow(ws, wpxName)
    If FindPlayerRow = 0 And gtName <> "" Then FindPlayerRow = FindNameRow(ws, gtName)
End Function

Private Function FindNameRow(ws As Worksheet, name As String) As Long
    Dim key As String
    key = NameKey(name)
    If key = "" Then Exit Function

    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If NameKey(ws.Cells(r, 1).Value) = key Then
            FindNameRow = r
            Exit Function
        End If
    Next r
End Function
