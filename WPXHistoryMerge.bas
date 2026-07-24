Attribute VB_Name = "WPXHistoryMerge"
' WPXHistoryMerge.bas
' Merges the legacy WPX Arena/HQ power history into GTStatsFINAL.xlsm for players
' whose roster ID exists in both the GT and WPX workbooks.
'
' It does NOT overwrite any existing GT Arena Power data. It appends the WPX
' history groups to the right of whatever history already exists on the GT
' Arena Power sheet, then stops so the next group start is always empty.
'
' To use:
'   1. Open GTStatsFINAL.xlsm.
'   2. Press Alt+F11 and import this file as a module.
'   3. Run MergeWPXHistory() from the macro list.
'   4. Select WPXStatsFinal.xlsm when prompted.

Option Explicit

Public Sub MergeWPXHistory()
    Dim wbWpx As Workbook
    Dim f As Variant
    Dim sameFolderPath As String

    ' Try to find WPXStatsFinal.xlsm in the same folder as this workbook first
    sameFolderPath = ThisWorkbook.Path & Application.PathSeparator & "WPXStatsFinal.xlsm"
    If Dir(sameFolderPath) <> "" Then
        f = sameFolderPath
    Else
        f = Application.GetOpenFilename( _
            "Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", , _
            "Select WPXStatsFinal.xlsm")
        If VarType(f) = vbBoolean Then Exit Sub
        If f = "False" Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set wbWpx = Workbooks.Open(CStr(f), ReadOnly:=True)

    Dim wsGTRoster As Worksheet, wsGTAP As Worksheet
    Dim wsWpxRoster As Worksheet, wsWpxAP As Worksheet
    Set wsGTRoster = ThisWorkbook.Sheets("Roster")
    Set wsGTAP = ThisWorkbook.Sheets("Arena Power")
    Set wsWpxRoster = wbWpx.Sheets("Roster")
    Set wsWpxAP = wbWpx.Sheets("Arena Power")

    Dim crossTeam As Collection
    Set crossTeam = New Collection
    Call BuildCrossTeamList(wsGTRoster, wsWpxRoster, crossTeam)

    Dim i As Long, p As Variant
    Dim mergedCount As Long
    mergedCount = 0
    For i = 1 To crossTeam.Count
        p = crossTeam(i)
        If MergePlayerHistory(wsGTAP, wsWpxAP, CStr(p(0)), CStr(p(1))) Then
            mergedCount = mergedCount + 1
        End If
    Next i

    wbWpx.Close SaveChanges:=False

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "WPX history merged for " & mergedCount & " cross-team player(s)." & vbCrLf & _
           "Total cross-team IDs found: " & crossTeam.Count, vbInformation
End Sub

' Build a list of cross-team players as Collection items: Array(gtName, wpxName, id)
Private Sub BuildCrossTeamList(wsGTRoster As Worksheet, wsWpxRoster As Worksheet, coll As Collection)
    Dim wpxMap As Object
    Set wpxMap = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long
    lastRow = MaxLastRow(wsWpxRoster, Array(2, 6, 10, 14))
    For r = 2 To lastRow
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 1, 2)
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 5, 6)
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 9, 10)
        Call AddToIdMap(wpxMap, wsWpxRoster, r, 13, 14)
    Next r

    lastRow = MaxLastRow(wsGTRoster, Array(2, 6, 10, 14))
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
    If IsNumeric(idVal) And Not IsEmpty(nameVal) And Trim(CStr(nameVal)) <> "" Then
        map(CLng(idVal)) = CStr(nameVal)
    End If
End Sub

Private Sub CheckAndAdd(coll As Collection, ws As Worksheet, map As Object, r As Long, idCol As Long, nameCol As Long)
    Dim idVal As Variant, id As Long
    idVal = ws.Cells(r, idCol).Value
    If IsNumeric(idVal) Then
        id = CLng(idVal)
        If map.Exists(id) Then
            Dim gtName As String, wpxName As String
            gtName = Trim(CStr(Nz(ws.Cells(r, nameCol).Value, "")))
            wpxName = Trim(CStr(map(id)))
            If gtName <> "" And wpxName <> "" Then
                coll.Add Array(gtName, wpxName, id)
            End If
        End If
    End If
End Sub

' Merges one player's WPX history into the GT Arena Power row.
' Returns True if any data was written, False if skipped or nothing to do.
Private Function MergePlayerHistory(wsGTAP As Worksheet, wsWpxAP As Worksheet, gtName As String, wpxName As String) As Boolean
    Dim gtRow As Long, wpxRow As Long
    gtRow = FindNameRow(wsGTAP, gtName)
    wpxRow = FindNameRow(wsWpxAP, wpxName)

    If gtRow = 0 Then
        Debug.Print "GT Arena Power row not found for " & gtName
        MergePlayerHistory = False
        Exit Function
    End If
    If wpxRow = 0 Then
        Debug.Print "WPX Arena Power row not found for " & wpxName
        MergePlayerHistory = False
        Exit Function
    End If

    ' If this player's WPX history is already present, skip to avoid duplicates.
    If WPXHistoryAlreadyMerged(wsGTAP, gtRow, wsWpxAP, wpxRow) Then
        Debug.Print "WPX history already merged for " & gtName
        MergePlayerHistory = False
        Exit Function
    End If

    Dim nextCol As Long
    nextCol = FindNextHistoryStart(wsGTAP, gtRow)

    ' Copy WPX current values (cols B-E) as the first history group, then every
    ' history group starting at col K (11).
    If CopyGroupIfPresent(wsWpxAP, wpxRow, 2, wsGTAP, gtRow, nextCol) Then nextCol = nextCol + 4

    Dim wpxCol As Long, wpxLastCol As Long
    wpxLastCol = wsWpxAP.Cells(wpxRow, wsWpxAP.Columns.Count).End(xlToLeft).Column
    For wpxCol = 11 To wpxLastCol Step 4
        If CopyGroupIfPresent(wsWpxAP, wpxRow, wpxCol, wsGTAP, gtRow, nextCol) Then
            nextCol = nextCol + 4
        End If
    Next wpxCol

    MergePlayerHistory = True
End Function

' Copies a 4-cell history group (Date, Level, Arena, HQ) from WPX to GT.
' Returns True if at least one cell in the WPX group had data.
Private Function CopyGroupIfPresent(wsSrc As Worksheet, rSrc As Long, cSrc As Long, _
    wsDst As Worksheet, rDst As Long, cDst As Long) As Boolean

    Dim i As Long, hasData As Boolean
    hasData = False
    For i = 0 To 3
        Dim v As Variant
        v = wsSrc.Cells(rSrc, cSrc + i).Value
        If Not IsEmpty(v) And Trim(CStr(v)) <> "" And Trim(CStr(v)) <> "-" Then
            hasData = True
        End If
    Next i

    If Not hasData Then
        CopyGroupIfPresent = False
        Exit Function
    End If

    For i = 0 To 3
        wsDst.Cells(rDst, cDst + i).Value = wsSrc.Cells(rSrc, cSrc + i).Value
    Next i
    CopyGroupIfPresent = True
End Function

' Returns True if any WPX history date is already present in the GT row.
Private Function WPXHistoryAlreadyMerged(wsGTAP As Worksheet, gtRow As Long, _
    wsWpxAP As Worksheet, wpxRow As Long) As Boolean

    Dim wpxDates As Collection
    Set wpxDates = New Collection

    Dim wpxCol As Long, wpxLastCol As Long, d As Variant
    wpxLastCol = wsWpxAP.Cells(wpxRow, wsWpxAP.Columns.Count).End(xlToLeft).Column

    ' Include current date (col B) and all history dates.
    d = wsWpxAP.Cells(wpxRow, 2).Value
    If IsDate(d) Then wpxDates.Add d

    For wpxCol = 11 To wpxLastCol Step 4
        d = wsWpxAP.Cells(wpxRow, wpxCol).Value
        If IsDate(d) Then wpxDates.Add d
    Next wpxCol

    If wpxDates.Count = 0 Then
        WPXHistoryAlreadyMerged = False
        Exit Function
    End If

    Dim gtLastCol As Long, gtCol As Long
    gtLastCol = wsGTAP.Cells(gtRow, wsGTAP.Columns.Count).End(xlToLeft).Column

    For gtCol = 11 To gtLastCol Step 4
        If IsDate(wsGTAP.Cells(gtRow, gtCol).Value) Then
            Dim gtDate As Date
            gtDate = CDate(wsGTAP.Cells(gtRow, gtCol).Value)
            For Each d In wpxDates
                If CDate(d) = gtDate Then
                    WPXHistoryAlreadyMerged = True
                    Exit Function
                End If
            Next d
        End If
    Next gtCol

    WPXHistoryAlreadyMerged = False
End Function

' Finds the next empty history group start column for the given row.
' History groups start at column K (11) and repeat every 4 columns.
Private Function FindNextHistoryStart(ws As Worksheet, r As Long) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column

    If lastCol < 11 Then
        FindNextHistoryStart = 11
        Exit Function
    End If

    Dim n As Long, s As Long
    n = (lastCol - 11) \ 4
    s = 11 + 4 * n
    If s <= lastCol Then s = s + 4
    FindNextHistoryStart = s
End Function

Private Function FindNameRow(ws As Worksheet, name As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(LCase(CStr(Nz(ws.Cells(r, 1).Value, "")))) = LCase(name) Then
            FindNameRow = r
            Exit Function
        End If
    Next r
    FindNameRow = 0
End Function

Private Function MaxLastRow(ws As Worksheet, cols As Variant) As Long
    Dim maxLr As Long, lr As Long, i As Long
    maxLr = 0
    For i = LBound(cols) To UBound(cols)
        lr = ws.Cells(ws.Rows.Count, CLng(cols(i))).End(xlUp).Row
        If lr > maxLr Then maxLr = lr
    Next i
    If maxLr < 2 Then maxLr = 2
    MaxLastRow = maxLr
End Function

Private Function Nz(v As Variant, defaultVal As Variant) As Variant
    If IsEmpty(v) Or v = "" Then
        Nz = defaultVal
    Else
        Nz = v
    End If
End Function
