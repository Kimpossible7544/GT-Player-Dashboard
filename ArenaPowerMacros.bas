Attribute VB_Name = "ArenaPowerMacros"
' -------------------------------------------------------
' Looks up a Score Import name against the Roster sheet
' Checks Player Name cols (B, F, J, N) and AKA cols (D, H, L, P)
' Returns the matching Player Name (original casing), or empty string
' -------------------------------------------------------
Function LookupRosterName(wsRoster As Worksheet, importName As String) As String

    Dim lastRow As Long
    Dim i As Long
    Dim playerName As String
    Dim aka As String
    Dim normImport As String
    Dim lastRowB As Long, lastRowF As Long, lastRowJ As Long, lastRowN As Long

    normImport = NormalizeArenaName(importName)

    lastRowB = wsRoster.Cells(wsRoster.Rows.count, "B").End(xlUp).row
    lastRowF = wsRoster.Cells(wsRoster.Rows.count, "F").End(xlUp).row
    lastRowJ = wsRoster.Cells(wsRoster.Rows.count, "J").End(xlUp).row
    lastRowN = wsRoster.Cells(wsRoster.Rows.count, "N").End(xlUp).row
    lastRow = Application.Max(lastRowB, lastRowF, lastRowJ, lastRowN)

    Dim nameCols(3) As Integer
    Dim akaCols(3) As Integer
    nameCols(0) = 2:  nameCols(1) = 6:  nameCols(2) = 10: nameCols(3) = 14
    akaCols(0) = 4:   akaCols(1) = 8:   akaCols(2) = 12: akaCols(3) = 16

    For i = 2 To lastRow
        Dim s As Integer
        For s = 0 To 3
            playerName = Trim(wsRoster.Cells(i, nameCols(s)).Value)
            If playerName <> "" Then
                If NormalizeArenaName(playerName) = normImport Then
                    LookupRosterName = playerName
                    Exit Function
                End If
                aka = Trim(wsRoster.Cells(i, akaCols(s)).Value)
                If aka <> "" Then
                    Dim akaParts() As String
                    Dim delim As String
                    If InStr(aka, ",") > 0 Then
                        delim = ","
                    ElseIf InStr(aka, "/") > 0 Then
                        delim = "/"
                    Else
                        delim = ""
                    End If
                    If delim = "" Then
                        If NormalizeArenaName(aka) = normImport Then
                            LookupRosterName = playerName
                            Exit Function
                        End If
                    Else
                        akaParts = Split(aka, delim)
                        Dim p As Integer
                        For p = 0 To UBound(akaParts)
                            If NormalizeArenaName(Trim(akaParts(p))) = normImport Then
                                LookupRosterName = playerName
                                Exit Function
                            End If
                        Next p
                    End If
                End If
            End If
        Next s
    Next i

    LookupRosterName = ""

End Function

' -------------------------------------------------------
' Normalizes a name for comparison only - does NOT affect display
' -------------------------------------------------------
Function NormalizeArenaName(val As String) As String
    Dim result As String
    Dim c As String
    Dim i As Integer

    result = ""
    val = UCase(Trim(val))

    For i = 1 To Len(val)
        c = Mid(val, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then
            result = result & c
        End If
    Next i

    NormalizeArenaName = result
End Function

' -------------------------------------------------------
Sub UpdateArenaScores()

    Dim wsMain As Worksheet
    Dim wsImport As Worksheet
    Dim wsRoster As Worksheet
    Dim lastRowMain As Long
    Dim lastRowImport As Long
    Dim i As Long, j As Long, c As Long
    Dim importName As String
    Dim rosterName As String
    Dim notFound As String
    Dim noData As String
    Dim matchFound As Boolean
    Dim newArenaStr As String
    Dim oldArenaStr As String
    Dim newHQStr As String
    Dim oldHQStr As String
    Dim firstArenaStr As String
    Dim firstHQStr As String
    Dim importDate As Variant
    Dim newRow As Long
    Dim mainName As String
    Dim foundInImport As Boolean
    Dim msg As String
    Dim lastCol As Long

    Set wsMain = ThisWorkbook.Sheets("Arena Power")
    Set wsImport = ThisWorkbook.Sheets("Score Import")
    Set wsRoster = ThisWorkbook.Sheets("Roster")

    lastRowMain = wsMain.Cells(wsMain.Rows.count, "A").End(xlUp).row
    lastRowImport = wsImport.Cells(wsImport.Rows.count, "I").End(xlUp).row

    If MsgBox("Update all player scores from Score Import sheet?", _
              vbYesNo + vbQuestion, "Update Scores") = vbNo Then Exit Sub

    notFound = ""
    noData = ""

    For i = 2 To lastRowImport

        importName = Trim(wsImport.Cells(i, "I").Value)
        If importName = "" Then GoTo NextImport

        rosterName = LookupRosterName(wsRoster, importName)
        If rosterName = "" Then
            notFound = notFound & "  - " & importName & " (no roster match)" & vbNewLine
            GoTo NextImport
        End If

        matchFound = False

        For j = 2 To lastRowMain
            If NormalizeArenaName(Trim(wsMain.Cells(j, "A").Value)) = NormalizeArenaName(rosterName) Then
                matchFound = True

                importDate = wsImport.Cells(i, "J").Value

                If CStr(importDate) <> CStr(wsMain.Cells(j, "B").Value) Then

                    lastCol = wsMain.Cells(j, wsMain.Columns.count).End(xlToLeft).Column
                    For c = lastCol To 11 Step -1
                        wsMain.Cells(j, c + 4).Value = wsMain.Cells(j, c).Value
                    Next c

                    wsMain.Cells(j, 11).Value = wsMain.Cells(j, "B").Value
                    wsMain.Cells(j, 12).Value = wsMain.Cells(j, "C").Value
                    wsMain.Cells(j, 13).Value = wsMain.Cells(j, "D").Value
                    wsMain.Cells(j, 14).Value = wsMain.Cells(j, "E").Value

                    wsMain.Cells(j, "B").Value = importDate
                    wsMain.Cells(j, "C").Value = wsImport.Cells(i, "K").Value
                    wsMain.Cells(j, "D").Value = CleanPower(CStr(wsImport.Cells(i, "L").Value))
                    wsMain.Cells(j, "E").Value = CleanPower(CStr(wsImport.Cells(i, "M").Value))

                End If

                newArenaStr = CleanPower(CStr(wsMain.Cells(j, "D").Value))
                oldArenaStr = CleanPower(CStr(wsMain.Cells(j, 13).Value))
                If IsNumeric(newArenaStr) And IsNumeric(oldArenaStr) And newArenaStr <> "" And oldArenaStr <> "" Then
                    wsMain.Cells(j, "F").Value = Round(CDbl(newArenaStr) - CDbl(oldArenaStr), 2)
                    ColorCell wsMain.Cells(j, "F"), Round(CDbl(newArenaStr) - CDbl(oldArenaStr), 2)
                Else
                    wsMain.Cells(j, "F").Value = "-"
                    wsMain.Cells(j, "F").Font.ColorIndex = xlAutomatic
                End If

                newHQStr = CleanPower(CStr(wsMain.Cells(j, "E").Value))
                oldHQStr = CleanPower(CStr(wsMain.Cells(j, 14).Value))
                If IsNumeric(newHQStr) And IsNumeric(oldHQStr) And newHQStr <> "" And oldHQStr <> "" Then
                    wsMain.Cells(j, "G").Value = Round(CDbl(newHQStr) - CDbl(oldHQStr), 2)
                    ColorCell wsMain.Cells(j, "G"), Round(CDbl(newHQStr) - CDbl(oldHQStr), 2)
                Else
                    wsMain.Cells(j, "G").Value = "-"
                    wsMain.Cells(j, "G").Font.ColorIndex = xlAutomatic
                End If

                firstArenaStr = GetFirstArenaPower(wsMain, j)
                If IsNumeric(newArenaStr) And newArenaStr <> "" And firstArenaStr <> "" Then
                    wsMain.Cells(j, "H").Value = Round(CDbl(newArenaStr) - CDbl(firstArenaStr), 2)
                    ColorCell wsMain.Cells(j, "H"), Round(CDbl(newArenaStr) - CDbl(firstArenaStr), 2)
                Else
                    wsMain.Cells(j, "H").Value = "-"
                    wsMain.Cells(j, "H").Font.ColorIndex = xlAutomatic
                End If

                firstHQStr = GetFirstHQPower(wsMain, j)
                If IsNumeric(newHQStr) And newHQStr <> "" And firstHQStr <> "" Then
                    wsMain.Cells(j, "I").Value = Round(CDbl(newHQStr) - CDbl(firstHQStr), 2)
                    ColorCell wsMain.Cells(j, "I"), Round(CDbl(newHQStr) - CDbl(firstHQStr), 2)
                Else
                    wsMain.Cells(j, "I").Value = "-"
                    wsMain.Cells(j, "I").Font.ColorIndex = xlAutomatic
                End If

                If wsMain.Cells(j, 12).Value = "" Then
                    wsMain.Cells(j, "J").Value = "-"
                ElseIf CStr(wsMain.Cells(j, "C").Value) <> CStr(wsMain.Cells(j, 12).Value) Then
                    wsMain.Cells(j, "J").Value = "Level up: " & wsMain.Cells(j, 12).Value & " -> " & wsMain.Cells(j, "C").Value
                Else
                    wsMain.Cells(j, "J").Value = "-"
                End If

                Exit For
            End If
        Next j

        If Not matchFound Then
            newRow = wsMain.Cells(wsMain.Rows.count, "A").End(xlUp).row + 1
            wsMain.Cells(newRow, "A").Value = rosterName
            wsMain.Cells(newRow, "B").Value = wsImport.Cells(i, "J").Value
            wsMain.Cells(newRow, "C").Value = wsImport.Cells(i, "K").Value
            wsMain.Cells(newRow, "D").Value = CleanPower(CStr(wsImport.Cells(i, "L").Value))
            wsMain.Cells(newRow, "E").Value = CleanPower(CStr(wsImport.Cells(i, "M").Value))
            wsMain.Cells(newRow, "F").Value = "-"
            wsMain.Cells(newRow, "G").Value = "-"
            wsMain.Cells(newRow, "H").Value = "-"
            wsMain.Cells(newRow, "I").Value = "-"
            wsMain.Cells(newRow, "J").Value = "-"
            notFound = notFound & "  + " & rosterName & " (new row added)" & vbNewLine
        End If

NextImport:
    Next i

    lastRowMain = wsMain.Cells(wsMain.Rows.count, "A").End(xlUp).row
    For j = 2 To lastRowMain
        mainName = Trim(wsMain.Cells(j, "A").Value)
        If mainName <> "" Then
            foundInImport = False
            For i = 2 To lastRowImport
                Dim resolvedName As String
                resolvedName = LookupRosterName(wsRoster, Trim(wsImport.Cells(i, "I").Value))
                If NormalizeArenaName(resolvedName) = NormalizeArenaName(mainName) Then
                    foundInImport = True
                    Exit For
                End If
            Next i
            If Not foundInImport Then
                noData = noData & "  - " & mainName & vbNewLine
            End If
        End If
    Next j

    Call GrayEmptyStaleRows

    msg = "Update complete!"
    If notFound <> "" Then msg = msg & vbNewLine & vbNewLine & "NEW players added:" & vbNewLine & notFound
    If noData <> "" Then msg = msg & vbNewLine & vbNewLine & "Players with NO import data:" & vbNewLine & noData
    If notFound = "" And noData = "" Then msg = msg & " All players matched successfully."

    MsgBox msg, vbInformation, "Update Complete"

End Sub

' -------------------------------------------------------
Function CleanPower(val As String) As String
    CleanPower = Replace(Replace(Replace(val, "Mil", ""), "M", ""), " ", "")
End Function

' -------------------------------------------------------
Sub ColorCell(cell As Range, diff As Double)
    If diff > 0 Then
        cell.Font.Color = RGB(0, 176, 80)
    ElseIf diff < 0 Then
        cell.Font.Color = RGB(255, 0, 0)
    Else
        cell.Font.ColorIndex = xlAutomatic
    End If
End Sub

' -------------------------------------------------------
Function GetFirstArenaPower(ws As Worksheet, rowNum As Long) As String
    Dim c As Long
    Dim lastCol As Long
    Dim val As String
    Dim firstVal As String

    firstVal = ""
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column

    For c = 13 To lastCol Step 4
        val = CleanPower(CStr(ws.Cells(rowNum, c).Value))
        If IsNumeric(val) And val <> "" Then firstVal = val
    Next c

    GetFirstArenaPower = firstVal
End Function

' -------------------------------------------------------
Function GetFirstHQPower(ws As Worksheet, rowNum As Long) As String
    Dim c As Long
    Dim lastCol As Long
    Dim val As String
    Dim firstVal As String

    firstVal = ""
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column

    For c = 14 To lastCol Step 4
        val = CleanPower(CStr(ws.Cells(rowNum, c).Value))
        If IsNumeric(val) And val <> "" Then firstVal = val
    Next c

    GetFirstHQPower = firstVal
End Function

' -------------------------------------------------------
Sub UpdateChangeColumn()

    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim j As Long
    Dim newArenaStr As String
    Dim oldArenaStr As String
    Dim newHQStr As String
    Dim oldHQStr As String
    Dim firstArenaStr As String
    Dim firstHQStr As String
    Dim hoStr As String

    Set wsMain = ThisWorkbook.Sheets("Arena Power")
    lastRow = wsMain.Cells(wsMain.Rows.count, "A").End(xlUp).row

    For j = 2 To lastRow
        If Trim(wsMain.Cells(j, "A").Value) <> "" Then

            newArenaStr = CleanPower(CStr(wsMain.Cells(j, "D").Value))
            newHQStr = CleanPower(CStr(wsMain.Cells(j, "E").Value))
            oldArenaStr = CleanPower(CStr(wsMain.Cells(j, 13).Value))
            oldHQStr = CleanPower(CStr(wsMain.Cells(j, 14).Value))
            firstArenaStr = GetFirstArenaPower(wsMain, j)
            firstHQStr = GetFirstHQPower(wsMain, j)

            ' Kimpossible: use the WPX arena power parked in column HO as the
            ' overall baseline, so Overall Arena Chg (H) = D - HO.
            If NormalizeArenaName(Trim(wsMain.Cells(j, "A").Value)) = NormalizeArenaName("Kimpossible7544") Then
                hoStr = CleanPower(CStr(wsMain.Cells(j, "HO").Value))
                If IsNumeric(hoStr) And hoStr <> "" Then firstArenaStr = hoStr
            End If

            If IsNumeric(newArenaStr) And IsNumeric(oldArenaStr) And newArenaStr <> "" And oldArenaStr <> "" Then
                wsMain.Cells(j, "F").Value = Round(CDbl(newArenaStr) - CDbl(oldArenaStr), 2)
                ColorCell wsMain.Cells(j, "F"), Round(CDbl(newArenaStr) - CDbl(oldArenaStr), 2)
            Else
                wsMain.Cells(j, "F").Value = "-"
                wsMain.Cells(j, "F").Font.ColorIndex = xlAutomatic
            End If

            If IsNumeric(newHQStr) And IsNumeric(oldHQStr) And newHQStr <> "" And oldHQStr <> "" Then
                wsMain.Cells(j, "G").Value = Round(CDbl(newHQStr) - CDbl(oldHQStr), 2)
                ColorCell wsMain.Cells(j, "G"), Round(CDbl(newHQStr) - CDbl(oldHQStr), 2)
            Else
                wsMain.Cells(j, "G").Value = "-"
                wsMain.Cells(j, "G").Font.ColorIndex = xlAutomatic
            End If

            If IsNumeric(newArenaStr) And newArenaStr <> "" And firstArenaStr <> "" Then
                wsMain.Cells(j, "H").Value = Round(CDbl(newArenaStr) - CDbl(firstArenaStr), 2)
                ColorCell wsMain.Cells(j, "H"), Round(CDbl(newArenaStr) - CDbl(firstArenaStr), 2)
            Else
                wsMain.Cells(j, "H").Value = "-"
                wsMain.Cells(j, "H").Font.ColorIndex = xlAutomatic
            End If

            If IsNumeric(newHQStr) And newHQStr <> "" And firstHQStr <> "" Then
                wsMain.Cells(j, "I").Value = Round(CDbl(newHQStr) - CDbl(firstHQStr), 2)
                ColorCell wsMain.Cells(j, "I"), Round(CDbl(newHQStr) - CDbl(firstHQStr), 2)
            Else
                wsMain.Cells(j, "I").Value = "-"
                wsMain.Cells(j, "I").Font.ColorIndex = xlAutomatic
            End If

        End If
    Next j

    MsgBox "Change columns updated!", vbInformation

End Sub

' -------------------------------------------------------
Sub GrayEmptyStaleRows()

    Dim wsMain As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim j As Long
    Dim k As Long
    Dim grayColor As Long

    grayColor = RGB(180, 180, 180)
    Set wsMain = ThisWorkbook.Sheets("Arena Power")

    lastRow = wsMain.Cells(wsMain.Rows.count, "A").End(xlUp).row
    lastCol = wsMain.Cells(1, wsMain.Columns.count).End(xlToLeft).Column

    For j = 2 To lastRow
        If Trim(wsMain.Cells(j, "A").Value) <> "" Then
            For k = 11 To lastCol Step 4
                Call ApplyGray(wsMain, j, k, grayColor, 4)
            Next k
        End If
    Next j

    MsgBox "Graying complete!", vbInformation

End Sub

' -------------------------------------------------------
Sub ApplyGray(ws As Worksheet, rowNum As Long, dateCol As Long, grayColor As Long, groupSize As Long)

    Dim c As Long
    Dim groupDate As Date
    Dim dateVal As Variant
    Dim isValidPastDate As Boolean

    isValidPastDate = False
    dateVal = ws.Cells(rowNum, dateCol).Value

    If dateVal <> "" Then
        On Error Resume Next
        groupDate = CDate(dateVal)
        If Err.Number = 0 Then
            If groupDate < Date Then isValidPastDate = True
        End If
        On Error GoTo 0
    End If

    If isValidPastDate Then
        For c = dateCol To dateCol + groupSize - 1
            If ws.Cells(rowNum, c).Value = "" Then
                ws.Cells(rowNum, c).Interior.Color = grayColor
            Else
                ws.Cells(rowNum, c).Interior.ColorIndex = xlNone
            End If
        Next c
    Else
        For c = dateCol To dateCol + groupSize - 1
            ws.Cells(rowNum, c).Interior.ColorIndex = xlNone
        Next c
    End If

End Sub

' -------------------------------------------------------
' Fixes casing of all names in Arena Power col A
' by direct comparison against Roster name columns
' -------------------------------------------------------
Sub FixArenaSheetCasing()

    Dim wsMain As Worksheet
    Dim wsRoster As Worksheet
    Dim lastRowMain As Long
    Dim lastRowRoster As Long
    Dim i As Long, j As Long, s As Long
    Dim cellName As String
    Dim candidate As String
    Dim fixed As Long

    Dim nameCols(3) As Integer
    nameCols(0) = 2
    nameCols(1) = 6
    nameCols(2) = 10
    nameCols(3) = 14

    Set wsMain = ThisWorkbook.Sheets("Arena Power")
    Set wsRoster = ThisWorkbook.Sheets("Roster")

    lastRowMain = wsMain.Cells(wsMain.Rows.count, "A").End(xlUp).row
    lastRowRoster = Application.Max( _
        wsRoster.Cells(wsRoster.Rows.count, "B").End(xlUp).row, _
        wsRoster.Cells(wsRoster.Rows.count, "F").End(xlUp).row, _
        wsRoster.Cells(wsRoster.Rows.count, "J").End(xlUp).row, _
        wsRoster.Cells(wsRoster.Rows.count, "N").End(xlUp).row)

    fixed = 0

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For j = 2 To lastRowMain
        cellName = Trim(wsMain.Cells(j, "A").Value)
        If cellName = "" Then GoTo NextRow

        For i = 2 To lastRowRoster
            For s = 0 To 3
                candidate = Trim(wsRoster.Cells(i, nameCols(s)).Value)
                If UCase(candidate) = UCase(cellName) And candidate <> cellName Then
                    wsMain.Cells(j, "A").Value = candidate
                    fixed = fixed + 1
                    GoTo NextRow
                End If
            Next s
        Next i

NextRow:
    Next j

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Done - " & fixed & " name(s) corrected.", vbInformation, "Fix Casing Complete"

End Sub

