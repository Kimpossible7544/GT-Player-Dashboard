Attribute VB_Name = "ImportKimWPXArena"
' ============================================================
' ImportKimWPXArena - VBA Macro Module
' Writes a one-time snapshot of Kimpossible's WPX "Arena Power"
' row (current + full history) onto the GT "Arena Power" sheet
' as a labeled block, parked out of the way to the right.
'
' Placement: the block starts at column START_COL (default 220).
' The sheet's weekly history begins at column K (11) and each
' week uses 4 columns, so 52 weeks fill through column 218; the
' block sits past that with a gap, so ~a year of new weekly data
' can be added before it is reached.
'
' Power values are written as text in the "X.XX M" format to
' match the rest of the sheet. Dates and levels are written as
' real values.
'
' USAGE:
'   1. Import this .bas: Developer > Visual Basic > File > Import File
'   2. Run "ImportKimWPXArenaBlock" from the Macros dialog.
' ============================================================

Option Explicit

Private Const SHEET_NAME As String = "Arena Power"
Private Const START_COL  As Long = 220   ' first column of the block (past 52 weeks)

Public Sub ImportKimWPXArenaBlock()

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet """ & SHEET_NAME & """ not found.", _
               vbExclamation, "Missing Sheet"
        Exit Sub
    End If

    Dim titleRow As Long, hdrRow As Long, datRow As Long
    titleRow = 1
    hdrRow = 2
    datRow = 3

    Application.ScreenUpdating = False

    ws.Cells(titleRow, START_COL).Value = "Kimpossible - WPX Arena Power (imported snapshot)"

    ' --- Header row ---
    ws.Cells(hdrRow, START_COL + 0).Value = "Player"
    ws.Cells(hdrRow, START_COL + 1).Value = "Date"
    ws.Cells(hdrRow, START_COL + 2).Value = "Level"
    ws.Cells(hdrRow, START_COL + 3).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 4).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 5).Value = "Arena Chg"
    ws.Cells(hdrRow, START_COL + 6).Value = "HQ Chg"
    ws.Cells(hdrRow, START_COL + 7).Value = "Overall Arena Chg"
    ws.Cells(hdrRow, START_COL + 8).Value = "Overall HQ Chg"
    ws.Cells(hdrRow, START_COL + 9).Value = "Level Note"
    ws.Cells(hdrRow, START_COL + 10).Value = "Date"
    ws.Cells(hdrRow, START_COL + 11).Value = "Level"
    ws.Cells(hdrRow, START_COL + 12).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 13).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 14).Value = "Date"
    ws.Cells(hdrRow, START_COL + 15).Value = "Level"
    ws.Cells(hdrRow, START_COL + 16).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 17).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 18).Value = "Date"
    ws.Cells(hdrRow, START_COL + 19).Value = "Level"
    ws.Cells(hdrRow, START_COL + 20).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 21).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 22).Value = "Date"
    ws.Cells(hdrRow, START_COL + 23).Value = "Level"
    ws.Cells(hdrRow, START_COL + 24).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 25).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 26).Value = "Date"
    ws.Cells(hdrRow, START_COL + 27).Value = "Level"
    ws.Cells(hdrRow, START_COL + 28).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 29).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 30).Value = "Date"
    ws.Cells(hdrRow, START_COL + 31).Value = "Level"
    ws.Cells(hdrRow, START_COL + 32).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 33).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 34).Value = "Date"
    ws.Cells(hdrRow, START_COL + 35).Value = "Level"
    ws.Cells(hdrRow, START_COL + 36).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 37).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 38).Value = "Date"
    ws.Cells(hdrRow, START_COL + 39).Value = "Level"
    ws.Cells(hdrRow, START_COL + 40).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 41).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 42).Value = "Date"
    ws.Cells(hdrRow, START_COL + 43).Value = "Level"
    ws.Cells(hdrRow, START_COL + 44).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 45).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 46).Value = "Date"
    ws.Cells(hdrRow, START_COL + 47).Value = "Level"
    ws.Cells(hdrRow, START_COL + 48).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 49).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 50).Value = "Date"
    ws.Cells(hdrRow, START_COL + 51).Value = "Level"
    ws.Cells(hdrRow, START_COL + 52).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 53).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 54).Value = "Date"
    ws.Cells(hdrRow, START_COL + 55).Value = "Level"
    ws.Cells(hdrRow, START_COL + 56).Value = "Arena Power"
    ws.Cells(hdrRow, START_COL + 57).Value = "HQ Power"
    ws.Cells(hdrRow, START_COL + 58).Value = "Date"
    ws.Cells(hdrRow, START_COL + 59).Value = "Level"
    ws.Cells(hdrRow, START_COL + 60).Value = "Arena Power"

    ' --- Kimpossible WPX data row ---
    ws.Cells(datRow, START_COL + 0).Value = "Kimpossible7544"
    ws.Cells(datRow, START_COL + 1).Value = DateSerial(2026, 7, 9)
    ws.Cells(datRow, START_COL + 2).Value = 31
    ws.Cells(datRow, START_COL + 3).Value = "157.70 M"
    ws.Cells(datRow, START_COL + 4).Value = "226.90 M"
    ws.Cells(datRow, START_COL + 5).Value = "4.30 M"
    ws.Cells(datRow, START_COL + 6).Value = "18.40 M"
    ws.Cells(datRow, START_COL + 7).Value = "61.70 M"
    ws.Cells(datRow, START_COL + 8).Value = "35.30 M"
    ws.Cells(datRow, START_COL + 9).Value = "-"
    ws.Cells(datRow, START_COL + 10).Value = DateSerial(2026, 7, 2)
    ws.Cells(datRow, START_COL + 11).Value = 31
    ws.Cells(datRow, START_COL + 12).Value = "153.40 M"
    ws.Cells(datRow, START_COL + 13).Value = "208.50 M"
    ws.Cells(datRow, START_COL + 14).Value = DateSerial(2026, 6, 26)
    ws.Cells(datRow, START_COL + 15).Value = 31
    ws.Cells(datRow, START_COL + 16).Value = "145.00 M"
    ws.Cells(datRow, START_COL + 17).Value = "211.40 M"
    ws.Cells(datRow, START_COL + 18).Value = DateSerial(2026, 6, 19)
    ws.Cells(datRow, START_COL + 19).Value = 31
    ws.Cells(datRow, START_COL + 20).Value = "137.80 M"
    ws.Cells(datRow, START_COL + 21).Value = "212.40 M"
    ws.Cells(datRow, START_COL + 22).Value = DateSerial(2026, 6, 12)
    ws.Cells(datRow, START_COL + 23).Value = 30
    ws.Cells(datRow, START_COL + 24).Value = "133.90 M"
    ws.Cells(datRow, START_COL + 25).Value = "208.80 M"
    ws.Cells(datRow, START_COL + 26).Value = DateSerial(2026, 6, 7)
    ws.Cells(datRow, START_COL + 27).Value = 30
    ws.Cells(datRow, START_COL + 28).Value = "131.10 M"
    ws.Cells(datRow, START_COL + 29).Value = "179.20 M"
    ws.Cells(datRow, START_COL + 30).Value = DateSerial(2026, 6, 5)
    ws.Cells(datRow, START_COL + 31).Value = 30
    ws.Cells(datRow, START_COL + 32).Value = "129.10 M"
    ws.Cells(datRow, START_COL + 33).Value = "191.60 M"
    ws.Cells(datRow, START_COL + 34).Value = DateSerial(2026, 5, 30)
    ws.Cells(datRow, START_COL + 35).Value = 30
    ws.Cells(datRow, START_COL + 36).Value = "125.20 M"
    ws.Cells(datRow, START_COL + 38).Value = DateSerial(2026, 5, 18)
    ws.Cells(datRow, START_COL + 39).Value = 30
    ws.Cells(datRow, START_COL + 40).Value = "106.20 M"
    ws.Cells(datRow, START_COL + 42).Value = DateSerial(2026, 5, 16)
    ws.Cells(datRow, START_COL + 43).Value = 30
    ws.Cells(datRow, START_COL + 44).Value = "105.80 M"
    ws.Cells(datRow, START_COL + 46).Value = DateSerial(2026, 5, 16)
    ws.Cells(datRow, START_COL + 47).Value = 30
    ws.Cells(datRow, START_COL + 48).Value = "105.80 M"
    ws.Cells(datRow, START_COL + 50).Value = DateSerial(2026, 5, 11)
    ws.Cells(datRow, START_COL + 51).Value = 30
    ws.Cells(datRow, START_COL + 52).Value = "102.00 M"
    ws.Cells(datRow, START_COL + 54).Value = DateSerial(2026, 5, 4)
    ws.Cells(datRow, START_COL + 55).Value = 29
    ws.Cells(datRow, START_COL + 56).Value = "96.20 M"
    ws.Cells(datRow, START_COL + 58).Value = DateSerial(2026, 5, 1)
    ws.Cells(datRow, START_COL + 59).Value = 29
    ws.Cells(datRow, START_COL + 60).Value = "96.00 M"

    ws.Cells(titleRow, START_COL).Font.Bold = True
    Application.ScreenUpdating = True

    MsgBox "Kimpossible's WPX Arena Power snapshot written starting at column " & _
           START_COL & " (row " & titleRow & ").", vbInformation, "Done"

End Sub
