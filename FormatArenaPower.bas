Attribute VB_Name = "FormatArenaPower"
' ============================================================
' FormatArenaPower — VBA Macro Module
' Rewrites the power numbers on the "Arena Power" sheet as text
' in the "161.66 M" format.
'
' Which cells are formatted (row 1 is the header, data from row 2):
'   * D = Arena Power, E = HQ Power
'   * F,G = session deltas, H,I = overall deltas
'   * History arena/HQ columns (the arena & HQ cells in each
'     4-column history group: date, level, arena, HQ starting at K)
' Skipped: name, dates, level, and level-note cells.
'
' Values already in millions (157.7, "153.4 M") are kept as-is;
' full raw counts (>= 100000, e.g. 161660000) are divided by
' 1,000,000. Blank and "-" cells are left untouched.
'
' USAGE:
'   1. Import this .bas file:
'      Developer > Visual Basic > File > Import File
'   2. Run "FormatArenaPowerNumbers" from the Macros dialog,
'      or assign it to a button.
' ============================================================

Option Explicit

Private Const SHEET_NAME As String = "Arena Power"
Private Const FIRST_ROW  As Long = 2

Public Sub FormatArenaPowerNumbers()

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet """ & SHEET_NAME & """ not found.", _
               vbExclamation, "Missing Sheet"
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long, off As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < FIRST_ROW Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    For r = FIRST_ROW To lastRow
        ' Current arena/HQ + deltas: columns D..I (4..9)
        For c = 4 To 9
            FormatPowerCell ws.Cells(r, c)
        Next c

        ' History groups of 4 starting at column K (11):
        '   11=date, 12=level, 13=arena, 14=HQ, 15=date, ...
        ' (c - 11) Mod 4 = 2 -> arena, = 3 -> HQ
        For c = 13 To lastCol
            off = (c - 11) Mod 4
            If off = 2 Or off = 3 Then FormatPowerCell ws.Cells(r, c)
        Next c
    Next r

CleanUp:
    Application.ScreenUpdating = True

End Sub

' --- Convert a single cell to "0.00 M" text (in millions) ------
Private Sub FormatPowerCell(ByVal cell As Range)

    Dim v As Variant
    v = cell.Value

    If IsError(v) Then Exit Sub

    Dim s As String
    s = Trim(CStr(v))
    If s = "" Or s = "-" Then Exit Sub

    ' Strip any existing "M"/"Mil" suffix and spaces
    s = Replace(s, "Mil", "", , , vbTextCompare)
    s = Replace(s, "M", "", , , vbTextCompare)
    s = Replace(s, " ", "")
    s = Trim(s)

    If Not IsNumeric(s) Then Exit Sub

    Dim n As Double
    n = CDbl(s)
    If Abs(n) >= 100000 Then n = n / 1000000    ' raw count -> millions

    cell.Value = Format(n, "0.00") & " M"

End Sub
