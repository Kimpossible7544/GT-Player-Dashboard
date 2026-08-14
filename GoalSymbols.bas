Attribute VB_Name = "GoalSymbols"
' ============================================================
' GoalSymbols  VBA Macro Module
' Fills column N of the weekly sheets with a goal indicator
' based on the weekly total in column M:
'
'   * Gold star   -> M >= 20,000,000  (weekly goal hit)
'   * Red "X"     -> data exists in C:G AND M < 20,000,000
'   * (blank)     -> no data in C:G
'
' USAGE  RUN NOW (fill/refresh the whole sheet):
'   1. Import this .bas file into your workbook:
'      Developer > Visual Basic > File > Import File
'   2. Open the weekly sheet you want to update
'   3. Run the macro "UpdateGoalSymbols" from the Macros dialog,
'      or assign it to a button:
'      Insert > Shapes > pick a shape > right-click >
'      Assign Macro > select "UpdateGoalSymbols"
'
' USAGE  AUTOMATIC (update as soon as M changes):
'   Paste the Workbook_SheetChange handler shown in the comment
'   block at the bottom of this module into the "ThisWorkbook"
'   object (VBA editor > double-click ThisWorkbook). It calls
'   UpdateGoalSymbolsRow whenever a relevant cell changes.
' ============================================================

Option Explicit

Public Const GOAL         As Double = 20000000   ' 20 million weekly goal
Private Const TOTAL_COL   As Long = 12           ' column L  weekly total
Private Const SYMBOL_COL  As Long = 14           ' column N  star / X
Private Const DATA_FIRST  As Long = 3            ' column C
Private Const DATA_LAST   As Long = 7            ' column G
Private Const FIRST_ROW   As Long = 2            ' row 1 is the header
Private Const LAST_ROW    As Long = 105          ' last row to process (N2:N105)

' STAR is built at runtime via StarChar() (ChrW 9733) so the glyph
' survives .bas import regardless of file encoding.
Private Const CROSS       As String = "X"        ' red X

Private Function StarChar() As String
    StarChar = ChrW(9733)                        ' gold star glyph
End Function

' --- Fill / refresh every data row on the ACTIVE sheet ---------
Public Sub UpdateGoalSymbols()

    Dim ws       As Worksheet
    Dim r        As Long

    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    On Error GoTo CleanUp
    For r = FIRST_ROW To LAST_ROW
        UpdateGoalSymbolsRow ws, r
    Next r

CleanUp:
    Application.ScreenUpdating = True

End Sub

' --- Apply the rule to a single row ----------------------------
Public Sub UpdateGoalSymbolsRow(ByVal ws As Worksheet, ByVal r As Long)

    If r < FIRST_ROW Or r > LAST_ROW Then Exit Sub

    Dim cell As Range
    Set cell = ws.Cells(r, SYMBOL_COL)

    Dim totalVal As Variant
    totalVal = ws.Cells(r, TOTAL_COL).Value

    If Not IsError(totalVal) Then
        If IsNumeric(totalVal) Then
            If CDbl(totalVal) >= GOAL Then
                ' --- Goal hit: gold star ---
                cell.Value = StarChar()
                cell.Font.Color = RGB(255, 200, 0)   ' gold
                cell.Font.Bold = False
                cell.Font.Italic = False
                cell.HorizontalAlignment = xlCenter
                Exit Sub
            End If
        End If
    End If

    ' --- Under goal: red X only when the row has C:G data ---
    If HasRowData(ws, r) Then
        cell.Value = CROSS
        cell.Font.Color = RGB(192, 0, 0)         ' red
        cell.Font.Bold = True                    ' bold italic red X
        cell.Font.Italic = True
        cell.HorizontalAlignment = xlCenter
    Else
        cell.ClearContents
        cell.Font.ColorIndex = xlColorIndexAutomatic
        cell.Font.Bold = False
        cell.Font.Italic = False
        cell.HorizontalAlignment = xlGeneral
    End If

End Sub

' --- True if any cell in C:G of the row is non-empty -----------
' An "x" marks an excused day, so it is treated like an empty cell: a row of
' nothing but excused days is not a missed goal.
Private Function HasRowData(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long
    Dim v As Variant
    For c = DATA_FIRST To DATA_LAST
        v = ws.Cells(r, c).Value
        If IsError(v) Then
            HasRowData = True
            Exit Function
        ElseIf Trim(CStr(v)) <> "" And LCase(Trim(CStr(v))) <> "x" Then
            HasRowData = True
            Exit Function
        End If
    Next c
    HasRowData = False
End Function

' ============================================================
' OPTIONAL  AUTOMATIC UPDATING
' Paste the following into the "ThisWorkbook" object so column N
' refreshes the moment column C:G or M changes on any sheet.
' (Remove the leading apostrophes when you paste it there.)
' ------------------------------------------------------------
' Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'     Dim cell As Range
'     Dim touched As Range
'
'     ' Only react to changes in the C:G data block or the L total
'     On Error Resume Next
'     Set touched = Application.Intersect(Target, Sh.Range("C:G,L:L"))
'     On Error GoTo 0
'     If touched Is Nothing Then Exit Sub
'
'     Application.EnableEvents = False
'     On Error GoTo CleanUp
'     For Each cell In touched.Cells
'         If cell.Row >= 2 Then GoalSymbols.UpdateGoalSymbolsRow Sh, cell.Row
'     Next cell
' CleanUp:
'     Application.EnableEvents = True
' End Sub
' ============================================================
