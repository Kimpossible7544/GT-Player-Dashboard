Attribute VB_Name = "ArchiveArenaPlayer"
' ============================================================
' ArchiveArenaPlayer — VBA Macro Module
' Moves the selected player row from "Arena Power" to
' "Arena Power - Archived Players"
'
' USAGE:
'   1. Import this .bas file into your workbook:
'      Developer > Visual Basic > File > Import File
'   2. On the "Arena Power" sheet, select any cell in the
'      player's row you want to archive
'   3. Click the "Archive Player" button (or run the macro
'      from the Macros dialog)
'
' BUTTON SETUP:
'   Insert > Shapes > pick a shape > right-click > Assign Macro
'   > select "ArchiveSelectedPlayer"
' ============================================================

Option Explicit

Private Const SRC_SHEET  As String = "Arena Power"
Private Const DEST_SHEET As String = "Arena Power - Archived Players"

Public Sub ArchiveSelectedPlayer()

    Dim wsSrc  As Worksheet
    Dim wsDest As Worksheet
    Dim srcRow As Long
    Dim destRow As Long
    Dim lastCol As Long
    Dim playerName As String

    ' --- Verify we are on the Arena Power sheet ---
    If ActiveSheet.Name <> SRC_SHEET Then
        MsgBox "Please run this from the """ & SRC_SHEET & """ sheet.", _
               vbExclamation, "Wrong Sheet"
        Exit Sub
    End If

    Set wsSrc = ActiveSheet
    srcRow = Selection.Row

    ' --- Don't allow archiving the header row ---
    If srcRow = 1 Then
        MsgBox "Select a player row, not the header.", _
               vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    ' --- Get the player name from column A ---
    playerName = Trim(CStr(wsSrc.Cells(srcRow, 1).Value))
    If playerName = "" Then
        MsgBox "No player name found in column A of the selected row.", _
               vbExclamation, "Empty Row"
        Exit Sub
    End If

    ' --- Confirm with the user ---
    If MsgBox("Archive """ & playerName & """ to " & DEST_SHEET & "?" & vbNewLine & _
              vbNewLine & "This will remove the row from " & SRC_SHEET & ".", _
              vbYesNo Or vbQuestion, "Confirm Archive") <> vbYes Then
        Exit Sub
    End If

    ' --- Create the archive sheet if it doesn't exist ---
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(DEST_SHEET)
    On Error GoTo 0

    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsDest.Name = DEST_SHEET

        ' Copy the header row from the source sheet
        lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy _
            Destination:=wsDest.Cells(1, 1)

        ' Match the source column widths
        Dim c As Long
        For c = 1 To lastCol
            wsDest.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
        Next c
    End If

    ' --- Find the next empty row on the archive sheet ---
    destRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1

    ' --- Determine the last used column for this row ---
    lastCol = wsSrc.Cells(srcRow, wsSrc.Columns.Count).End(xlToLeft).Column
    ' Also check header extent in case data row is shorter
    Dim headerLastCol As Long
    headerLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If headerLastCol > lastCol Then lastCol = headerLastCol

    ' --- Copy the full row (values + formatting) ---
    wsSrc.Range(wsSrc.Cells(srcRow, 1), wsSrc.Cells(srcRow, lastCol)).Copy _
        Destination:=wsDest.Cells(destRow, 1)

    ' --- Delete the row from the source sheet ---
    wsSrc.Rows(srcRow).Delete Shift:=xlUp

    MsgBox """" & playerName & """ archived successfully.", _
           vbInformation, "Done"

End Sub
