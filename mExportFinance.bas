Attribute VB_Name = "mExportFinance"
Sub Export_New_Items()

OpenExportWorkbook
CopyExportsheets
CreateExportTable
Count_Difference_In_Rows
DeleteExportSheet
CloseExportWorkbook
    
End Sub
Sub OpenExportWorkbook()
'Open a Workbook

Workbooks.Open "D:\_Files\BUDGET-BOOK\.data\Export2024.csv"

'NOTE
'THIS WILL NEED CHANGED BASED ON LOCATION

End Sub
Sub CopyExportsheets()

Dim exportWorkbook As Workbook
Dim financeWorkbook As Workbook

'NOTE
'THESE WILL NEED TO CHANGE IF NECCESSARY FILES ARE RENAMED
Set exportWorkbook = Workbooks("Export2024.csv")
Set financeWorkbook = Workbooks("Transactions2024.xlsm")

Dim exportWorksheet As Worksheet
Set exportWorksheet = exportWorkbook.Worksheets("Export2024")

exportWorksheet.Copy After:=financeWorkbook.Sheets(1)

End Sub
Sub ClearExportTable()
Dim ws As Worksheet
Set ws = ActiveSheet

Dim tbl As ListObject
Set tbl = ws.ListObjects("Summary")

    'Definf Sheet and table name
     With Sheets("MAIN").ListObjects("Summary")
        'Check If any data exists in the table
        If Not .DataBodyRange Is Nothing Then
            'Clear Content from the table
            .DataBodyRange.ClearContents
        End If
    End With
Clear_Rows

End Sub
Sub Clear_Rows()

Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects("Summary")

'Dim rowCount As Integer
'Set rowCount = tbl.DataBodyRange.Rows.Count

For i = 1 To tbl.DataBodyRange.Rows.Count
    RemoveRowToTable
Next i


Set rowCount = Nothing

End Sub

Sub CreateExportTable()

'
' CreateExportTable2 Macro
'
'
Dim rw As Integer
rw = (Range("A1").End(xlDown).Row)


ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(4, 1), Cells(rw, 7)), , xlYes).Name = _
    "Export"
  
End Sub
Sub Count_Difference_In_Rows()

'NOTE
'THIS WILL NEED TO CHANGE IF MAIN FILE IS RENAMED
Dim wb As Workbook
Set wb = Workbooks("Transactions2024.xlsm")

Dim exportWorksheet As Worksheet
Dim financeWorksheet As Worksheet
Set exportWorksheet = wb.Worksheets("Export2024")
Set financeWorksheet = wb.Worksheets("MAIN")

Dim newRows As Integer
newRows = (exportWorksheet.ListObjects("Export").ListRows.Count) - (financeWorksheet.ListObjects("Summary").ListRows.Count)
    'Create New Rows
For i = 1 To newRows
    financeWorksheet.ListObjects("Summary").ListRows.Add 1
Next i

Dim srcRow As Range
Dim destRow As Range

For i = 1 To newRows
    Set srcRow = exportWorksheet.ListObjects("Export").ListRows(i).Range
    Set destRow = financeWorksheet.ListObjects("Summary").ListRows(i).Range
    
    srcRow.Copy
    destRow.PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    'exportRow.Rows(i).Copy
    'financeRows.Rows(i).Add
Next i

End Sub

Sub DeleteExportSheet()
'Modify Here
Application.DisplayAlerts = False
Sheets("Export2024").Delete
Application.DisplayAlerts = True

End Sub
Sub CloseExportWorkbook()
'Close a Workbook
'Modify Here
Workbooks("Export2024.csv").Close SaveChanges:=True
End Sub


Sub RemoveRowToTable()

Dim ws As Worksheet
Set ws = ActiveSheet

Dim tbl As ListObject
Set tbl = ws.ListObjects("Summary")

tbl.ListRows(1).Delete

End Sub

Sub AddRowToTable()

Dim ws As Worksheet
Set ws = ActiveSheet

Dim tbl As ListObject
Set tbl = ws.ListObjects("Summary")

'add a row at the end of the table
'tbl.ListRows.Add
'add a row as the fifth row of the table (counts the headers as a row)
tbl.ListRows.Add 1

End Sub

Sub Copy_Paste_Below_Last_Cell()
'Find the last used row in both sheets and copy and paste data below existing data.

Dim wsCopy As Worksheet
Dim wsDest As Worksheet
Dim lCopyLastRow As Long
Dim lDestLastRow As Long

    'Set Variables for copy and destination sheets
    Set wsCopy = Workbooks("Export.csv").Worksheets("Export")
    Set wsDest = Workbooks("FINANCE.xlsm").Worksheets("MAIN")
    
    '1. Find last used row in the copy range based on data in column A
    lCopyLastRow = ws.Copy.Cells(wsCopy.Rows.Count, "A").End(xlUp).Row
    
    '2. Find first blank row in the destination range based on data in column A
    'Offset property moves down 1 row
    lDestLastRow = wDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Offset(1).Row
    
    '3. Copy & Paste Data
    wsCopy.Range("A2:D" & lCopyLastRow).Copy _
        wsDest.Range("A" & lDestLastRow)
End Sub


Sub Find_Next_Empty_Row()

Dim Rng As Range

On Error Resume Next
Set Rng = Range("Summary[[Transaction Number]]").SpecialCells(xlCellTypeBlanks)
On Error GoTo 0
If Not Rng Is Nothing Then
    Rng.Select
End If

End Sub

Sub Find_Next_Empty_Row2()

Dim wb As Workbook
Dim exportWorksheet As Worksheet
Dim Rng As Range
Dim rw As Long

Set wb = Workbooks("FINANCE.xlsm")
Set exportWorksheet = Worksheets("Export")

'MsgBox (Range("A1").End(xlDown).Row) - 4
rw = (Range("A1").End(xlDown).Row)
'ActiveCell.Value = exportWorksheet.Rows.Count
'Range("A" & Rows.Count)

End Sub

Sub Count_Rows()

Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects("Summary")
MsgBox tbl.Range.Rows.Count
MsgBox tbl.HeaderRowRange.Rows.Count
MsgBox tbl.DataBodyRange.Rows.Count
Set tbl = Nothing

End Sub







