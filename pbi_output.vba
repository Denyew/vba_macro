Public bifolder As String
Public bisheet As Worksheet
Public wb As Workbook
Public ws As Worksheet
Public wb4 As Workbook
Public ws4 As Worksheet
Public export As Workbook
Public startsheet As Worksheet
Public endsheet As Worksheet
Public deleteSheets As Collection
Public i As Integer

Sub BI_input_folder()

Set wb = ThisWorkbook

Set bisheet = ThisWorkbook.Sheets("PowerBI")

With Application.fileDialog(msoFileDialogFolderPicker)

    .Title = "Select a folder"
    .AllowMultiSelect = False
    If .Show = -1 Then
     bifolder = .SelectedItems(1)
     
    Else
    MsgBox "No folder selected"
    
    End If
    
End With

bisheet.Range("A4").Value = bifolder


End Sub
    
Sub clear_4cat()

'Sub to clear sheets
      
Set startsheet = ThisWorkbook.Sheets("4")
Set endsheet = ThisWorkbook.Sheets("7")
Set deleteSheets = New Collection

For Each ws In ThisWorkbook.Sheets
    If ws.Index > startsheet.Index And ws.Index < endsheet.Index Then
        deleteSheets.Add ws.Name
    End If
Next ws

Application.DisplayAlerts = False

    For i = 1 To deleteSheets.Count
        ThisWorkbook.Sheets(deleteSheets(i)).Delete
    Next i
    Application.DisplayAlerts = True
    
End Sub

Sub import_4cat()

Call clear_4cat

Dim filepath4 As String

Set wb = ThisWorkbook

filepath4 = ThisWorkbook.Sheets("PowerBi").Range("A4") & "\4.csv"

MsgBox (filepath4)

'Finds file named 4.csv

If filepath4 = "" Then

    MsgBox "File not found"
    
Else

    Set wb4 = Workbooks.Open(filepath4)
    
    Set ws4 = wb4.Sheets(1)
    
        ws4.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        
        wb4.Close False
        
    wb.Sheets(wb.Sheets.Count).Name = "4Cat"

End If

Dim ws4_imp As Worksheet

Set ws4_imp = wb.Sheets("4Cat")

Dim lastRow As Long

lastRow = ws4_imp.Cells(ws4_imp.Rows.Count, 1).End(xlUp).Row

With ws4_imp.Range("A1:E" & lastRow)
    .AutoFilter Field:=2, Criteria1:="="
End With

'Replacing missing values with a certain value as impact is minimal
            
    On Error Resume Next
    ws4_imp.Range("B2:B" & lastRow).SpecialCells(xlCellTypeBlanks).Value = "xxx"
    On Error GoTo 0
    
    ws4_imp.AutoFilterMode = False

Dim ws4_Pivot As Worksheet
Dim pivottablerange4 As Range
Dim pivotTable As pivotTable
Dim pivotcache As pivotcache

Set pivottablerange4 = ws4_imp.UsedRange
Set ws4_Pivot = ThisWorkbook.Sheets.Add
ws4_Pivot.Name = "4CatPivot"

Set pivotcache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivottablerange4)

Set pivotTable = pivotcache.CreatePivotTable(TableDestination:=ws4_Pivot.Cells(1, 1), TableName:="4Cat")

With pivotTable

    .PivotFields("Store").Orientation = xlRowField
    .PivotFields("Store").Position = 1
    
    .PivotFields("4 Cat").Orientation = xlRowField
    .PivotFields("4 Cat").Position = 2
    
    .PivotFields("Month").Orientation = xlColumnField
    .PivotFields("Month").Position = 1
    
    .PivotFields("Day").Orientation = xlColumnField
    .PivotFields("Day").Position = 2
    
    .PivotFields("Sum of Total").Orientation = xlDataField
    .PivotFields("Sum of Total").NumberFormat = "#,##0.00"
    
End With

For Each DataField In pivotTable.DataFields
        DataField.NumberFormat = "#,##0.00"
    Next DataField

pivotTable.PivotFields("Month").Subtotals(1) = False
ThisWorkbook.Sheets("4CatPivot").Move after:=startsheet
ThisWorkbook.Sheets("4Cat").Move before:=endsheet

ws4_Pivot.Activate
    
End Sub
