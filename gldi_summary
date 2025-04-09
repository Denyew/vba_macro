Public sourceworkbook As Workbook
Public targetworkbook As Workbook
Public sourcesheet As Worksheet
Public GLsheet As Worksheet
Public wb As Workbook
Public mainsheet As Worksheet
Public ws As Worksheet
Public lastCol As Long
Public lastRow As Long
Public filepath As String
Public i As Long

Sub Select_File()

 filepath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Select an Excel File")

    If filepath <> "False" Then
        
    MsgBox (Dir(filepath) & " is being processed")
        
    Else
    
    MsgBox "No File Selected"
    
    End If
    
Call GL_modify
    
End Sub


Sub GL_modify()

Set wb = ThisWorkbook

Set mainsheet = wb.Sheets("Main")

Application.DisplayAlerts = False

For Each ws In wb.Sheets

    If ws.Name <> mainsheet.Name Then
    
    ws.Delete
    
    End If
    
Next ws

Application.DisplayAlerts = True

Set sourceworkbook = Workbooks.Open(filepath)

Set targetworkbook = ThisWorkbook

Set sourcesheet = sourceworkbook.Sheets(1)

    sourcesheet.Copy After:=targetworkbook.Sheets(sourceworkbook.Sheets.Count)
    
    Set GLsheet = targetworkbook.Sheets(targetworkbook.Sheets.Count)
    
    GLsheet.Name = "GLDI"
    
    sourceworkbook.Close False

lastCol = GLsheet.Cells(1, GLsheet.Columns.Count).End(xlToLeft).Column
lastRow = GLsheet.Cells(GLsheet.Rows.Count, 1).End(xlUp).Row


    GLsheet.Cells(1, lastCol + 1).Value = "Acc Desc"
    GLsheet.Cells(1, lastCol + 2).Value = "Total"
    
    
Dim coa_path As String
Dim coa As Workbook
Dim coa_acc As Worksheet

'Set is only for objects
coa_path = wb.Sheets("Main").Range("A5").Value
Set coa = Workbooks.Open(coa_path)
Set coa_acc = coa.Sheets(1)

GLsheet.Range(GLsheet.Cells(2, lastCol + 1), GLsheet.Cells(lastRow, lastCol + 1)).Formula = _
"=XLOOKUP(C2,'" & coa.Name & "'!A:A, '" & coa.Name & "'! B:B, ""Not Found"")"

Dim colDebit As Long
Dim colCredit As Long
Dim headerDebit As String
Dim headerCredit As String

    headerDebit = "Debit"
    headerCredit = "Credit"
    
        For i = 1 To lastCol
            If CleanTrim(GLsheet.Cells(1, i).Value) = headerDebit Then
                colDebit = i
            ElseIf CleanTrim(GLsheet.Cells(1, i).Value) = headerCredit Then
                colCredit = i
            End If
        Next i
      
If colDebit > 0 And colCredit > 0 Then

    GLsheet.Range(GLsheet.Cells(2, lastCol + 2), GLsheet.Cells(lastRow, lastCol + 2)).FormulaR1C1 = _
    "=RC" & colCredit & "-RC" & colDebit

Else

    MsgBox "One or more headers not found"

End If

coa.Close SaveChanges:=False

Call Create_Pivot

End Sub

'Function to trim header

Function CleanTrim(str As String) As String
    
    CleanTrim = Trim(Replace(Replace(str, Chr(160), ""), Chr(32), ""))
        
End Function


Sub Create_Pivot()

    Dim pivotws As Worksheet
    Dim pivotTable As pivotTable
    Dim pivotCache As pivotCache
    Dim datarange As Range
    Dim GLSheetPivot As Worksheet
    Dim wb1 As Workbook
    
Set wb1 = ThisWorkbook

Set GLSheetPivot = wb1.Sheets("GLDI")
    
Set datarange = GLSheetPivot.Range(GLSheetPivot.Cells(1, 1), GLSheetPivot.Cells(lastRow, lastCol + 2))

Set pivotws = ThisWorkbook.Sheets.Add

    pivotws.Name = "Pivot"
    
Set pivotCache = ThisWorkbook.PivotCaches.Create(xlDatabase, datarange)

Set pivotTable = pivotCache.CreatePivotTable(pivotws.Cells(1, 1), "Pivot")

    With pivotTable
        .PivotFields("Account").Orientation = xlRowField
        .PivotFields("Acc Desc").Orientation = xlRowField
        .PivotFields("Cost Center").Orientation = xlColumnField
        .AddDataField .PivotFields("Total"), "Sum of Total", xlSum
        .RowAxisLayout xlTabularRow
        .PivotFields("Account").Subtotals(1) = False
        .PivotFields("Account").PivotFilters.Add Type:=xlCaptionIsGreaterThan, Value1:="30000"
        

    End With
    
targetworkbook.Sheets("Pivot").Activate
Range("C3").Select
ActiveWindow.FreezePanes = True



End Sub
