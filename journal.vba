Public template_path As String
Public jnl_template As Workbook
Public jnl_sheet As Worksheet
Public journal_range As Range
Public journalName As String
Public journalNo As Integer
Public journalDesc As String
Public accDate As String
Public month As String
Public year As String
Public year_short As String
' period = e.g. P2-25
Public period As String
'period1 = e.g. Feb 2025
Public period1 As String
Public journalCreator As String
Public source_path As String
Public source_wb As Workbook
Public source_ws As Worksheet
Public source_pivot As Worksheet
Public lastPivotRow As Long, i As Long, nextJnlRow As Long


Sub initialize_control()

Set control = Workbooks.Open("\Desktop\Journal Processor.xlsm")

Set control_sheet = control.Sheets("Main")

journalNo = control_sheet.Range("B8").Value

journalDesc = control_sheet.Range("B9").Value

month = control_sheet.Range("B10").Value

year_short = control_sheet.Range("B11").Value

year = control_sheet.Range("B12").Value

period1 = control_sheet.Range("B13").Value

accDate = control_sheet.Range("B14").Value

journalCreator = control_sheet.Range("B15").Value
 
period = control_sheet.Range("C13").Value


End Sub

Sub initialize_template()

Dim ws_jnl As Worksheet
Dim lastRow_jnl As Long

template_path = "\Desktop\Manual Journal Template.xlsx"

Set jnl_template = Workbooks.Open(template_path)

Set jnl_sheet = jnl_template.Worksheets("Single Journal")

lastRow_jnl = jnl_sheet.Cells(jnl_sheet.Rows.Count, "C").End(xlUp).row

Set journal_range = jnl_sheet.Range("C18:V" & lastRow_jnl)

jnl_sheet.Range("V18:V" & lastRow_jnl).NumberFormat = "General"

journal_range.Range("E18:V" & lastRow_jnl).ClearContents

End Sub

Sub initialize_source()

Dim control_wb As Workbook
Dim control_ws As Worksheet
Dim export_path As String

Set control_wb = Workbooks.Open("\Desktop\Journal Processor.xlsm")

Set control_ws = control_wb.Sheets("Main")

export_path = control_ws.Range("B6").Value

Set source_wb = Workbooks.Open(export_path)

Set source_ws = source_wb.Sheets("Pivot")

source_ws.Copy after:=jnl_sheet

Set source_pivot = jnl_template.Sheets("Pivot")

lastPivotRow = source_pivot.Cells(source_pivot.Rows.Count, "A").End(xlUp).row

nextJnlRow = 18

End Sub

Sub Sales_jnl()

    Call initialize_control
    Call initialize_template
    Call initialize_source


Dim product As String, delivery As String
Dim store As Variant

delivery = ""

For i = 4 To lastPivotRow

    store = source_pivot.Cells(i, "A").Value
    product = source_pivot.Cells(i, "B").Value
    
    
    
        If product <> delivery Then
        
            With jnl_sheet
            
                'fixed value
                .Cells(nextJnlRow, "E").Value = ""
                .Cells(nextJnlRow, "J").Value = ""
                .Cells(nextJnlRow, "K").Value = ""
                .Cells(nextJnlRow, "L").Value = ""
                .Cells(nextJnlRow, "M").Value = ""
                .Cells(nextJnlRow, "N").Value = ""
                
                .Cells(nextJnlRow, "F").Value = ""
                .Cells(nextJnlRow, "G").Value = store
                .Cells(nextJnlRow, "H").Value = ""
                .Cells(nextJnlRow, "I").Value = product
                .Cells(nextJnlRow, "O").Formula = "=ROUND(XLOOKUP(1,(Pivot!$A:$A = G" & nextJnlRow & ")*(Pivot!$B:$B=I" & nextJnlRow & "),Pivot!$E:$E,"""",0),2)"
                .Cells(nextJnlRow, "V").Formula = "="" & RIGHT($G" & nextJnlRow & ",3) & "" " & period & """"
                                                
            End With
            
        nextJnlRow = nextJnlRow + 1
        
    Dim lso_store As String
    
    If store = 1 Then
        new_store = 001
        Else
        new_store = store
    End If
        
            With jnl_sheet
            
                'fixed value
                .Cells(nextJnlRow, "E").Value = ""
                .Cells(nextJnlRow, "J").Value = ""
                .Cells(nextJnlRow, "K").Value = ""
                .Cells(nextJnlRow, "L").Value = ""
                .Cells(nextJnlRow, "M").Value = ""
                .Cells(nextJnlRow, "N").Value = ""
                
                .Cells(nextJnlRow, "F").Value = ""
                .Cells(nextJnlRow, "G").Value = new_store
                .Cells(nextJnlRow, "H").Value = ""
                .Cells(nextJnlRow, "I").Value = product
                .Cells(nextJnlRow, "P").Formula = "=O" & (nextJnlRow - 1)
                .Cells(nextJnlRow, "V").Formula = "="" & RIGHT($G" & nextJnlRow & ",3) & "" " & period & """"
                
            End With
            
        nextJnlRow = nextJnlRow + 1
        
        End If
        
    Next i
     
'Populating journal details

jnl_sheet.Range("D8").Value = "abc" & month & " " & year_short & "/1" & " " & journalCreator
jnl_sheet.Range("D9").Value = "abc"
jnl_sheet.Range("D11").Value = accDate

'Saving journal

Dim filename As String
Dim savepath As String

filename = "abc " & month & "-" & year_short & " - 19 " & journalCreator & " - " & "abc" & ".xlsx"

savepath = "\Desktop\" & filename

jnl_template.SaveAs filename:=savepath, FileFormat:=xlOpenXMLWorkbook

MsgBox "Sales Journal has been saved", vbInformation

End Sub
