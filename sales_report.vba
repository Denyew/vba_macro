Sub import()

  'Imports selected .txt file and automatically converts into table

    Dim fileDialog As fileDialog
    Dim filepath As String
    Dim ws As Worksheet
    Dim importsheet As Worksheet
    Dim mainsheet As Worksheet
    Dim monthToFilter As String
    Dim lastRow As Long
    Dim cell As Range
    Dim monthDict As Object
    Dim yearDict As Object
    Dim monthList As String
    Dim yearList As String

'To initiate fileDialog to choose file
    
    If IsEmpty(Sheets("Main").Range("A6").Value) = True Then

        MsgBox "Please choose a folder"
        
        Call Folder_path

    Else
        
        Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
        
            fileDialog.Title = "Choose .txt file"
            
            fileDialog.InitialFileName = Range("A6").Value & "\"
            
            fileDialog.Filters.Clear
            
            fileDialog.Filters.Add "Text Files", "*.txt"
            
                If fileDialog.Show = -1 Then
                
                filepath = fileDialog.SelectedItems(1)
                
                Range("A4").Value = filepath
                
                End If
        
    End If
       
'Delimits selected .txt file
    
    Set importsheet = Sheets("Import")

        importsheet.Cells.ClearContents

    With importsheet.QueryTables.Add(Connection:="TEXT;" & filepath, Destination:=importsheet.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = True
            .Refresh BackgroundQuery:=False
    End With
    
'Filters for values for cross month data
    
    Set mainsheet = Sheets("Main")
    Set monthDict = CreateObject("Scripting.Dictionary")
    Set yearDict = CreateObject("Scripting.Dictionary")
    
    lastRow = importsheet.Cells(importsheet.Rows.Count, "A").End(xlUp).Row
    
    For Each cell In importsheet.Range("B2:B" & lastRow)
    
        If IsDate(cell.Value) Then
            monthDict(Format(cell.Value, "mmmm")) = 1
            yearDict(Year(cell.Value)) = 1
        End If
    Next cell
    
    monthList = Join(monthDict.Keys, ",")
    yearList = Join(yearDict.Keys, ",")
    
        'Creates dropdown list for multiple months
        
            mainsheet.Range("B9:B10").ClearContents
    
            With mainsheet.Range("B9").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=monthList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
                
            With mainsheet.Range("B10").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=yearList
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
            
                'If more than 1 month and/or year, warn the user of cross month/year data

                If UBound(Split(monthList, ",")) > 0 Or UBound(Split(yearList, ",")) > 0 Then
                
                MsgBox "Warning : Data imported contains cross-month data"
                
                End If
                

End Sub

  Public importsheet As Worksheet
  Public exportsheet As Worksheet
        
Sub export()

'Filters for month and year criteria selected and exports based on selected criteria
  
    Dim mainsheet As Worksheet
    Dim lastRow As Long
    Dim lastRowexp As Long
    Dim monthToFilter As String
    Dim yearToFilter As String
    Dim cell As Range
    Dim filterRange As Range
    Dim copyRange As Range


Set importsheet = Sheets("Import")
Set exportsheet = Sheets("Export")
Set mainsheet = Sheets("Main")

    exportsheet.Cells.ClearContents
    
    'Declaration of values to filter
    monthToFilter = mainsheet.Range("B9").Value
    yearToFilter = mainsheet.Range("B10").Value
    
    'Insertion of helper columns for filtering
    lastRow = importsheet.Cells(importsheet.Rows.Count, "B").End(xlUp).Row
    
    importsheet.Columns("C:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    importsheet.Cells(1, 3).Value = "Month"
    importsheet.Cells(1, 4).Value = "Year"
    
    importsheet.Range("C2:C" & lastRow).Formula = "=TEXT(B2,""mmmm"")"
    importsheet.Range("D2:D" & lastRow).Formula = "=YEAR(B2)"
    importsheet.Range("D2:D" & lastRow).NumberFormat = "General"
    
    'Selects helper columns to filter
    Set filterRange = importsheet.Range("A1").CurrentRegion
    
        filterRange.AutoFilter Field:=3, Criteria1:=monthToFilter
        filterRange.AutoFilter Field:=4, Criteria1:=yearToFilter
             
                
    On Error Resume Next
    Set copyRange = filterRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
         
    If copyRange Is Nothing Then
    
    MsgBox "No visible cells filtered"
    
    Else
        copyRange.Copy Destination:=exportsheet.Range("A1")
    
    End If
    
    importsheet.AutoFilterMode = False
    exportsheet.AutoFilterMode = False
    importsheet.Columns("C:D").Delete
    exportsheet.Columns("C:D").Delete

    exportsheet.Range("B2:B" & lastRow).NumberFormat = "mm/dd/yyyy"
    

End Sub

Sub Append()

'Appends export to a selected file, that contains sales data of the month"

Dim fd As fileDialog
Dim folderD As fileDialog
Dim selectedFile As String
Dim selectedFolder As String
Dim initialFolder As String
Dim mainsheet As Worksheet
Dim folderAddress As Range
Dim sourceworkbook As Workbook
Dim sourceworkbookpath As String
sourceworkbookpath = "\Desktop\Sales Report Importer.xlsm"
Set sourceworkbook = Workbooks.Open(sourceworkbookpath)

    'Activates the control sheet

    sourceworkbook.Activate
    
    Set mainsheet = Sheets("Main")
    
    
    'Selects folder of target file

If mainsheet.Range("A15") = "" Then

    Set folderD = Application.fileDialog(msoFileDialogFolderPicker)
    
    folderD.Title = "Select a Folder"
    
    If folderD.Show = -1 Then
    
        selectedFolder = folderD.SelectedItems(1)
        
        Set folderAddress = mainsheet.Range("A15")
        
        folderAddress.Value = selectedFolder

    Else
        
        folderAddress = mainsheet.Range("A15").Value
    
    End If

End If
    
    If mainsheet.Range("A17") = "" Then

        Set fd = Application.fileDialog(msoFileDialogFilePicker)
        
        fd.InitialFileName = folderAddress & "\"
        
        fd.Title = "Select a File"
        
        fd.AllowMultiSelect = False
        
        fd.Filters.Clear
        
        fd.Filters.Add "All Files", "*.*"
        
            If fd.Show = -1 Then
            
                selectedFile = fd.SelectedItems(1)
                
                mainsheet.Range("A17").Value = Dir(selectedFile)
            
            End If
        
    Else
    
    mainsheet.Range("A17").Value = Dir(selectedFile)

    End If
    
MsgBox (Dir(selectedFile))

Dim targetworkbook As Workbook
Dim targetsheet As Worksheet
Dim lastRow As Long
Dim targetworkbookpath As String

    If selectedFile = "" Then
    
        targetworkbookpath = mainsheet.Range("A15").Value & "\" & mainsheet.Range("A17").Value & ".xlsx"
        
      Set targetworkbook = Workbooks.Open(targetworkbookpath)
    
    Else
    
        Set targetworkbook = Workbooks.Open(Dir(selectedFile))
    
    End If

Set targetsheet = targetworkbook.Sheets("Sales Input File")

lastRow = targetsheet.Cells(targetsheet.Rows.Count, "A").End(xlUp).Row

Dim sourcerange As Range
Dim targetrange As Range

Set exportsheet = sourceworkbook.Sheets("Export")

    Dim lastRowSource As Long
    Dim lastColSource As Long
    Dim lastRowTarget As Long

lastRowSource = exportsheet.Cells(exportsheet.Rows.Count, "A").End(xlUp).Row
lastColSource = exportsheet.Cells(2, exportsheet.Columns.Count).End(xlToLeft).Column

Set sourcerange = exportsheet.Range(exportsheet.Cells(2, 1), exportsheet.Cells(lastRowSource, lastColSource))

Set targetrange = targetsheet.Cells(lastRow + 1, 1)

sourcerange.Copy targetrange

End Sub
