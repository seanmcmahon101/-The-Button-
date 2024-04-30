Sub ImportAndFilterData()
    Dim fd As FileDialog
    Dim selectedFile As String
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim lastRow As Long, lastColumn As Long, targetRow As Long, i As Long
    Dim customerIDCol As Long, itemDescriptionCol As Long
    Dim isCustomerIDColFound As Boolean, isItemDescriptionColFound As Boolean
    
    ' Set the target workbook and sheet
    Set targetWorkbook = ThisWorkbook
    Set targetSheet = targetWorkbook.Sheets("UK Report")
    
    ' Directly clear the contents of columns A to R in the target sheet
    With targetSheet
        .Range("A:R").ClearContents
    End With
    

    ' Prompt the user to select the file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm; *.csv"
        .AllowMultiSelect = False
        If .Show = True Then
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "No file was selected.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Open the selected file
    Set sourceWorkbook = Workbooks.Open(selectedFile)
    Set sourceSheet = sourceWorkbook.Sheets(1)
    
    ' Remove the first row if it's a title row
    sourceSheet.Rows(1).Delete Shift:=xlUp
    
    ' Initialize flags to false
    isCustomerIDColFound = False
    isItemDescriptionColFound = False
    
    ' Refresh lastColumn after deleting the title row
    lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    
    ' Find the necessary columns
    For i = 1 To lastColumn
        If Trim(UCase(sourceSheet.Cells(1, i).Value)) = "CUSTOMERID" Then
            customerIDCol = i
            isCustomerIDColFound = True
        ElseIf Trim(UCase(sourceSheet.Cells(1, i).Value)) = "ITEMDESCRIPTION" Then
            itemDescriptionCol = i
            isItemDescriptionColFound = True
        End If
    Next i
    
    ' Check if both necessary columns were found
    If Not isCustomerIDColFound Or Not isItemDescriptionColFound Then
        Dim missingColumns As String
        If Not isCustomerIDColFound Then missingColumns = "CustomerID"
        If Not isItemDescriptionColFound Then
            If missingColumns <> "" Then missingColumns = missingColumns & " and "
            missingColumns = missingColumns & "ItemDescription"
        End If
        
        MsgBox "Critical columns missing: " & missingColumns, vbCritical
        sourceWorkbook.Close False
        Exit Sub
    End If
    
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, customerIDCol).End(xlUp).Row
    targetRow = 2 ' Assuming the first row of the target sheet is for headers
    
    ' Copy data from the source to the target, starting from the second row
    For i = 2 To lastRow
    If UCase(sourceSheet.Cells(i, customerIDCol).Value) <> "NPI" And _
       UCase(sourceSheet.Cells(i, customerIDCol).Value) <> "SALES" And _
       UCase(sourceSheet.Cells(i, customerIDCol).Value) <> "INTMAN" And _
       sourceSheet.Cells(i, itemDescriptionCol).Value <> "" Then
        ' Copy headers from sourceSheet to targetSheet
        sourceSheet.Range("A1:R1").Copy Destination:=targetSheet.Range("A1")

        ' Copy data from columns A to R
        sourceSheet.Range("A" & i & ":R" & i).Copy Destination:=targetSheet.Cells(targetRow, 1)
        
        targetRow = targetRow + 1
    End If
Next i

    ' Close the source workbook without saving
    sourceWorkbook.Close False
    MsgBox "The process is complete.", vbInformation, "Process Complete"
End Sub


Sub ShowMyUserForm()
    UserForm1.Show
End Sub



