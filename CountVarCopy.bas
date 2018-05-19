Attribute VB_Name = "CountVarCopy"
Sub Count()

Dim Rowcount As Integer     'to count how many empty rows before the next value
Dim RefNumber As Variant    'to capture the Seller Agreement Number
Dim RowNumber As Integer    'the actual number of the row where the RefNumber is in the worksheet
Dim RefType As Variant      'the type of seller agreement
Dim i_elem As Integer       'the ith element of the array
Dim j_elem As Integer       'the jth element of the array
Dim TotRows                 'the total number of rows to cover

Dim ListArray(1 To 1000000, 1 To 4) As Variant

RowNumber = 0

i_elem = 0

TotRows = InputBox("Enter the last row number")

For i = 1 To TotRows

Rowcount = 0

j_elem = 1

If ActiveSheet.Cells(i, 1) = "Seller Agreement Number:" Then
    
    i_elem = i_elem + 1

    RefNumber = ActiveSheet.Cells(i, 4)
    
    ListArray(i_elem, j_elem) = RefNumber
    
    j_elem = j_elem + 1
    
    RowNumber = i
    
    ListArray(i_elem, j_elem) = RowNumber

    While ActiveSheet.Cells(i + 1, 4) = ""

        Rowcount = Rowcount + 1
        
        i = i + 1
        
    Wend
    
    j_elem = j_elem + 1
    
    RefType = ActiveSheet.Cells(i + 1, 4)
    
    ListArray(i_elem, j_elem) = RefType
    
    j_elem = j_elem + 1
    
    ListArray(i_elem, j_elem) = Rowcount
    
    j_elem = 1
    

End If

Next

ActiveWorkbook.Sheets.Add

ActiveSheet.Cells(1, 1) = "Reference Number"
ActiveSheet.Cells(1, 2) = "Row Number"
ActiveSheet.Cells(1, 3) = "Reference Type"
ActiveSheet.Cells(1, 4) = "Row Count"

For K = 1 To i_elem

    For l = 1 To 4
    
    ActiveSheet.Cells(K + 1, l) = ListArray(K, l)
    
Next

Next

End Sub

