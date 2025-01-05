Attribute VB_Name = "ClearandSend"
Sub ClearExpensesIncomes()

    'Select the range
    Range("A10:E10").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    'Clear the selected range
    Selection.ClearContents
    
End Sub

Sub ClearOutput()

    'Clear start date
    Range("G6").Select
    Selection.ClearContents
    
    'Clear end date
    Range("I6").Select
    Selection.ClearContents
    
    
    Range("A47:D47").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("G47:J47").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

End Sub
Sub SendToData()
    Dim wsInput As Worksheet
    Dim wsData As Worksheet
    Dim sourceRow As Long
    Dim insertRow As Long
    Dim transactionDate As Date

    ' Set worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsData = ThisWorkbook.Sheets("Data")

    ' Loop through rows in the Input sheet
    sourceRow = 10
    Do While wsInput.Cells(sourceRow, "A").Value <> ""
        ' Get the transaction date from Input
        transactionDate = wsInput.Cells(sourceRow, "A").Value

        ' Find the correct position in Data sheet for the transaction
        insertRow = FindInsertPosition(wsData, transactionDate)

        ' Insert the transaction at the correct position
        InsertTransaction wsData, wsInput, insertRow, sourceRow

        ' Move to the next row in Input
        sourceRow = sourceRow + 1
    Loop

    Set pt2 = wsData.PivotTables("PivotTable2")
    Set pt3 = wsData.PivotTables("PivotTable3")
    

    pt2.RefreshTable
    pt3.RefreshTable

    MsgBox "Data transferred and sorted successfully!", vbInformation, "Success"
End Sub

Private Function FindInsertPosition(wsData As Worksheet, transactionDate As Date) As Long
    Dim rowNum As Long
    rowNum = 2 ' Start at row 2, assuming headers are in row 1

    ' Find the first row where the date is less than the transaction date
    Do While wsData.Cells(rowNum, "A").Value <> ""
        If wsData.Cells(rowNum, "A").Value < transactionDate Then Exit Do
        rowNum = rowNum + 1
    Loop

    ' Return the row number for insertion
    FindInsertPosition = rowNum
End Function

Private Sub InsertTransaction(wsData As Worksheet, wsInput As Worksheet, insertRow As Long, sourceRow As Long)
    Dim lastRow As Long
    Dim col As Integer

    ' Find the last occupied row in Data
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row

    ' Shift rows manually to make space for the new transaction
    If lastRow >= insertRow Then
        For col = 1 To 5
            wsData.Range(wsData.Cells(insertRow, col), wsData.Cells(lastRow, col)).Offset(1, 0).Value = _
                wsData.Range(wsData.Cells(insertRow, col), wsData.Cells(lastRow, col)).Value
        Next col
    End If

    ' Copy the transaction from Input to Data
    wsData.Cells(insertRow, "A").Value = wsInput.Cells(sourceRow, "A").Value ' Date
    wsData.Cells(insertRow, "A").NumberFormat = "yyyy-mm-dd;@"
    wsData.Cells(insertRow, "B").Value = wsInput.Cells(sourceRow, "B").Value ' Type
    wsData.Cells(insertRow, "C").Value = wsInput.Cells(sourceRow, "C").Value ' Item
    wsData.Cells(insertRow, "D").Value = wsInput.Cells(sourceRow, "D").Value ' Category
    wsData.Cells(insertRow, "E").Value = wsInput.Cells(sourceRow, "E").Value ' Price
    wsData.Cells(insertRow, "E").NumberFormat = "$#,##0.00"
End Sub


