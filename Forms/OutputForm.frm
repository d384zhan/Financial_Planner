VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputForm 
   Caption         =   "Output"
   ClientHeight    =   5110
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   4040
   OleObjectBlob   =   "OutputForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OutputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Output_Click()
    Dim WB As Workbook
    Dim ws As Worksheet
    Dim wsData As Worksheet
    Dim startDate As Date
    Dim endDate As Date
    Dim intReadRow As Integer
    Dim intWriteRow1 As Integer
    Dim intWriteRow2 As Integer
    Dim transactionDate As Date
    Dim dataEmpty As Boolean

    ' Set worksheets
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Output")
    Set wsData = WB.Worksheets("Data")

    intReadRow = 2
    intWriteRow1 = 47
    intWriteRow2 = 47

    ' Check if Data sheet is empty from row 2 onward
    dataEmpty = wsData.Cells(2, "A").Value = ""

    ' Validate start date
    If Not IsDateInputValid(txtDay1.Value, txtMonth1.Value, txtYear1.Value) Then
        MsgBox "Please enter a valid start date.", vbExclamation, "Invalid Start Date"
        Exit Sub
    End If

    ' Validate end date
    If Not IsDateInputValid(txtDay2.Value, txtMonth2.Value, txtYear2.Value) Then
        MsgBox "Please enter a valid end date.", vbExclamation, "Invalid End Date"
        Exit Sub
    End If

    ' Convert start and end dates to actual dates
    startDate = DateSerial(txtYear1.Value, txtMonth1.Value, txtDay1.Value)
    endDate = DateSerial(txtYear2.Value, txtMonth2.Value, txtDay2.Value)

    ' Ensure start date is earlier than or equal to end date
    If startDate > endDate Then
        MsgBox "Start date must be earlier than or equal to end date.", vbExclamation, "Invalid Date Range"
        Exit Sub
    End If

    ' Set dates in the worksheet for reference
    ws.Cells(6, "G").Value = startDate
    ws.Cells(6, "I").Value = endDate
    ws.Cells(6, "G").NumberFormat = "yyyy-mm-dd;@"
    ws.Cells(6, "I").NumberFormat = "yyyy-mm-dd;@"

    ' Loop through the "Data" sheet to filter and output transactions
    Do While wsData.Cells(intReadRow, "A").Value <> ""
        transactionDate = wsData.Cells(intReadRow, "A").Value

        If transactionDate >= startDate And transactionDate <= endDate Then
            If wsData.Cells(intReadRow, "B").Value = "Income" Then
                ws.Cells(intWriteRow1, "A").Value = transactionDate
                ws.Cells(intWriteRow1, "B").Value = wsData.Cells(intReadRow, "C").Value
                ws.Cells(intWriteRow1, "C").Value = wsData.Cells(intReadRow, "D").Value
                ws.Cells(intWriteRow1, "D").Value = wsData.Cells(intReadRow, "E").Value

                ws.Cells(intWriteRow1, "A").NumberFormat = "yyyy-mm-dd;@"
                ws.Cells(intWriteRow1, "D").NumberFormat = "$#,##0.00"

                intWriteRow1 = intWriteRow1 + 1
            ElseIf wsData.Cells(intReadRow, "B").Value = "Expense" Then
                ws.Cells(intWriteRow2, "G").Value = transactionDate
                ws.Cells(intWriteRow2, "H").Value = wsData.Cells(intReadRow, "C").Value
                ws.Cells(intWriteRow2, "I").Value = wsData.Cells(intReadRow, "D").Value
                ws.Cells(intWriteRow2, "J").Value = wsData.Cells(intReadRow, "E").Value

                ws.Cells(intWriteRow2, "G").NumberFormat = "yyyy-mm-dd;@"
                ws.Cells(intWriteRow2, "J").NumberFormat = "$#,##0.00"

                intWriteRow2 = intWriteRow2 + 1
            End If
        End If

        intReadRow = intReadRow + 1
    Loop

    MsgBox "Transactions outputted for selected period.", vbInformation, "Success"
    
    ' Call RefreshCharts only if data is not empty
    If Not dataEmpty Then
        Call RefreshCharts
    Else
        MsgBox "You must input some transactions first!"
    End If
    Unload Me
End Sub

Private Sub TodayBtn_Click()
    Dim WB As Workbook
    Dim ws As Worksheet
    Dim wsData As Worksheet
    Dim earliestDate As Date
    Dim todayDate As Date
    Dim intReadRow As Integer
    Dim intWriteRow1 As Integer
    Dim intWriteRow2 As Integer
    Dim transactionDate As Date
    Dim dataEmpty As Boolean

    ' Set worksheets
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Output")
    Set wsData = WB.Worksheets("Data")

    intReadRow = 2
    intWriteRow1 = 47
    intWriteRow2 = 47

    ' Check if Data sheet is empty from row 2 onward
    dataEmpty = wsData.Cells(2, "A").Value = ""

    ' Define the earliest and today's date
    earliestDate = DateSerial(1900, 1, 1)
    todayDate = Date

    ' Set dates in the worksheet for reference
    ws.Cells(6, "G").Value = earliestDate
    ws.Cells(6, "I").Value = todayDate
    ws.Cells(6, "G").NumberFormat = "yyyy-mm-dd;@"
    ws.Cells(6, "I").NumberFormat = "yyyy-mm-dd;@"

    ' Loop through the "Data" sheet to filter and output transactions
    Do While wsData.Cells(intReadRow, "A").Value <> ""
        transactionDate = wsData.Cells(intReadRow, "A").Value

        If transactionDate >= earliestDate And transactionDate <= todayDate Then
            If wsData.Cells(intReadRow, "B").Value = "Income" Then
                ws.Cells(intWriteRow1, "A").Value = transactionDate
                ws.Cells(intWriteRow1, "B").Value = wsData.Cells(intReadRow, "C").Value
                ws.Cells(intWriteRow1, "C").Value = wsData.Cells(intReadRow, "D").Value
                ws.Cells(intWriteRow1, "D").Value = wsData.Cells(intReadRow, "E").Value

                ws.Cells(intWriteRow1, "A").NumberFormat = "yyyy-mm-dd;@"
                ws.Cells(intWriteRow1, "D").NumberFormat = "$#,##0.00"

                intWriteRow1 = intWriteRow1 + 1
            ElseIf wsData.Cells(intReadRow, "B").Value = "Expense" Then
                ws.Cells(intWriteRow2, "G").Value = transactionDate
                ws.Cells(intWriteRow2, "H").Value = wsData.Cells(intReadRow, "C").Value
                ws.Cells(intWriteRow2, "I").Value = wsData.Cells(intReadRow, "D").Value
                ws.Cells(intWriteRow2, "J").Value = wsData.Cells(intReadRow, "E").Value

                ws.Cells(intWriteRow2, "G").NumberFormat = "yyyy-mm-dd;@"
                ws.Cells(intWriteRow2, "J").NumberFormat = "$#,##0.00"

                intWriteRow2 = intWriteRow2 + 1
            End If
        End If

        intReadRow = intReadRow + 1
    Loop

    MsgBox "Transactions up to today's date have been outputted.", vbInformation, "Success"
    
    ' Call RefreshCharts only if data is not empty
    If Not dataEmpty Then
        Call RefreshCharts
    Else
        MsgBox "You must input some transactions first!"
    End If
    Unload Me
End Sub
Private Function IsDateInputValid(day As Variant, month As Variant, year As Variant) As Boolean
    Dim maxDays As Integer

    ' Ensure day, month, and year are numeric
    If Not (IsNumeric(day) And IsNumeric(month) And IsNumeric(year)) Then
        IsDateInputValid = False
        Exit Function
    End If

    ' Convert to integers
    day = CInt(day)
    month = CInt(month)
    year = CInt(year)

    ' Validate year, month, and day ranges
    If year < 1900 Or month < 1 Or month > 12 Then
        IsDateInputValid = False
        Exit Function
    End If

    ' Determine the maximum valid days for the given month and year
    Select Case month
        Case 1, 3, 5, 7, 8, 10, 12
            maxDays = 31
        Case 4, 6, 9, 11
            maxDays = 30
        Case 2
            ' February: Check for leap year
            If (year Mod 4 = 0 And year Mod 100 <> 0) Or (year Mod 400 = 0) Then
                maxDays = 29
            Else
                maxDays = 28
            End If
    End Select

    ' Validate the day range
    If day < 1 Or day > maxDays Then
        IsDateInputValid = False
        Exit Function
    End If

    ' If all checks pass, the date is valid
    IsDateInputValid = True
End Function

