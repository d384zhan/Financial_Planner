Attribute VB_Name = "Budget"
Function GetNetBalance() As Double
    Dim ws As Worksheet
    Dim todayDate As Date
    Dim netBalance As Double
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet to "Data"
    Set ws = ThisWorkbook.Sheets("Data")

    ' Get today's date
    todayDate = Date

    ' Initialize net balance
    netBalance = 0

    ' Find the last row in column A (dates)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Loop through the rows to sum values in column E up to today's date
    For i = 1 To lastRow
        If IsDate(ws.Cells(i, "A").Value) And ws.Cells(i, "A").Value <= todayDate Then
            netBalance = netBalance + ws.Cells(i, "E").Value
        End If
    Next i

    ' Return the net balance
    GetNetBalance = netBalance
End Function

Function CalculateBudget(valuesRange As Range) As Double
    Dim netBalance As Double
    Dim expenses As Double
    Dim Budget As Double

    ' Get the net balance up to today
    netBalance = GetNetBalance()

    ' Calculate the sum of the provided range (expenses)
    expenses = Application.Sum(valuesRange)

    ' Calculate the budget
    Budget = netBalance - expenses

    ' Return the budget
    CalculateBudget = Budget
End Function

Sub RefreshBudget()
'
' RefreshBudget Macro
'

'
    Range("M16:XFC19").Select
    ActiveCell.Formula2R1C1 = "=CalculateBudget(R[-6]C[-9]:R[1048560]C[-9])"
    Range("M20").Select
End Sub


