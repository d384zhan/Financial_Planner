VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinancialAdvice 
   Caption         =   "UserForm1"
   ClientHeight    =   3255
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   8100
   OleObjectBlob   =   "FinancialAdviceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinancialAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set worksheet to "Goals"
    Set ws = ThisWorkbook.Sheets("Goals")

    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Clear combo box
    cboxGoal.Clear

    ' Populate combo box with goals starting from row 10
    For i = 10 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            cboxGoal.AddItem ws.Cells(i, 1).Value
        Else
            Exit For
        End If
    Next i

   
End Sub

Private Sub SubmitBtn_Click()
    Dim wsGoals As Worksheet
    Dim wsData As Worksheet
    Dim selectedGoal As String
    Dim goalRow As Long
    Dim goalAmount As Double
    Dim goalDueDate As Date
    Dim amountContributed As Double
    Dim netBalance As Double
    Dim futureBudget As Double
    Dim remainingAmount As Double
    Dim daysRemaining As Long
    Dim dailySavings As Double
    Dim totalContributed As Double
    Dim Budget As Double
    Dim i As Long
    Dim found As Boolean

    Set wsGoals = ThisWorkbook.Sheets("Goals")
    Set wsData = ThisWorkbook.Sheets("Data")

    ' Check if a goal is selected
    If cboxGoal.Value = "" Then
        MsgBox "Please select a goal.", vbExclamation
        Exit Sub
    End If

    ' Find the selected goal starting from row 10
    selectedGoal = cboxGoal.Value
    found = False
    For i = 10 To wsGoals.Cells(wsGoals.Rows.Count, "A").End(xlUp).row
        If CStr(wsGoals.Cells(i, 1).Value) = CStr(selectedGoal) Then
            goalRow = i
            found = True
            Exit For
        End If
    Next i

    ' If goal is not found, show an error
    If Not found Then
        MsgBox "Goal '" & selectedGoal & "' not found in the 'Goals' sheet.", vbExclamation
        Exit Sub
    End If

    ' Retrieve goal details from "Goals" sheet
    goalAmount = wsGoals.Cells(goalRow, 2).Value
    goalDueDate = wsGoals.Cells(goalRow, 3).Value
    amountContributed = wsGoals.Cells(goalRow, 4).Value

    ' Calculate remaining amount for the goal
    remainingAmount = goalAmount - amountContributed

    ' Calculate net balance based on transactions before the due date
    netBalance = 0
    For i = 2 To wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
        Dim transactionDate As Date
        Dim transactionAmount As Double

        transactionDate = wsData.Cells(i, 1).Value
        transactionAmount = wsData.Cells(i, 5).Value ' Transaction amount

        If transactionDate <= goalDueDate Then
            netBalance = netBalance + transactionAmount
        End If
    Next i

    ' Calculate total contributed amount from column D of "Goals"
    totalContributed = WorksheetFunction.Sum(wsGoals.Range("D10:D" & wsGoals.Cells(wsGoals.Rows.Count, "D").End(xlUp).row))

    ' Retrieve the budget from cell M16
    Budget = wsGoals.Cells(16, "M").Value

    ' Calculate future budget
    futureBudget = netBalance - totalContributed

    ' Calculate days remaining until the goal's due date
    daysRemaining = DateDiff("d", Date, goalDueDate)

    ' Provide financial advice
    If daysRemaining < 0 Then
        ' Goal is overdue
        MsgBox "Financial Advice:" & vbNewLine & _
               "Your goal '" & selectedGoal & "' is overdue by " & Abs(daysRemaining) & " days!" & vbNewLine & _
               "Current budget: $" & Format(Budget, "0.00") & vbNewLine & _
               "You need to contribute an additional $" & Format(remainingAmount, "0.00") & " to reach your goal." & vbNewLine & _
               "Achieve this goal ASAP!.", vbExclamation
    ElseIf futureBudget >= remainingAmount Then
        ' Sufficient balance to meet the goal
        dailySavings = remainingAmount / daysRemaining
        MsgBox "Financial Advice:" & vbNewLine & _
               "You are on track to meet your goal '" & selectedGoal & "'!" & vbNewLine & _
               "Projected budget by due date: $" & Format(futureBudget, "0.00") & vbNewLine & _
               "Additional contribution needed: $" & Format(remainingAmount, "0.00") & vbNewLine & _
               "Time remaining: " & daysRemaining & " days" & vbNewLine & _
               "Suggested daily contributions to maintain consistency: $" & Format(dailySavings, "0.00"), vbInformation
    Else
        ' Insufficient balance to meet the goal
        MsgBox "Financial Advice:" & vbNewLine & _
               "You are not currently on track to meet your goal '" & selectedGoal & "'." & vbNewLine & _
               "Projected budget by due date: $" & Format(futureBudget, "0.00") & vbNewLine & _
               "Additional budget needed: $" & Format(remainingAmount - futureBudget, "0.00") & vbNewLine & _
               "Time remaining: " & daysRemaining & " days" & vbNewLine & _
               "Consider adjusting expenses or seeking additional income to close the gap.", vbExclamation
    End If
End Sub
Private Sub GenAdvice_Click()
    ' Declarations
    Dim ws As Worksheet
    Dim message As String
    Dim i As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim incomeUpToToday As Double, expensesUpToToday As Double
    Dim projectedIncome As Double, projectedExpenses As Double
    Dim netBalance As Double
    Dim emergencyFundsNeeded As Double
    Dim currentRow As Long
    Dim lastRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    If ws Is Nothing Then
        MsgBox "The sheet 'Data' does not exist.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Initialize sums
    incomeUpToToday = 0
    expensesUpToToday = 0
    projectedIncome = 0
    projectedExpenses = 0

    ' Define date ranges
    startDate = Date
    endDate = DateAdd("m", 6, startDate)

    ' Find the last row in the data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Loop through data and calculate sums
    For currentRow = 2 To lastRow
        Dim transDate As Date
        Dim transType As String
        Dim amount As Double

        transDate = ws.Cells(currentRow, "A").Value
        transType = LCase(ws.Cells(currentRow, "B").Value)
        amount = ws.Cells(currentRow, "E").Value

        If transDate <= Date Then
            If transType = "income" Then
                incomeUpToToday = incomeUpToToday + amount
            ElseIf transType = "expense" Then
                expensesUpToToday = expensesUpToToday + Abs(amount)
            End If
        ElseIf transDate > Date And transDate <= endDate Then
            If transType = "income" Then
                projectedIncome = projectedIncome + amount
            ElseIf transType = "expense" Then
                projectedExpenses = projectedExpenses + Abs(amount)
            End If
        End If
    Next currentRow

    ' Calculate emergency funds needed for the next 6 months
    emergencyFundsNeeded = projectedExpenses

    ' Calculate net balance
    netBalance = (incomeUpToToday - expensesUpToToday) + (projectedIncome - projectedExpenses)

    ' Analyze Income Categories from Pivot Table
    message = AnalyzePivotTable(ws, "M9", 0.5, "Income", _
        "Diversify your asset allocation for category: ", True)

    ' Analyze Expense Categories from Pivot Table
    message = message & AnalyzePivotTable(ws, "P5", 0.4, "Expense", _
        "Cut down expenses for category: ", False)

    ' Add financial forecast to the message
    message = message & vbCrLf & "6-Month Financial Forecast:" & vbCrLf
    If netBalance >= 0 Then
        message = message & "You have enough emergency funds in place. Budget at least $" & _
            Format(emergencyFundsNeeded, "0.00") & " to cover your projected expenses for the next 6 months." & vbCrLf
    Else
        message = message & "You need to either reduce your projected expenses or increase your income. " & _
            "You should budget at least $" & Format(emergencyFundsNeeded, "0.00") & " as emergency funds." & vbCrLf
    End If

    ' Output the final advice
    MsgBox message, vbInformation, "Financial Advice"
End Sub

' Function to analyze pivot table
Function AnalyzePivotTable(ws As Worksheet, startCell As String, threshold As Double, _
    tableType As String, adviceMessage As String, isIncome As Boolean) As String
    Dim pivotRange As Range
    Dim totalValue As Double
    Dim i As Long
    Dim categoryValue As Double
    Dim message As String
    Dim explanation As String

    On Error Resume Next
    Set pivotRange = ws.Range(startCell).CurrentRegion
    On Error GoTo 0

    If pivotRange Is Nothing Then
        AnalyzePivotTable = tableType & " pivot table not found starting at " & startCell & "." & vbCrLf
        Exit Function
    End If

    ' Exclude the last row (Grand Total)
    Dim lastDataRow As Long
    lastDataRow = pivotRange.Rows.Count - 1

    ' Calculate total value
    If isIncome Then
        totalValue = Application.WorksheetFunction.Sum(pivotRange.Cells(2, 2).Resize(lastDataRow - 1))
    Else
        totalValue = 0
        For i = 2 To lastDataRow
            totalValue = totalValue + Abs(pivotRange.Cells(i, 2).Value)
        Next i
    End If

    message = tableType & " Analysis:" & vbCrLf
    For i = 2 To lastDataRow ' Skip header and Grand Total
        If isIncome Then
            categoryValue = pivotRange.Cells(i, 2).Value
        Else
            categoryValue = Abs(pivotRange.Cells(i, 2).Value)
        End If

        If categoryValue > threshold * totalValue Then
            If isIncome Then
                explanation = "This category dominates your income, which poses a risk if the source changes unexpectedly."
            Else
                explanation = "This category exceeds 40% of total expenses, indicating a disproportionate allocation."
            End If
            message = message & "- " & adviceMessage & pivotRange.Cells(i, 1).Value & ". " & explanation & vbCrLf
        End If
    Next i

    ' If no categories exceed the threshold, provide a positive message
    If message = tableType & " Analysis:" & vbCrLf Then
        If isIncome Then
            message = message & "Your income distribution is diversified." & vbCrLf
        ElseIf tableType = "Expense" Then
            message = message & "Your expense distribution is OK." & vbCrLf
        End If
    End If

    message = message & vbCrLf
    AnalyzePivotTable = message
End Function




