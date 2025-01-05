VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoalCompletion 
   Caption         =   "UserForm1"
   ClientHeight    =   3269
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   3680
   OleObjectBlob   =   "GoalCompletionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoalCompletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim ws As Worksheet, lastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Goals")
    cboxGoal.Clear

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' Populate combo box with goals
    For i = 10 To lastRow
        If ws.Cells(i, "A").Value <> "" Then cboxGoal.AddItem ws.Cells(i, "A").Value
    Next i

    ' Warn user if no goals are available
    If cboxGoal.ListCount = 0 Then
        MsgBox "No goals available. Please add goals first.", vbExclamation
    End If
End Sub

Private Sub SubmitBtn_Click()
    Dim ws As Worksheet, lastRow As Long, goalRow As Long
    Dim goalName As Variant, contribution As Double, goalAmount As Double
    Dim currentContribution As Double, completionPercentage As Double, remainingAmount As Double
    Dim i As Long, found As Boolean
    Dim netBalance As Double

    ' Calculate the net balance
    netBalance = GetNetBalance()

    Set ws = ThisWorkbook.Sheets("Goals")
    goalName = cboxGoal.Value
    contribution = Val(txtAmount.Value)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    found = False

    ' Check if a goal is selected
    If goalName = "" Then
        MsgBox "Please select a goal.", vbExclamation
        Exit Sub
    End If

    ' Check if contribution is valid
    If contribution <= 0 Then
        MsgBox "Please enter a valid contribution amount.", vbExclamation
        Exit Sub
    End If

    ' Check if contribution exceeds the net balance
    If contribution > netBalance Or contribution > ws.Cells(16, "M").Value Then
        MsgBox "The contribution exceeds the available net balance of " & Format(netBalance, "$#,##0.00") & " or your budget. Try contributing less.", vbExclamation
        Exit Sub
    End If

    ' Search for the goal in the worksheet
    For i = 10 To lastRow
        If CStr(ws.Cells(i, "A").Value) = CStr(goalName) Then
            goalRow = i
            found = True
            Exit For
        End If
    Next i

    ' Check if the goal was found
    If Not found Then
        MsgBox "Goal '" & goalName & "' not found.", vbExclamation
        Exit Sub
    End If

    ' Calculate goal progress
    goalAmount = ws.Cells(goalRow, "B").Value
    currentContribution = ws.Cells(goalRow, "D").Value ' Current contribution amount
    remainingAmount = goalAmount - currentContribution

    ' Prevent over-contribution
    If contribution > remainingAmount Then
        MsgBox "You only need " & Format(remainingAmount, "$#,##0.00") & " to complete this goal.", vbExclamation
        Exit Sub
    End If

    ' Update contribution
    currentContribution = currentContribution + contribution
    completionPercentage = (currentContribution / goalAmount) * 100

    ' Store the updated contribution amount and percentage
    ws.Cells(goalRow, "D").Value = currentContribution
    ws.Cells(goalRow, "D").NumberFormat = "$#,##0.00"
    ws.Cells(goalRow, "E").Value = completionPercentage / 100
    ws.Cells(goalRow, "E").NumberFormat = "0.00%"

    ' Handle goal completion
    If completionPercentage >= 100 Then
        ' Clear the contents of columns A to F for the completed goal
        ws.Range(ws.Cells(goalRow, "A"), ws.Cells(goalRow, "F")).ClearContents
        
        ' Shift rows below up
        If goalRow < lastRow Then
            ws.Range(ws.Cells(goalRow + 1, "A"), ws.Cells(lastRow, "F")).Cut Destination:=ws.Cells(goalRow, "A")
            ws.Range(ws.Cells(lastRow, "A"), ws.Cells(lastRow, "F")).ClearContents ' Clear the last row after shifting
        End If
        
        MsgBox "Congratulations! Goal '" & goalName & "' has been completed.", vbInformation
    Else
        MsgBox "You contributed " & Format(contribution, "$#,##0.00") & ". Progress: " & Format(completionPercentage, "0.00") & "%.", vbInformation
    End If

End Sub

Private Function GetNetBalance() As Double
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


