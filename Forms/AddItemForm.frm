VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddItemForm 
   Caption         =   "Add Expense/Income"
   ClientHeight    =   6167
   ClientLeft      =   120
   ClientTop       =   460
   ClientWidth     =   7740
   OleObjectBlob   =   "Untitled.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    With AddItemForm.cboxType
        .AddItem "Income"
        .AddItem "Expense"
    End With
    

    With AddItemForm.cboxFrequency
        .AddItem "One-time"
        .AddItem "Recurring"
    End With

  
    With AddItemForm.cboxRecurrence
        .AddItem "Weekly"
        .AddItem "Monthly"
        .AddItem "Annually"
    End With


    AddItemForm.cboxCategory.Clear
    AddItemForm.Recurrence.Visible = False
    AddItemForm.cboxRecurrence.Visible = False
    AddItemForm.Instances.Visible = False
    AddItemForm.txtInstances.Visible = False
End Sub

Private Sub cboxFrequency_Change()
    
    If AddItemForm.cboxFrequency.Value = "Recurring" Then
        AddItemForm.Recurrence.Visible = True
        AddItemForm.cboxRecurrence.Visible = True
        AddItemForm.Instances.Visible = True
        AddItemForm.txtInstances.Visible = True
    Else
        AddItemForm.Recurrence.Visible = False
        AddItemForm.cboxRecurrence.Visible = False
        AddItemForm.Instances.Visible = False
        AddItemForm.txtInstances.Visible = False
    End If
End Sub

Private Sub cboxType_Change()

    AddItemForm.cboxCategory.Clear
    
   
    If AddItemForm.cboxType.Value = "Expense" Then
  
        With AddItemForm.cboxCategory
            .AddItem "Shopping"
            .AddItem "Bills"
            .AddItem "Entertainment"
            .AddItem "Food"
            .AddItem "Other"
        End With
    ElseIf AddItemForm.cboxType.Value = "Income" Then
        With AddItemForm.cboxCategory
            .AddItem "Salary"
            .AddItem "Bonus"
        End With
    End If
End Sub

Private Sub SubmitBtn_Click()

    Dim WB As Workbook
    Dim ws As Worksheet
    Dim goalsWs As Worksheet
    Dim initialDate As Date
    Dim intRow As Long
    Dim recurrencePeriod As String
    Dim Instances As Integer
    Dim nextDate As Date
    Dim i As Integer
    Dim Budget As Double
    Dim userResponse As VbMsgBoxResult

    ' Set workbook and worksheets
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Input")
    Set goalsWs = WB.Worksheets("Goals")
    intRow = 10

    ' Check if Item field is filled
    If txtItem.Value = "" Then
        MsgBox "Please enter an item", vbExclamation, "Missing Item"
        Exit Sub
    End If

    ' Validate the entered date
    If Not IsDateInputValid(txtDay.Value, txtMonth.Value, txtYear.Value) Then
        MsgBox "Please enter a valid date.", vbExclamation, "Invalid Date"
        Exit Sub
    End If

    ' Validate Type field
    If cboxType.Value = "" Then
        MsgBox "Please select a type", vbExclamation, "Missing Type"
        Exit Sub
    End If

    ' Validate Category field
    If cboxCategory.Value = "" Then
        MsgBox "Please select a category", vbExclamation, "Missing Category"
        Exit Sub
    End If

    ' Validate Price field
    If txtPrice.Value = "" Or Not IsNumeric(txtPrice.Value) Then
        MsgBox "Please enter a valid price.", vbExclamation, "Invalid Price"
        Exit Sub
    End If

    ' Check if expense exceeds budget
    Budget = goalsWs.Range("M16").Value
    If cboxType.Value = "Expense" And Abs(txtPrice.Value) > Budget Then
        userResponse = MsgBox("The expense you are trying to input exceeds your budget of $" & _
                              Format(Budget, "0.00") & ". This may put you into debt. Do you want to proceed?", _
                              vbExclamation + vbYesNo, "Budget Exceeded")
        If userResponse = vbNo Then Exit Sub
    End If

    ' Find the next empty row in the Input sheet
    Do While ws.Cells(intRow, "A").Value <> ""
        intRow = intRow + 1
    Loop

    ' Parse the initial date
    initialDate = DateSerial(txtYear.Value, txtMonth.Value, txtDay.Value)

    ' Add the initial transaction
    ws.Cells(intRow, "A").Value = initialDate
    ws.Cells(intRow, "A").NumberFormat = "yyyy-mm-dd;@"
    ws.Cells(intRow, "B").Value = cboxType.Value
    ws.Cells(intRow, "C").Value = txtItem.Value
    ws.Cells(intRow, "D").Value = cboxCategory.Value
    If cboxType.Value = "Income" Then
        ws.Cells(intRow, "E").Value = txtPrice.Value
        ws.Cells(intRow, "E").NumberFormat = "$#,##0.00"
    ElseIf cboxType.Value = "Expense" Then
        txtPrice.Value = txtPrice.Value * -1
        ws.Cells(intRow, "E").Value = txtPrice.Value
        ws.Cells(intRow, "E").NumberFormat = "$#,##0.00"
    End If

    ' Handle Recurring Transactions
    If cboxFrequency.Value = "Recurring" Then
        If cboxRecurrence.Value = "" Then
            MsgBox "Please select a recurrence type", vbExclamation, "Missing Recurrence"
            Exit Sub
        End If

        If Not IsNumeric(txtInstances.Value) Or CInt(txtInstances.Value) <= 0 Then
            MsgBox "Please enter a valid number of instances.", vbExclamation, "Invalid Instances"
            Exit Sub
        End If

        recurrencePeriod = cboxRecurrence.Value
        Instances = CInt(txtInstances.Value)
        nextDate = initialDate

        For i = 1 To Instances - 1
            intRow = intRow + 1

            ' Determine the next recurring date
            Select Case recurrencePeriod
                Case "Monthly"
                    nextDate = DateAdd("m", 1, nextDate)
                Case "Weekly"
                    nextDate = nextDate + 7
                Case "Annually"
                    nextDate = DateAdd("yyyy", 1, nextDate)
            End Select

            ' Add the recurring transaction
            ws.Cells(intRow, "A").Value = nextDate
            ws.Cells(intRow, "A").NumberFormat = "yyyy-mm-dd;@"
            ws.Cells(intRow, "B").Value = cboxType.Value
            ws.Cells(intRow, "C").Value = txtItem.Value
            ws.Cells(intRow, "D").Value = cboxCategory.Value
            ws.Cells(intRow, "E").Value = txtPrice.Value
            ws.Cells(intRow, "E").NumberFormat = "$#,##0.00"
        Next i
    End If

    MsgBox "Transaction(s) added successfully!", vbInformation, "Success"
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

