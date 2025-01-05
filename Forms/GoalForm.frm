VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GoalForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4361
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   4600
   OleObjectBlob   =   "GoalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GoalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Populate cboxPriority with options
    With cboxPriority
        .AddItem "Urgent"
        .AddItem "Casual"
        .AddItem "Long Term"
    End With
End Sub

Private Sub SubmitBtn_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowToInsert As Long
    Dim goal As String
    Dim amount As Double
    Dim dateValue As Date
    Dim priority As String
    Dim i As Long
    Dim inserted As Boolean

    Set ws = ThisWorkbook.Sheets("Goals") ' Use the "Goals" worksheet
    lastRow = Application.Max(10, ws.Cells(ws.Rows.Count, "A").End(xlUp).row) ' Determine the last row of data

    ' Validate the date input
    If Not IsDateInputValid(txtDay.Value, txtMonth.Value, txtYear.Value) Then
        MsgBox "Invalid date entered. Please correct it.", vbExclamation
        Exit Sub
    End If

    ' Validate goal and amount input
    If txtGoal.Value = "" Then
        MsgBox "Goal cannot be empty.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtAmount.Value) Then
        MsgBox "Amount must be a numeric value.", vbExclamation
        Exit Sub
    End If

    ' Prepare the values for insertion
    goal = txtGoal.Value
    amount = CDbl(txtAmount.Value)
    dateValue = DateSerial(txtYear.Value, txtMonth.Value, txtDay.Value)
    priority = cboxPriority.Value

    ' Find the correct row for insertion
    rowToInsert = 10
    inserted = False

    For i = 10 To lastRow
        ' Compare priorities: Urgent > Casual > Long Term
        If ws.Cells(i, 6).Value = "" Or PriorityOrder(ws.Cells(i, 6).Value) > PriorityOrder(priority) Then
            rowToInsert = i
            inserted = True
            Exit For
        ElseIf PriorityOrder(ws.Cells(i, 6).Value) = PriorityOrder(priority) Then
            ' Compare dates for the same priority
            If ws.Cells(i, 3).Value = "" Or CDate(ws.Cells(i, 3).Value) > dateValue Then
                rowToInsert = i
                inserted = True
                Exit For
            End If
        End If
    Next i

    If Not inserted Then
        rowToInsert = lastRow + 1
    End If

    ' Ensure the insertion row is valid
    If rowToInsert > lastRow + 1 Then rowToInsert = lastRow + 1

    ' Shift only columns A to F down
    If ws.Cells(rowToInsert, 1).Value <> "" Then
        ws.Range(ws.Cells(rowToInsert, 1), ws.Cells(rowToInsert, 6)).Insert Shift:=xlDown
    End If

    ' Write the data to the worksheet
    ws.Cells(rowToInsert, 1).Value = goal
    ws.Cells(rowToInsert, 2).Value = amount
    ws.Cells(rowToInsert, 3).Value = dateValue
    ws.Cells(rowToInsert, 6).Value = priority

    ' Format the inserted data
    ws.Cells(rowToInsert, 2).NumberFormat = "$#,##0.00" ' Format amount as currency
    ws.Cells(rowToInsert, 3).NumberFormat = "yyyy-mm-dd" ' Format date

    MsgBox "Goal added successfully!", vbInformation
End Sub

Function PriorityOrder(priority As String) As Integer
    Select Case priority
        Case "Urgent": PriorityOrder = 1
        Case "Casual": PriorityOrder = 2
        Case "Long Term": PriorityOrder = 3
        Case Else: PriorityOrder = 99 ' Default for empty or unknown priorities
    End Select
End Function



Private Function IsDateInputValid(day As Variant, month As Variant, year As Variant) As Boolean
    Dim maxDays As Integer

    ' Validate Inputs
    If Not IsNumeric(day) Or Not IsNumeric(month) Or Not IsNumeric(year) Then Exit Function

    ' Convert Inputs to Integers
    day = CInt(day)
    month = CInt(month)
    year = CInt(year)

    ' Validate Year and Month
    If year < 1900 Or month < 1 Or month > 12 Then Exit Function

    ' Determine Max Days for Month
    Select Case month
        Case 1, 3, 5, 7, 8, 10, 12: maxDays = 31
        Case 4, 6, 9, 11: maxDays = 30
        Case 2
            If (year Mod 4 = 0 And year Mod 100 <> 0) Or (year Mod 400 = 0) Then
                maxDays = 29
            Else
                maxDays = 28
            End If
    End Select

    ' Validate Day
    If day >= 1 And day <= maxDays Then IsDateInputValid = True
End Function

