VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveTransaction 
   Caption         =   "UserForm1"
   ClientHeight    =   5103
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   3180
   OleObjectBlob   =   "RemoveTransactionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Hide all combo boxes except cboxType
    cboxItem.Visible = False
    cboxInstances.Visible = False
    listInstances.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    
    ' Populate cboxType with "Income" and "Expense"
    cboxType.Clear
    cboxType.AddItem "Income"
    cboxType.AddItem "Expense"
End Sub

Private Sub cboxType_Change()
    Dim ws As Worksheet
    Dim intRow As Integer
    Dim itemSet As Collection
    Dim currentType As String

    Set ws = ThisWorkbook.Worksheets("Data")
    Set itemSet = New Collection
    currentType = cboxType.Value
    
    ' Show cboxItem and clear existing items
    cboxItem.Visible = True
    cboxItem.Clear
    Label3.Visible = True

    ' Loop through column B to find matching type and populate unique items in column C
    intRow = 2
    On Error Resume Next ' Prevent error for duplicate items in the collection
    Do While ws.Cells(intRow, "B").Value <> ""
        If ws.Cells(intRow, "B").Value = currentType Then
            itemSet.Add ws.Cells(intRow, "C").Value, CStr(ws.Cells(intRow, "C").Value)
        End If
        intRow = intRow + 1
    Loop
    On Error GoTo 0 ' Resume normal error handling

    ' Add unique items to cboxItem
    For Each Item In itemSet
        cboxItem.AddItem Item
    Next Item
End Sub

Private Sub cboxItem_Change()
    ' Make cboxInstances and listInstances visible
    cboxInstances.Visible = True
    listInstances.Visible = True
    Label4.Visible = True

    ' Populate cboxInstances with options
    cboxInstances.Clear
    cboxInstances.AddItem "One"
    cboxInstances.AddItem "Multiple"
    
    ' Populate listInstances with occurrences of the selected item and type
    UpdateListInstances
End Sub

Private Sub cboxInstances_Change()
    ' Update listInstances to allow or disallow multi-select based on cboxInstances selection
    If cboxInstances.Value = "One" Then
        listInstances.MultiSelect = fmMultiSelectSingle
    ElseIf cboxInstances.Value = "Multiple" Then
        listInstances.MultiSelect = fmMultiSelectMulti
    End If
End Sub

Private Sub UpdateListInstances()
    Dim wsData As Worksheet
    Dim intRow As Integer
    Dim currentType As String
    Dim currentItem As String

    Set wsData = ThisWorkbook.Worksheets("Data")
    currentType = cboxType.Value
    currentItem = cboxItem.Value

    ' Clear the list box
    listInstances.Clear

    ' Populate the list box with all occurrences of the selected item and type
    intRow = 2
    Do While wsData.Cells(intRow, "A").Value <> ""
        If wsData.Cells(intRow, "B").Value = currentType And wsData.Cells(intRow, "C").Value = currentItem Then
            ' Add the date to the list in yyyy-mm-dd format
            listInstances.AddItem Format(wsData.Cells(intRow, "A").Value, "yyyy-mm-dd")
        End If
        intRow = intRow + 1
    Loop
End Sub
Private Sub SubmitBtn_Click()
    Dim wsData As Worksheet
    Dim wsOutput As Worksheet
    Dim intRow As Integer
    Dim currentType As String
    Dim currentItem As Variant ' Use Variant to handle both numbers and text
    Dim i As Integer
    Dim dateItem As Variant
    Dim deleted As Boolean
    Dim validDate As Boolean

    ' Set worksheets
    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsOutput = ThisWorkbook.Worksheets("Output")
    currentType = cboxType.Value
    currentItem = cboxItem.Value
    deleted = False

    ' Iterate through the selected items in listInstances
    For i = 0 To listInstances.ListCount - 1
        If listInstances.Selected(i) Then
            dateItem = listInstances.List(i)

            ' Validate that the selected item is a valid date
            validDate = IsDate(dateItem)
            If Not validDate Then
                MsgBox "Invalid date detected in the list. Please ensure all dates are valid.", vbExclamation, "Invalid Date"
                Exit Sub
            End If

            ' Format the date as yyyy-mm-dd
            dateItem = Format(dateItem, "yyyy-mm-dd")

            ' Remove the first matching instance from "Data"
            intRow = 2
            Do While wsData.Cells(intRow, "A").Value <> ""
                If Format(wsData.Cells(intRow, "A").Value, "yyyy-mm-dd") = dateItem And _
                   wsData.Cells(intRow, "B").Value = currentType And _
                   CStr(wsData.Cells(intRow, "C").Value) = currentItem Then
                    
                    ' Clear the row contents (columns A-E) and shift up
                    wsData.Range(wsData.Cells(intRow, "A"), wsData.Cells(intRow, "E")).Delete Shift:=xlUp
                    deleted = True
                    Exit Do ' Stop after removing one instance
                End If
                intRow = intRow + 1
            Loop

            ' Perform additional actions
            startDate = wsOutput.Range("G6").Value
            endDate = wsOutput.Range("I6").Value
            
            Call ClearOutput
            wsOutput.Cells(6, "G").Value = startDate
            wsOutput.Cells(6, "I").Value = endDate
            Call Output
        End If ' Close the If statement for listInstances.Selected(i)
    Next i ' Close the For loop
    
    dataEmpty = wsData.Cells(2, "A").Value = ""
    ' Refresh charts and pivot tables
    If Not dataEmpty Then
        Call RefreshCharts
    End If
    ' Feedback to the user
    If deleted Then
        MsgBox "Selected transactions removed successfully!"
    Else
        MsgBox "No transactions matched the criteria."
    End If

    Unload Me
End Sub


