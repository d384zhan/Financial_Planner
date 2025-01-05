Attribute VB_Name = "RefreshandOutput"
Sub RefreshCharts()


    Dim ws2 As Worksheet
    Dim pt1 As PivotTable
    Dim pt2 As PivotTable
    Dim pt3 As PivotTable
    Dim startDate As Date
    Dim endDate As Date
    

    Set ws1 = ThisWorkbook.Sheets("Data")
    Set ws2 = ThisWorkbook.Sheets("Output")
    Set pt1 = ws1.PivotTables("PivotTable1")
    Set pt2 = ws1.PivotTables("PivotTable2")
    Set pt3 = ws1.PivotTables("PivotTable3")
    

    startDate = (ws2.Range("G6").Value)
    endDate = (ws2.Range("I6").Value)
    
 
    With pt1.PivotFields("Date")
        .ClearAllFilters
        .PivotFilters.Add Type:=xlDateBetween, Value1:=startDate, Value2:=endDate
    End With
    

    pt1.RefreshTable
    pt2.RefreshTable
    pt3.RefreshTable

End Sub

Sub Output()

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Output")


    intReadRow = 2
    intWriteRow1 = 47
    intWriteRow2 = 47

    Do While (Worksheets("Data").Cells(intReadRow, "A") <> "")
    

                If (Worksheets("Data").Cells(intReadRow, "A") >= ws.Cells(6, "G") And Worksheets("Data").Cells(intReadRow, "A") <= ws.Cells(6, "I")) Then
                    
                    If (Worksheets("Data").Cells(intReadRow, "B").Value = "Income") Then
                    
              
                        ws.Cells(intWriteRow1, "A") = Worksheets("Data").Cells(intReadRow, "A")
                        ws.Cells(intWriteRow1, "B") = Worksheets("Data").Cells(intReadRow, "C")
                        ws.Cells(intWriteRow1, "C") = Worksheets("Data").Cells(intReadRow, "D")
                        ws.Cells(intWriteRow1, "D") = Worksheets("Data").Cells(intReadRow, "E")
            
            
                        ws.Cells(intWriteRow1, "A").NumberFormat = "yyyy-mm-dd;@"
                        ws.Cells(intWriteRow1, "D").NumberFormat = "$#,##0.00"
        
                        intWriteRow1 = intWriteRow1 + 1
                    
                    End If
                    
                    If (Worksheets("Data").Cells(intReadRow, "B").Value = "Expense") Then
                    

                        ws.Cells(intWriteRow2, "G") = Worksheets("Data").Cells(intReadRow, "A")
                        ws.Cells(intWriteRow2, "H") = Worksheets("Data").Cells(intReadRow, "C")
                        ws.Cells(intWriteRow2, "I") = Worksheets("Data").Cells(intReadRow, "D")
                        ws.Cells(intWriteRow2, "J") = Worksheets("Data").Cells(intReadRow, "E")
                  
                        ws.Cells(intWriteRow2, "G").NumberFormat = "yyyy-mm-dd;@"
                        ws.Cells(intWriteRow2, "J").NumberFormat = "$#,##0.00"
                    
                        intWriteRow2 = intWriteRow2 + 1
                        
                    End If
                
                End If
            
      
                intReadRow = intReadRow + 1
    
            Loop

End Sub
