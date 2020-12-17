Sub forEachWs()

    For Each ws In ActiveWorkbook.Worksheets
        ws.Select
        Call Stock_Summary
        Call Greatest_Summary
    Next

End Sub

Sub Stock_Summary()

    ' Use an integer to store the curent row of the summary table
    Dim SummaryRow As Integer
    SummaryRow = 2

    ' Create the summary table column headings
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    ' Store the Stock summary values
    Dim Ticker As String
    Dim OpenValue As Double
    Dim TotalStockVol As Double
    
    ' Get the first opening stock value
    OpenValue = Cells(2, 3).Value
    TotalStockVol = 0
    
    ' Find the last row in the sheet
    Dim lastrow As Double
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row of the data table
    For i = 2 To lastrow
    
        'Add to Stock Volume Total
        TotalStockVol = TotalStockVol + Cells(i, 7).Value
        
        ' Check if the 'Ticker' code is changing
        If Cells(i, 1) <> Cells(i + 1, 1) Then 'MsgBox Cells(i, 1).Value
        
            ' Update the summary table
            Cells(SummaryRow, 9).Value = Cells(i, 1).Value
            Cells(SummaryRow, 10).Value = Cells(i, 6).Value - OpenValue
            If OpenValue <> 0 Then
                Cells(SummaryRow, 11).Value = Cells(SummaryRow, 10).Value / OpenValue 'Closing - Opening /100
            Else
                Cells(SummaryRow, 11).Value = 0
            End If
            Cells(SummaryRow, 12).Value = TotalStockVol
            
            ' Set conditional formatting on the yearly change column
            If Cells(SummaryRow, 10).Value < 0 Then
                Cells(SummaryRow, 10).Interior.ColorIndex = 3
            Else
                Cells(SummaryRow, 10).Interior.ColorIndex = 4
            End If
            
            ' Reset total stock volume, store the next opening stock value, goto the next row in the summary table
            TotalStockVol = 0
            OpenValue = Cells(i + 1, 3).Value
            SummaryRow = SummaryRow + 1
                
        End If
        
    Next i
    
    'Set the Percentage Change Column to a Percentage type, autofit table
    Range("K2" & ":K" & lastrow).NumberFormat = "0.00%"
    Range("I:L").Columns.AutoFit
   
End Sub

Sub Greatest_Summary()

    ' Create the Greatest summary table row and column headings
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
          
    ' Find the last row in the summary table
    Dim lastrowSmry As Integer
    lastrowSmry = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Set the values to the first record.
    Cells(2, 16).Value = Cells(2, 9).Value
    Cells(2, 17).Value = Cells(2, 11).Value
    Cells(3, 16).Value = Cells(2, 9).Value
    Cells(3, 17).Value = Cells(2, 11).Value
    Cells(4, 16).Value = Cells(2, 9).Value
    Cells(4, 17).Value = Cells(2, 12).Value
    
    ' Loop through each row of the data table and update if highest
    For i = 3 To lastrowSmry
        
        ' Check if this row will update the Greatest table
        If Cells(i, 11).Value > Cells(2, 17).Value Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
        End If
        If Cells(i, 11).Value < Cells(3, 17).Value Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
        End If
        If Cells(i, 12).Value > Cells(4, 17).Value Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
        End If
        
    Next i
    
    ' Autofit the table
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("O:Q").Columns.AutoFit

End Sub

Sub forEachWs2()

    For Each ws In ActiveWorkbook.Worksheets
        ws.Range("I:Q").Columns.Delete
    Next

End Sub
