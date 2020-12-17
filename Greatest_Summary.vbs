Sub Greatest_Summary()

    ' This procedure will create the second summary table in the worksheet - Greatest Summary
    ' Note: StockSummary Procedure must be run first

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
    
    ' Loop through each row of the data table and update if greater than previous
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
    
    ' Autofit the table, bold headings
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("P1:Q1").Font.Bold = True
    Range("O2:O5").Font.Bold = True
    Range("O2:O5").Font.Italic = True
    Range("O:Q").Columns.AutoFit

End Sub