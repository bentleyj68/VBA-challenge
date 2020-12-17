Sub forEachWs()

    ' This is the first and only procedure called by the user in the spreadsheet
    ' It can be invoked by including a button on one of the worksheets or run manually    

    ' Go through each worksheet (each year) and display both summary tables
    For Each ws In ActiveWorkbook.Worksheets
        ws.Select
        Call Stock_Summary
        Call Greatest_Summary
        Range("A1").Select
    Next

End Sub
