Sub forEachWs2()

    ' This procedure can be used to clear the summary tables from the entire workbook
    ' It is useful when testing the other procedures and recreating the summary tables

    For Each ws In ActiveWorkbook.Worksheets
        ws.Range("I:Q").Columns.Delete
    Next

End Sub