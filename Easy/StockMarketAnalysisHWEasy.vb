Sub TickerTally()

'Declare miscellany items

Dim StockName As String
Dim StockTableRow As Integer

'Build ForEach to gather parameters for worksheet

For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Perform ForNext loop to gather total stock volume and collect individual names

    For i = 2 To lastrow

        StockTableRow = 2


        If Cells(i, 1) <> Cells(i + 1, 1) Then
            StockName = ws.Cells(i, 1).Value
            StockTotal = StockTotal + ws.Cells(i, 7).Value

            'Create table of total volumes and corresponding ticker names

            StockTableRow = StockTableRow + 1


            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Total Stock Value"
            ws.Range("I" & StockTableRow).Value = StockName
            ws.Range("J" & StockTableRow).Value = StockTotal
        

        End If



    Next i
    
    'Reset Totals
    StockTotal = 0
    StockTableRow = 2


Next ws


End Sub