Sub YearlyAssessment()

' Declare placeholder items and their values
Dim StockName As String
Dim StockTotal As Double
Dim StockTableRow As Integer
Dim opVal As Double
Dim clVal As Double
Dim opVal_row As Integer
Dim YrChange As Double
Dim PctChange As Double


StockTableRow = 2

opVal_row = 2


' Build For Each Loop to go though worksheets

For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' Build For Loop to check and collect year data in cells

    For i = 2 To lastrow

        If Cells(i, 1) <> Cells(i + 1, 1) Then
            StockName = ws.Cells(i, 1).Value
            StockTotal = StockTotal + ws.Cells(i, 7).Value

            opVal = ws.Cells(opVal_row, 3).Value
            clVal = ws.Cells(i, 6).Value
            YrChange = clVal - opVal

            If opVal = 0 Then
                PctChange = 0
            Else
                PctChange = YrChange / opVal
            End If

            ' Build Table and populate

            ws.Range("I1").Value = "Ticker"
            ws.Range("L1").Value = "Total Stock Value"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("I" & StockTableRow).Value = StockName
            ws.Range("L" & StockTableRow).Value = StockTotal
            ws.Range("J" & StockTableRow).Value = YrChange
            ws.Range("K" & StockTableRow).Value = PctChange
            ws.Range("K" & StockTableRow).NumberFormat = "0.00%"

            If ws.Range("J" & StockTableRow).Value > 0 Then
                ws.Range("J" & StockTableRow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & StockTableRow).Interior.ColorIndex = 3
            End If

            ' Reset Values

            StockTableRow = StockTableRow + 1        

        End If

    Next i

    StockTotal = 0
    StockTableRow = 2
    opVal_Row = 2

Next ws

End Sub
