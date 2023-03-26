Sub stockmarket():


'i j k as counter for loops
Dim i, j, k As Integer

'Populate the column labels
For Each ws In Worksheets
ws.Cells(1, 9).Value = ("Ticker")
ws.Cells(1, 10).Value = ("Yearly Change")
ws.Cells(1, 11) = ("Percent Change")
ws.Cells(1, 12) = ("Total Stock Volume")
ws.Cells(1, 15) = ("Ticker")
ws.Cells(1, 16) = ("Value")
ws.Cells(2, 14) = ("Greatest % Increase")
ws.Cells(3, 14) = ("Greatest % Decrease")
ws.Cells(4, 14) = ("Greatest total Volume")

'start with sum =0
Sum = 0

'define lastrow
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'counters initialized
k = 2
j = 2
'start looping through first column to the last row
For i = 2 To lastrow
    If i = 2 Then
        open_price = ws.Cells(i, 3).Value
    End If
'add sum with volume values per ticker
    Sum = Sum + ws.Cells(i, 7).Value
    If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    
'total stock volume
        ws.Cells(j, 12).Value = Sum
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
'yearly change
        yearly_change = ws.Cells(i, 6).Value - open_price
        ws.Cells(j, 10).Value = yearly_change
        
'color formatting
        If yearly_change > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        ElseIf yearly_change < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
'percent change
        percent_change = yearly_change / open_price
        ws.Cells(j, 11).Value = percent_change
 'increment j and reset sum
        j = j + 1
        Sum = 0
        open_price = ws.Cells(i + 1, 3).Value
    
    End If
Next i

'Max and Min Fx for largest and smallest values with match function
greatest_increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
ws.Range("P2") = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
ws.Range("O2") = Cells(greatest_increase + 1, 9)

greatest_decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
ws.Range("P3") = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
ws.Range("O3") = Cells(greatest_decrease + 1, 9)

greatest_total = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
ws.Range("O4") = ws.Cells(greatest_total + 1, 9)

'Format cells with % and autofit
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("I:Q").Columns.AutoFit

Next ws
End Sub