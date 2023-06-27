Sub VBA_challenge()

Dim ws As Worksheet

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim summary_table As Integer
summary_table = 2

Dim total_volume As Double
total_volume = 0

Dim opening_price As Double
opening_price = ws.Cells(2, 3).Value

Dim closing_price As Double

Dim yearly_change As Double

Dim percent_change As Double

For i = 2 To lastrow

total_volume = total_volume + ws.Cells(i, 7).Value

If ws.Cells(i, 1) <> ws.Cells(i + 1, 1).Value Then
    ws.Cells(summary_table, 9).Value = ws.Cells(i, 1).Value
    
    ws.Cells(summary_table, 12).Value = total_volume
    
    closing_price = ws.Cells(i, 6).Value
    yearly_change = closing_price - opening_price
    ws.Cells(summary_table, 10).Value = yearly_change
        
        If ws.Cells(summary_table, 10).Value < 0 Then
            ws.Cells(summary_table, 10).Interior.ColorIndex = 3
        Else: ws.Cells(summary_table, 10).Interior.ColorIndex = 4
        End If
    
    percent_change = yearly_change / opening_price
    ws.Cells(summary_table, 11).Value = FormatPercent(percent_change)
    
    summary_table = summary_table + 1
    total_volume = 0
    opening_price = ws.Cells(i + 1, 3).Value

End If

Next i

'greatest total volume
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L3001"))
ws.Range("P4") = WorksheetFunction.Index(ws.Range("I2:I3001"), WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L2:L3001"), 0))

'greatest percent increase
ws.Range("Q2") = FormatPercent(WorksheetFunction.Max(ws.Range("K2:K3001")))
ws.Range("P2") = WorksheetFunction.Index(ws.Range("I2:I3001"), WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K2:K3001"), 0))

'greatest percent decrease
ws.Range("Q3") = FormatPercent(WorksheetFunction.Min(ws.Range("K2:K3001")))
ws.Range("P3") = WorksheetFunction.Index(ws.Range("I2:I3001"), WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K2:K3001"), 0))

Next ws

End Sub