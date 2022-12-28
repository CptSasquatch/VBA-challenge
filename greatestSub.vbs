Sub greatest()

'columns 15, 16 & 17
'Declare Worksheet
Dim ws As Worksheet

'Looping through all sheets
For Each ws In Worksheets

    Dim increase As String
    Dim decrease As String
    Dim volume As String

    'set max
    Dim max As Double
    
    max = 0
    
    'set min
    Dim min As Double
    
    min = 0
    
    'set total
    Dim total As Double
    
    total = 0
    
    'Determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Insert Ticker, Value and categories

        ws.Cells(1, 16).Value = "Ticker"
        
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Loop through all stocks
        For i = 2 To LastRow
        
        'set greatest increase
        If ws.Cells(i, 11).Value > max Then
        
            max = ws.Cells(i, 11).Value
            
            ws.Cells(2, 17).Value = max
            
            increase = ws.Cells(i, 9).Value
            
            ws.Cells(2, 16).Value = increase
            
        'set greatest decrease
        ElseIf ws.Cells(i, 11).Value < min Then
        
            min = ws.Cells(i, 11).Value
            
            ws.Cells(3, 17).Value = min
            
            decrease = ws.Cells(i, 9).Value
            
            ws.Cells(3, 16).Value = decrease
            
        'set greatest total
        ElseIf ws.Cells(i, 12).Value > total Then
        
            total = ws.Cells(i, 12).Value
            
            ws.Cells(4, 17).Value = total
            
            volume = ws.Cells(i, 9).Value
            
            ws.Cells(4, 16).Value = volume
            
        End If
        Next i
        
'set column formats
ws.Range("Q2", "Q3").NumberFormat = "0.00%"
        
Next ws
        
End Sub