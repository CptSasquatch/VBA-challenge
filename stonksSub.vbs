Sub stonks()

'Declare Worksheet
Dim ws As Worksheet

'Looping through all sheets
For Each ws In Worksheets

    'Insert Ticker, Yearly Change, Percent Change and Total Stock Volume Columns
    
        Dim ticker As String
        
        Dim totalVolume As Double
        
        Dim tickerRow As Long
        
        Dim LastRow As Long
        
        Dim opening As Double
        
        Dim closing As Double
         
        'set initial row value
        tickerRow = 2
        
        'Determine the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'set column formats
        ws.Range("K2", "K" & LastRow).NumberFormat = "0.00%"

        'set opening price
        opening = ws.Cells(2, 3).Value
        
        'Loop through all stocks
        For i = 2 To LastRow
        
            'set Ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            
            'determine closing price
            closing = ws.Cells(i, 6).Value
            
            'add to total for each Ticker
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'add ticker to ticker column
            ws.Range("I" & tickerRow).Value = ticker
            
            'add total to volume column
            ws.Range("L" & tickerRow).Value = totalVolume
            
            'add yearly change
            ws.Range("J" & tickerRow).Value = closing - opening
            
            'add percent change
            ws.Range("K" & tickerRow).Value = 1 * ((closing - opening) / opening)
            
            'add 1 to row
            tickerRow = tickerRow + 1
            
            'reset Volume
            totalVolume = 0

            'set opening price
            opening = ws.Cells(i+1,3).Value
            
        Else
        
            'add to totalVolume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
        End If
        
        Next i
    
Next ws

End Sub