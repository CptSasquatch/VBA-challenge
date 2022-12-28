Sub cellColors()

'Declare Worksheet
Dim ws As Worksheet

Dim LastRow As Long

'Looping through all sheets
For Each ws In Worksheets

        'Determine the last row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

        'set conditional format
        If ws.Cells(i, 10).Value >= 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        Else
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
        Next i
        
Next ws

End Sub