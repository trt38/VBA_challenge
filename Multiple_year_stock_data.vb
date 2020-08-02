Sub stocktest()

On Error Resume Next

For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
	Dim tickername As String
	Dim ychange As Double
	Dim pchange As Double
	Dim totalvolume As Double

	Dim greatincr As Double
	Dim greatdecr As Double
	Dim greattovo As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Summary_Table_Row = 2

    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                   
            tickername = ws.Cells(i, 1).Value
            yopen = ws.Cells(i, 3).Value
            yclose = ws.Cells(i, 6).Value
            ychange = yclose - yopen
            pchange = (yclose - yopen) / yclose
            totalvolume = totalvolume + ws.Cells(i + 1, 7).Value
                
                ws.Range("I" & Summary_Table_Row).Value = tickername
                ws.Range("J" & Summary_Table_Row).Value = ychange
                ws.Range("K" & Summary_Table_Row).Value = pchange
                ws.Range("L" & Summary_Table_Row).Value = totalvolume
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            totalvolume = 0
            
            Else
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            
        End If
        
    
        If (ws.Cells(i, 10) >= 0) Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
            ElseIf (ws.Cells(i, 10) < 0) Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    
        ws.Cells(i, 11).NumberFormat = "0.00%"
         
    Next i

    For i = 2 To lastrow
        
        If ws.Cells(i + 1, 9).Value <> ws.Cells(i, 9).Value Then
        
            greatincr = WorksheetFunction.Max(ws.Range("K:K"))
            greatdecr = WorksheetFunction.Min(ws.Range("K:K"))
            greattovo = WorksheetFunction.Max(ws.Range("L:L"))
            
            ws.Cells(2, 16).Value = greatincr
            ws.Cells(3, 16).Value = greatdecr
            ws.Cells(4, 16).Value = greattovo
                                
            match_gi = WorksheetFunction.Match(greatincr, ws.Range("K:K"), 0)
            ws.Cells(2, 15).Value = ws.Cells(match_gi, 9)
            
            match_gd = WorksheetFunction.Match(greatdecr, ws.Range("K:K"), 0)
            ws.Cells(3, 15).Value = ws.Cells(match_gd, 9)
            
            match_tv = WorksheetFunction.Match(greattovo, ws.Range("L:L"), 0)
            ws.Cells(4, 15).Value = ws.Cells(match_tv, 9)
                  
            Exit For
            
        End If
        
                
    Next i
        

Next ws

End Sub

