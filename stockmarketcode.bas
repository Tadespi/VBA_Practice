Sub stockmarketdata()

    For Each ws In Worksheets
        
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim tickcount As Long
        Dim lastrowA As Long
        Dim lastrowI As Long
        Dim PercentChange As Double
        Dim GreatIncrease As Double
        Dim GreatDecrease As Double
        Dim GreatVolume As Double
        
        WorksheetName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "yearly change"
        ws.Cells(1, 11).Value = "percent change"
        ws.Cells(1, 12).Value = "total stock volume"
        ws.Cells(1, 16).Value = "ticker"
        ws.Cells(1, 17).Value = "value"
        ws.Cells(2, 15).Value = "greatest percent increase"
        ws.Cells(3, 15).Value = "greatest percent decrease"
        ws.Cells(4, 15).Value = "greatest total volume"
        
        tickcount = 2
        
        j = 2
        
        lastrowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            For i = 2 To lastrowA
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ws.Cells(tickcount, 9).Value = ws.Cells(i, 1).Value
                
                ws.Cells(tickcount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    If ws.Cells(tickcount, 10).Value < 0 Then
                    
                    ws.Cells(tickcount, 10).Interior.ColorIndex = 3
                    
                    Else
                    
                    ws.Cells(tickcount, 10).Interior.ColorIndex = 4
                    
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    ws.Cells(tickcount, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(tickcount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                ws.Cells(tickcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                tickcount = tickcount + 1
                
                j = i + 1
                
                End If
            
            Next i
        
        lastrowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        GreatVolume = ws.Cells(2, 12).Value
        GreatIncrease = ws.Cells(2, 11).Value
        GreatDecrease = ws.Cells(2, 11).Value
        
            For i = 2 To lastrowI
            
                If ws.Cells(i, 12).Value > GreatVolume Then
                GreatVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVolume = GreatVolume
                
                End If
                
                If ws.Cells(i, 11).Value > GreatIncrease Then
                GreatIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncrease = GreatIncrease
                
                End If
                
                If ws.Cells(i, 11).Value < GreatDecrease Then
                GreatDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecrease = GreatDecrease
                
                End If
            
            ws.Cells(2, 17).Value = Format(GreatIncrease, "percent")
            ws.Cells(3, 17).Value = Format(GreatDecrease, "percent")
            ws.Cells(4, 17).Value = Format(GreatVolume, "scientific")
            
            Next i
            
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
    
End Sub

