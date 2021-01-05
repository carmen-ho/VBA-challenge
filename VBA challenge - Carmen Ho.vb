Sub stock_analysis()
    
    'Set dimensions
    Dim i As Long
    Dim j As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim start As Long
    Dim lastrow As Long
    
    
    'Set column titles
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Initialize values
    total_stock_volume = 0
    yearly_change = 0
    j = 0
    start = 2
   
    'Calculate last row number
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'For loop
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            'ticker = Cells(i, 1).Value
       
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
   
            'For zero total volume
            If total_stock_volume = 0 Then
                
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = 0
                Range("L" & 2 + j).Value = 0
        
            Else
                'For non-zero total volume, find the new start number
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If
                
                'Calculate change
                yearly_change = Cells(i, 6) - Cells(start, 3)
                percent_change = Round((yearly_change / Cells(start, 3) * 100), 2)
 
                'Start new ticker
                start = i + 1
                
                'print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(yearly_change, 2)
                Range("K" & 2 + j).Value = percent_change & "%"
                Range("L" & 2 + j).Value = total_stock_volume
                
                'Highlight positives/negatives
                Select Case yearly_change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
                
            End If
            
            'Reset variables for new stock ticker
            total_stock_volume = 0
            yearly_change = 0
            j = j + 1
        
        Else
   
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
     
        End If
    Next i
   
   

End Sub
   