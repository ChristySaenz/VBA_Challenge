Attribute VB_Name = "Module3"
Sub StockYr()
    'Apply to all Worksheets (referenced: https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
    
    Dim WS_Count As Integer
    Dim j As Integer
    
    WS_Count = ActiveWorkbook.Worksheets.Count
        
    For j = 1 To WS_Count
        
    'Finding Last Row
    
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        
    'Naming new columns (This Works)
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
    'Auto-fit columns (Works)
     
        Columns("A:Z").EntireColumn.AutoFit
        
   
        Dim i As Long
        Dim tickerrow As Integer
        Dim yropen As Double
        Dim yrend As Double
        Dim yrchange As Double
        
        tickerrow = 2
        TotalStock = 0
        k = 0
        
        For i = 2 To lastrow
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                TotalStock = TotalStock + Cells(i, 7).Value
                         
                If TotalStock = 0 Then
                    Range("I" & 2 + k).Value = Cells(i, 1).Value
                    Range("J" & 2 + k).Value = 0
                    Range("K" & 2 + k).Value = "%" & 0
                    Range("L" & 2 + k).Value = 0
                    
                Else
                    If Cells(tickerrow, 3) = 0 Then
                    
                        For find_value = tickerrow To i
                            If Cells(find_value, 3).Value <> 0 Then
                                tickerrow = find_value
                                Exit For
                            End If
                        Next find_value
                         
                    End If
                    yrchange = Cells(i, 6).Value - Cells(tickerrow, 3).Value
                    perchange = yrchange / Cells(tickerrow, 3).Value
                    tickerrow = i + 1
                    
                    Range("I" & 2 + k).Value = Cells(i, 1).Value
                    Range("J" & 2 + k).Value = yrchange
                    Range("K" & 2 + k).Value = perchange
                    Range("L" & 2 + k).Value = TotalStock
                    
'Make Positive Yearly change Green (May need to pull out 0)
     
                    If yrchange > 0 Then
                        Range("J" & 2 + k).Interior.ColorIndex = 4
            
'Make negative yearly change Red
      
                    ElseIf yrchange < 0 Then
                        Range("J" & 2 + k).Interior.ColorIndex = 3
                        
                    Else
                        Range("J" & 2 + k).Interior.ColorIndex = 0
                        
                    End If
                    
                End If
                    
                TotalStock = 0
                yrchange = 0
                k = k + 1
                                
            Else
                TotalStock = TotalStock + Cells(i, 7).Value
            End If
            
        Next i
    Next j
End Sub
    


'Also, need to make this run on the whole workbook



