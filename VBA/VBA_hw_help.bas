Attribute VB_Name = "Module2"
Sub StockDataInProgress()

    For Each ws In Worksheets

        Dim Ticker_Name As String
        
        Dim rng As Range
        
        Dim Yearly_Change As Double
        
        Dim year_open As Double
        
        Dim Year_Close As Long
        
        Dim percent_change As Long
        
        Dim Yearly_Ratio As Double
        
        ws.Cells(1, 9).Value = "Ticker"
            
        ws.Cells(1, 10).Value = "Yearly Change"
            
        ws.Cells(1, 11).Value = "Percent Change"
            
        ws.Cells(1, 12).Value = "Total Stock Volume"
            
        Dim Summary_Table_Row As Integer
        
        Summary_Table_Row = 2
        
        start_value = 2
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Year_Close = ws.Cells(i, 6).Value
                
                year_open = ws.Cells(start_value, 3)
                
                Yearly_Change = Year_Close - year_open
                    
                    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                    
                    start_value = i + 1
                    
                    If year_open <> 0 Then
                    
                        Yearly_Ratio = Yearly_Change / year_open
                        
                    Else
                    
                        Yearly_Ratio = 0
                        
                    End If
                    
                    percent_change = Yearly_Ratio * 100
                    
                    ws.Range("K" & Summary_Table_Row).Value = percent_change
            
                Ticker_Name = ws.Cells(i, 1).Value
                
                    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                    
                    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                    
                    ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
                    
                    Ticker_Total = 0
                    
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                
                ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                
                Else
                   
                   ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                   
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
                    
                Else
                
                    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                    
            End If
                    
        Next i
        
    Next ws

End Sub

