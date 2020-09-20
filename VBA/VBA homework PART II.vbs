'VBA homework PART II

Sub stockdata():

    For Each ws In Worksheets

        Dim summary_row As Double
        Dim start_value As Double
        Dim ticker As String
        Dim ticker_value As Double
        Dim open_value As Double
        Dim close_value As Double
        Dim percent_change As Double
        Dim year_change As Double
        Dim yearly_value As Double
    
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Volume"
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
    ' MsgBox to show column headers

        summary_row = 2
        start_value = 2
    ' Determine the lastRow count
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                open_value = ws.Cells(start_value, 6).Value
                close_value = ws.Cells(i, 3).Value
                year_change = open_value - close_value
                
                ws.Range("J" & summary_row).Value = year_change
                
                start_value = i + 1
                
                If year_open <> 0 Then
                    yearly_value = year_change / open_value
                    
                Else
                    yearly_value = 0
                    
                End If
                         
                percent_change = year_change * 100
                ws.Range("K" & summary_row).Value = percent_change
                
                                
                ticker = ws.Cells(i, 1).Value
                ticker_value = ticker_value + ws.Cells(i, 7).Value
                
                ws.Range("I" & summary_row).Value = ticker
                ws.Range("L" & summary_row).Value = ticker_value
                
                ticker_value = 0
                
            ' Coloring in cells with conditional formatting
            
                If ws.Range("J" & summary_row).Value < 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                End If
                
                summary_row = summary_row + 1
                
            Else
                ticker_value = ticker_value + ws.Cells(i, 7).Value
            
            End If
            
        Next i
    Next ws
End Sub


