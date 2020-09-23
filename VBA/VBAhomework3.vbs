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
        Dim yearly_percent As Double

        Dim percent_increased As Double
        Dim percent_decreased As Double
        Dim max_ticker As String
        Dim min_ticker As String
        Dim max_ticker_volume As String
        Dim max_volume As Double
        
    
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
    ' MsgBox to show column headers
        summary_row = 2
        start_value = 2

        current_spreadsheet = True

        
        percent_increased = 0
        percent_decreased = 0
        max_volume = 0
        max_ticker_volume = " "
        max_ticker = " "
        min_ticker = " "

    ' Determine the lastRow count
        lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row

        close_price = 0
        open_price = ws.Cells(i + 1, 3).Value
    
        For i = 2 To lastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                open_value = ws.Cells(start_value, 6).Value
                close_value = ws.Cells(i, 3).Value
                year_change = close_value - open_value
                
                ws.Range("J" & summary_row).Value = year_change
                
                start_value = i + 1
            

                percent_change = year_change
                ws.Range("K" & summary_row).Value = percent_change
                ws.Range("K" & summary_row).NumberFormat = "0%"
                
                                
                ticker = ws.Cells(i, 1).Value
                ticker_value = ticker_value + ws.Cells(i, 7).Value
                
                ws.Range("I" & summary_row).Value = ticker
                ws.Range("L" & summary_row).Value = ticker_value
                
                ticker_value = 0
            
                If ws.Range("J" & summary_row).Value < 0 Then
                    ws.Range("J" & summary_row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & summary_row).Interior.ColorIndex = 4
                End If
                
                summary_row = summary_row + 1

                If year_open <> 0 Then
                    yearly_percent = (year_change / open_value) * 100
                        
                Else
                    yearly_percent = 0
                    
                End If
                    ticker_value = 0
                
            Else
                ticker_value = ticker_value + ws.Cells(i, 7).Value
            
            End If

        Next i

        yc_lastRow = ws.Cells(Rows.Count, Column + 10).End(xlUp).Row

        For j = 2 To yc_lastRow
        
            If Cells(j, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & yc_lastRow)) Then
                Cells(2, Column + 16).Value = Cells(j, Column + 9).Value
                Cells(2, Column + 17).Value = Cells(j, Column + 11).Value
                Cells(2, Column + 17).NumberFormat = "0.00%"
                    
            ElseIf Cells(j, Column + 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & yc_lastRow)) Then
                Cells(3, Column + 16).Value = Cells(j, Column + 9).Value
                Cells(3, Column + 17).Value = Cells(j, Column + 11).Value
                Cells(3, Column + 17).NumberFormat = "0.00%"
                    
            ElseIf Cells(j, Column + 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & yc_lastRow)) Then
                Cells(4, Column + 16).Value = Cells(j, Column + 9).Value
                Cells(4, Column + 17).Value = Cells(j, Column + 12).Value
            End If

        Next j

    Next ws
End Sub