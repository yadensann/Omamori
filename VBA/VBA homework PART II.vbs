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
        Dim yearly_percent As Double
        
        Dim current_spreadsheet As Boolean

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

                If year_open <> 0 Then
                    yearly_percent = (year_change / open_value) * 100
                        
                Else
                    yearly_percent = 0
                    
                End If

                If yearly_percent > percent_increased Then
                    percent_increased = yearly_percent
                    max_ticker = ticker

                ElseIf yearly_percent < percent_decreased Then
                    percent_decreased = yearly_percent
                    min_ticker = ticker
                End If

                If ticker_value > max_volume Then
                    max_volume = ticker_value
                    max_ticker_volume = ticker
                End If
                
                yearly_percent = 0
                ticker_value = 0
                
            Else
                ticker_value = ticker_value + ws.Cells(i, 7).Value
            
            End If

        Next i

            If Not current_spreadsheet Then
            ' CStr converts the expression into a String data type.
                'ws.Range("Q2").Value = (CStr(percent_increased) & "%")
                'ws.Range("Q3").Value = (CStr(percent_decreased) & "%")
                ws.Range("P2").Value = max_ticker
                ws.Range("P3").Value = min_ticker
                ws.Range("Q4").Value = max_volume
                ws.Range("P4").Value = max_ticker_volume
            Else
                current_spreadsheet = False
            End If

    Next ws
End Sub



' PART II OF HOMEWORK

Dim rng As Range 
Dim percent_increase As Double 

Set rng = ws.Cells("K" + summary_row).Value
percent_increase = Application.WorksheetFunction.Max(rng)

ws.Cells(2,15).value = percent_increase

' ws.Cells(Count, 4)=Application.WorksheetFunction.Max(Range(Cells(m, 1),Cells(n, 1)))


if (yearly_percent > MAX_PERCENT ) Then 
    MAX_PERCENT = 