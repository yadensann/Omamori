' VBA Homework
' --------------------------------------
Sub alpha_test():

'Assign and set worksheet

' Loop through all stocks in every worksheet?
 

' Create column headings
        Cells(1,9) = "Ticker"
        Cells(1,10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1,12) = "Total Stock Volume"
        Cells(2, 14) = "Greatest % Increase"
        Cells(3, 14) = "Greatest % Decrease"
        Cells(4, 14) = "Greatest Total Volume"
        Cells(1, 15) = "Ticker"
        Cells(1, 16) = "Value"
        ' MsgBox to show column headers

    ' Assign ticker
        
            
         
        


    ' Assign stock volume
            Dim ticker As String
            Dim total_stock As Double
            Dim stock_vol As Double
            Dim open_price As Double
            Dim close_price As Double


            stock_vol = 0

        For i = 2 to 1000

            ticker = Cells(i,1).Value 
            open_price = Cells(i, 3).Value
            close_price = Cells(i, 6).Value
            total_stock = Cells(i, 7).Value
            stock_vol = stock_vol + total_stock

            ' if ticker is the same as the value below it
            if ticker = Cells(i+1, 1).value Then
                ' keep going
            else 
                cells(2, 9)=ticker
        
        Next i
' msgbox for total volume
        MsgBox (stock_vol)

            

            
    

