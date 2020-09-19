' VBA Homework

' PART I          
    Sub stock_data():

    'Loop through all the sheets-------------

    For Each ws in Worksheets

        Dim worksheet_name as string 

            ' lastRow = ws.Cells(rows.count, 1).end(xlUp).Row
            worksheet_name = ws.Name 
        'Msgbox workshewt name-------------
        'Split the worksheet_name
             State = Split(worksheet_name, " ")
        'Msgbox worksheet_name--------------
        ' add info to the column 
             ws.Range("A1").EntireColumn.Insert
        ' adding info to all the rows 
             ws.Range("A2:A" & LastRow) = State(0)
        ' determine the last column number
            lastcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        ' rename year columns by looping through and renaming each 
            For i = 3 To lastcolumn
                yearheader = ws.Cells(1, i).value
            ' msgbox yearsplit (0)
                yearsplit = Split(yearheader, " ")
                ws.Cells(1, i).value = yearsplit(3)
            ' msgbox cells(1, i)
            ' msgbox yearsplit(3)
        Next i
        
        For i = 2 To lastRow
            For j = 3 To lastcolumn
            ' Changes from scientific notation to currency format
                ws.Cells(i, j).Style = "Currency"
            Next j
        Next i
    Next ws

' PART II
' Loop through every worksheet and select the state contents.
' Copy the state contents and paste it into the combined data tab
    Sub WellsFargo_ptII():
        ' add sheet
        sheets.add.name = "combined_Data"
        ' move created sheet to be first sheet 
        sheets("combined_data").move before:=Sheets(1)
        'specfiy location of combined sheet 
        set combined_sheet = worksheets("combined_data")

        'loop through all the sheets
        for each ws in worksheets
            'find last row of combined sheet after each paste 
            ' add 1 to get first empty row 
            lastRow = combined_sheet.Cells(Rows.count, "A").End(xlUp).Row + 1
            ' find the last row of each worksheet 
            ' subtract one to return the number of rows without header 
            lastRowstate = ws.Cells("A" & lastRow & ":G" & ((lastRowstate - 1)+ lastRow)).value = ws.Range("A2:G" & (lastRowstate + 1)).value
        next ws
        ' Copy the headers from sheet 1
        combined_sheet.Range("A1:G1").value = sheets(2).range("A1:G1".value)
        ' autofit to display data
        combined_sheet.Columns("A:G").autofit

    End sub

' ------------------------------------------------------
Sub stockdata():
' add ws. before all cells
    For Each ws In Worksheets
    
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

        Dim ticker As String
        Dim ticker_value As Double
        Dim year_open As Double
        Dim year_close As Double
        Dim yearly_diff As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim lastRow As Double
        Dim summary_row As Integer
        Dim start_value As Integer
            
        summary_row = 2
        start_value = 2
                
    ' Determine the lastRow count
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastRow
            
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                year_close = ws.Cells(1, 6).Value
                year_open = ws.Cells(start_value, 3)
                yearly_diff = year_close - year_open
                
                ws.Range("K" & summary_row).Value = yearly_change
                
                start_value = i + i
                
                If year_open <> 0 Then
                    yearly_change = yearly_diff / year_open
                Else
                    yearly_change = 0
                End If
                
            Range("K" & summary_row).Value = percent_change
            percent_change = yearly_change
            
            ws.Range("K" & summary_row).Style = "Percent"
            ticker = ws.Cells(i, 1).Value
            year_close = ws.Cells(1, Value)
                    
            ws.Range("I" & summary_row).Value = ticker
            ws.Range("L" & summary_row).Value = ticker_value
    
            ticker_value = 0
                           
            If ws.Range("J" & summary_row).Value < 0 Then
            ' Use conditional formatting
                ws.Range("J" & summary_row).Interior.Color = RGB(255, 0, 0)
                    
            Else
                ws.Range("J" & summar_row).Interior.Color.Index = 4
                        
            End If
                summary_row = summary_row + 1
        Else
            ticker_value = ticker * ws.Cells(1, 7).Value

        End If

        Next i
        
    Next ws
                 
End Sub





' -------------------------------------------------------------------
Sub wsloop():
    Dim ws_count as Integer
    dim i as integer 
      
      'set ws_count equal to the number of worksheets in the active workbook
      ws_count = numberofworksheets.worksheets.count

      ' Begin the loop
    for i = 1 to ws_count

        ' insert your code here 
        ' the following line shows how to reference a sheet within the loop
        'by displaying the worklsheet name in a dialog box

        msgbox activeworkbook.worksheets(i).name
    next i

end Sub
