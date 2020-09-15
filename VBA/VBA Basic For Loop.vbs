' Basic For Loop:
Sub forLoop():

' Create a variable to hold the counter 
    Dim i As Integer

' Loop through numbers with For loop
    For i = 1 to 50 
        ' Iterate through the rows placing a value of 1 throughout
        Cells(i, 1).value = 1
        ' Iterate through the columns placing a value of 5 throughout
        Cells(1, i).value = 5
        ' Places increasing values based upon the variable "i" in B2 to B51
        Cells(i + 1, 2).value = i + 1

    ' Call next iteration
    Next i

End Sub
