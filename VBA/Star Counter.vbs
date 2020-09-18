' Star Counter
' ---------------------------------
' Assign variables 
Sub starcounter():

' Create a variable that holds the starcounter variable that will be repeatedly used
    Dim starcounter As Integer

' Loop through each row
    For i = 2 To 51
    ' Initially set the counter to be 0 for each row after your first For loop
        starcounter = 0
' While in each row, loop through each star column 
        For j = 4 to 8
        ' If a column contains the word "Full-Star"...
            If (Cells(i, j).value = "Full-Star") Then 

        ' Add 1 to the starcounter 
                starcounter = starcounter + 1

            End if 
        Next j
    ' Don't Next i until we've iterated through each column in row i
        Cells(i, 9).value = starcounter 
    
    Next i
End sub 


        

