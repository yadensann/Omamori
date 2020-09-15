' Chicken Nuggets

' Part I
' ----------------------------------------
' Assign iterator 
Sub ChickenNug():
    
    Dim i As Integer
' For...Next
    For i = 1 to 50
' Assign columns with iterator 
    Cells(i , 1).Value = "I will eat"
    Cells(i, 2).Value = i + 2
    Cells(i, 3).Value = "chicken nuggets."
' Assign counter
    Next i

    MsgBox ("Finger Lickin Good")

End Sub