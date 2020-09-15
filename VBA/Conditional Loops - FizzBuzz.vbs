' Conditional Loops - FizzBuzz

Sub FizzBuzz():

' Create a variable to hold the counter
Dim i As Integer

For i = 2 To 100
    Cells(i, 1).Value = i
    ' Mod means divisible by
    If i Mod 3 = 0 And i Mod 5 = 0 Then
        
        Cells(i, 2).Value = "FizzBuzz"
    
    ElseIf i Mod 3 = 0 Then

        Cells(i, 2).Value = "Fizz"
        
        
    ElseIf i Mod 5 = 0 Then
    
        Cells(i, 2).Value = "Buzz"
            
    End If
    
Next i
    
End Sub
