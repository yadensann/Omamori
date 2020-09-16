' Student Gradebook
' ----------------------------------
' Sub formatter()
​
  ' Set the Font color to Red
  Range("A1").Font.ColorIndex = 3
​
  ' Set the Cell Colors to Red
  Range("A2:A5").Interior.ColorIndex = 3
​
  ' Set the Font Color to Green
  Range("B1").Font.ColorIndex = 4
​
  ' Set the Cell Colors to Green
  Range("B2:B5").Interior.ColorIndex = 4
​
  ' Set the Color Index to Blue
  Range("C1").Font.ColorIndex = 5
​
  ' Set the Cell Colors to Blue
  Range("C2:C5").Interior.ColorIndex = 5
​
  ' Set the Color Index to Magenta
  Range("D1").Font.ColorIndex = 7
​
  ' Set the Cell Colors to Magenta
  Range("D2:D5").Interior.ColorIndex = 7
​
  ' See this website for color guides: http://dmcritchie.mvps.org/excel/colors.htm
End Sub

Sub gradebook():

    Dim grade As Double 
    
    grade = Range("B2")

   ' Button message
    
    
    If grade >= 90 Then 
        Range("C2") = "Pass"
        Range("D2") = "A"
        Range ("C2").Interior.ColorIndex = 4
    Elseif Grade >= 80 And Grade < 90 Then
        Range("C2") = "Pass"
        Range("D2") = "B"
        Range ("C2").Interior.ColorIndex = 4
    Elseif Grade >= 70 And Grade < 80 Then 
        Range("C2") = "Warning"
        Range("D2") = "C"
        Range ("C2").Interior.ColorIndex = 6
    Elseif grade < 70 Then
        Range("C2") = "Fail"
        Range("D2") = "D"
    Elseif grade < 60 Then
        Range("C2") = "Fail"
        Range("D2") = "F"
        Range ("C2").Interior.ColorIndex = 3
    End If  

End Sub
