' Red and Black chessboard

Sub redandblack():

    Dim i As Integer
    Dim n As Integer


    For i = 1 to 8 

 
            If i Mod 2 = 0 Then

                For n = 1 to 8 
                
        
            Cells(i, n).Interior.ColorIndex = 3

            Else
            Cells (i, n).Interior.ColorIndex = 1

            End If
    
        Next n
    Next i
