'Chess Board

' Part I - insert chess pieces
' -----------------------------------------------------------
Sub ChessBoard():
    Range("A2:H2").Value = "Pawn"
    Range("A1, H1").Value = "Rook"
    Range("B1, G1").Value = "Knight"
    Range("C1, F1").Value = "Bishop"
    Range("D1, D8").Value = "Queen"
    Range("E1, E8").Value = "King"
    
    Range("A7:H7").Value = "Pawn"
    Range("A8, H8").Value = "Rook"
    Range("B8, G8").Value = "Knight"
    Range("C8, F8").Value = "Bishop"
'MsgBox to check ranges

' Optional Part II
' -----------------------------------------------------------
' Setting cell color formatting
For iterator = 1 to 8 
    For j = 1 to 8
        If iterator Mod 2 = 0 Then
            If j Mod 2 <> 0 Then
                Cells (iterator, j).Interior.ColorIndex = 1
            End If
        Else 
            If j Mod 2 = 0 Then 
            Cells (iterator, j).Interior.ColorIndex = 1
            End If
            
        End If
    Next j
Next iterator

' Setting text color
Range("a1:h2").Font.ColorIndex = 3
Range("a1:h2").Font.Bold = True

Range("a7:h8").Font.ColorIndex = 5
Range("a7:h8").Font.Bold = True 

' Setting cell height and width 
Range("a1:h8").RowHeight = 60
Range("a1:h8").ColumnWidth = 20

End Sub
    
