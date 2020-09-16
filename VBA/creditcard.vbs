Sub creditcard():
    Dim column = 1

    For i = 1 To 3 

        If Cells(i + 1, column).value <> Cells(i, column).value Then
        MsgBox (Cells(i, column).value)

        End If 

    Next i

End Sub