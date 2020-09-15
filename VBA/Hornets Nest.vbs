' Hornets Nest

Sub hornets():

    Dim hornets_count As Integer
    hornets_count = 0

    For x = 1 To 6
        For y = 1 To 7
            If Cells(x, y) = "Hornets" Then
                hornets_count = hornets_count + 1

            End If
        Next y
    Next x

MsgBox ("There are" & Str(hornets_count) & " hornets.")

End Sub


Sub hornet():
    



