' Hornets Nest

Sub hornets():
    
    
    hornets_count = 0

    For x = 1 To 6
        For y = 1 To 7
            If Cells(x, y) = "Hornets" Then
                hornets_count = hornets_count + 1

            End If
        Next y
    Next x

MsgBox ("There are" & Str(hornets_count) & " hornets.")

    Dim hornets_count As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bugs As Integer
    Dim bees As Integer
    Dim count As Integer
    Dim value As String

' Initialize count
    count = 0
    bugs = Range("L2")
    bees = Range("N2")

' For loop to count hornets
    For i = 1 To 6
        For j = 1 To 7
         
        ' Pull value from column
            value = Cells(i, j).value

        ' If conditional to find hornets
            If value = "Hornets" Then

                'If conditional replaces hornets with bugs or bees
                    If bugs <> 0 Then
                        Cells(i, j).value = "Bugs"
                        bugs = bugs + 1

                    ElseIf bees <> 0 Then
                        Cells(i, j).value = "Bees"
                        bees = bees + 1
                    End If
            End If
        
        Next j
    Next i

    ' Message box count if hornets are left
    If count <> 0 Then

    MsgBox ("Oh no! We still have " + count + " Hornets.")
    End If
    
    
End Sub