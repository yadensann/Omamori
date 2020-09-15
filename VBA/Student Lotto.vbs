' Student Lotto
' ------------------------------------------------

Sub Lotto():
' Define variables
    dim win1 as long
    dim win2 as long 
    dim win3 as long 
' Bonus wins 
    dim win4A as long
    dim win4B as long 
    dim win4C as long 

' For loop values
    dim i as long
    dim lotto_num as long
    dim first_name as string 
    dim last_name as string 

' Define top 3 lotto numbers 
    win1 = 3957481
    win2 = 5865187
    win3 = 2817729

' Define runner up lotto numbers 
    win4A = 2275339
    win4B = 5868182
    win4C = 1841402
' ------------------------------------------------
' For loop finding top 3 lotto numbers 
    For i = 2 to 1001
    ' Assign what column information is being pulled from 
        lotto_num = Cells(i, 3).value
        first_name = Cells(i, 1).value
        last_name = Cells(1, 2).value

        ' If conditional for first place and message 
            If lotto_num = win1 then
                Range("F2") = first_name
                Range("G2") = last_name
                Range("H2") = lotto_num
                MsgBox ("Congratulations" + first_name + " " + last_name + " !" + "You win nothing!")

        ' ElseIf conditional for second place 
            ElseIf lotto_num = win2 then
                Range("F3") = first_name
                Range("G3") = last_name
                Range("H3") = lotto_num

        ' ElseIf conditional for third place 
            ElseIf lotto_num = win3 then
                Range("F4") = first_name
                Range("G4") = last_name
                Range("H4") = lotto_num

            End if 

    Next i

' Bonus For loop for runner-up winner 

    For i = 2 to 1001

    lotto_num = Cells(i, 3).value
    first_name = Cells(i, 1).value
    last_name = Cells(i, 2).value

    ' If conditional for runner-up 
        If lotto_num = win4A or lotto_num = win4B or lotto_num = win4C then
            Range("F5") = first_name
            Range("G5") = last_name
            Range("H5") = lotto_num

            Exit For
        
        End If
    
    Next i

End Sub





