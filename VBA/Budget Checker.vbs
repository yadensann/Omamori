'Budget Checker
'Part I
' -----------------------------------------------
'Assign variables
Sub BudgetChecker():
    Dim Budget As Double 
    Dim Price As Double
    Dim Fees As Double
    Dim Total As Double
    Dim newPrice As Double 
    
'Assign ranges of variables
    Budget = Range("C3").value
    Price = Range("F3").value
    Fees = Range("H3").value
    Total = Range("L3").value

'Part II
' -----------------------------------------------
'Use values to calculate the total
    Total = Price * (1 + Fees)
    MsgBox ("Your total is $" + Str(Total))

' Make the if Then statement for budget
    If Budget > Total Then

        MsgBox ("Under budget")
    Else
        MsgBox ("Over budget") 
        
        newPrice = Budget / (1 + Fees)

' Change the price
        newPrice = Price
' Change the new total
        Total = newPrice * (1 + Fees)
    End If

End Sub


    




    





