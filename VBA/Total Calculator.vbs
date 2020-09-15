' Total Calculator

Sub totalcalc():

'Assign variables
    dim price as double
    dim tax as double
    dim quantity as double
    dim total as double 
  

' Assign the value of the cell variables are associated with 
    price = cells(2, 2).value
    tax = cells(2, 3).value
    quantity = cells(2, 4).value
    

' Assign total value and cell
    total = price * (1+ tax)
    Cells(2, 5).value = total
    
    MsgBox ("Total is $" + Str(total))

  
End Sub