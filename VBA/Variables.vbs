Sub Variables():

    ' Basic String Variable
    ' ----------------------------------------
    Dim name As String
    name = "Gandalf"

    MsgBox (name)

    ' Basic String Concatenation (Combination)
    ' ----------------------------------------
    Dim title As String
    title = "The Great"

    Dim fullname As String
    fullname = name + " " + title

    MsgBox (fullname)

    ' Basic Integer, Double, Long Variables
    ' ----------------------------------------
    Dim age1 As Integer
    Dim age2 As Integer
    age1 = 5
    age2 = 10

    Dim price As Double
    Dim tax As Double
    price = 19.99
    tax = 0.05

    Dim lightspeed As Long
    lightspeed = 299792458

    ' Basic Numeric manipulation
    ' ----------------------------------------
    MsgBox (age1 + age2)
    Cells(1, 1).Value = price * (1 + tax)

    ' String, Numeric Combination (Casting)
    ' ----------------------------------------
    MsgBox ("I am " + Str(age1) + " years old.")

    ' Booleans
    ' ----------------------------------------
    Dim money_grows_on_trees As Boolean
    money_grows_on_trees = False    
    
    ' Pulling variables from cells
    ' ----------------------------------------
    Dim number_from_workbook As Double
    number_from_workbook = Cells(1, 1).Value
    MsgBox ("The number in this cell is " + Str(age1))

    Dim text_from_workbook As String
    text_from_workbook = Cells(1, 1).Value
    MsgBox ("The text in this cell is " + text_from_workbook)



End Sub
