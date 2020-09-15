' Sentence Breaker

' Assign variables

Sub SentenceBreaker()
  Dim Words() As String
  Dim Sentence As String
  Dim one As Integer
  Dim two As Integer
  Dim three As Integer

  ' Sentence = "Any fool can know. The point is to understand."

  Sentence = Cells(1, 2)
  Words = Split(Sentence, " ")
  one = Cells(4, 1)
  two = Cells(5, 1)
  three = Cells(6, 1)

  Cells(4, 2) = Words(one - 1)
  Cells(4,2) = Words(Cells(4,1)-1)
  Cells(5, 2) = Words(two - 1)
  Cells(6, 2) = Words(three - 1)
  
  ' MsgBox (Words(8))

End Sub
