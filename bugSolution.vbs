Function f(a, b)
  'Explicitly convert both arguments to numbers to ensure consistent comparison
  Dim numA, numB
  numA = CDbl(a)
  numB = CDbl(b)
  If numA > numB Then
    MsgBox "a is greater than b"
  ElseIf numA < numB Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
End Function

f(10, "5")
f(10, 5) 
f("10", 5) 