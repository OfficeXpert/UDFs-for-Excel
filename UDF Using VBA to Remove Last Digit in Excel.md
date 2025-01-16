Function ClearLastDigit(inputValue As String) As String

If Len(inputValue) > 0 Then
ClearLastDigit = Left(inputValue, Len(inputValue) - 1)
Else
ClearLastDigit = inputValue
End If

End Function
