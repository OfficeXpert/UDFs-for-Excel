Public Function ClearCharLeft(cell_ref As String, char_num As Long)
 ClearCharLeft = Right(cell_ref, Len(cell_ref) - char_num)
 End Function
