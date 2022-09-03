Sub Proximo()
  Sheets("aba").Select
  l = range('100').End(xlUp).Row + 1
  Cells(l,1).Value = ""
End Sub
