Sub Abas()
  for each aba in Thisworkbook.Sheets
    if aba.name  <> "Varíavel" Then
      aba.Select
      Range("h5").Value = "Varíavel"
    End If
  Next
End Sub
