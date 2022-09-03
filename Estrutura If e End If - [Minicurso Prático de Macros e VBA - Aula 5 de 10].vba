Sub Alunos()
  Linha = 2
  Sheets('nome da Aba').select
  If Cells(linha,2).Value >= 6 Then
    Cells(linha,3).Value = "Aprovado"
  else
    Cells(linha,3).Value = "Reprovado"
  End If
End Sub
