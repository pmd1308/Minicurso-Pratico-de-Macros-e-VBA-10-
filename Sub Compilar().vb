Sub Compilar()
    'Limpar aba XXX
    Sheets("XXX").Select
    If Range("A2").Value <> "" Then
        ult_linha = Range("A1000000").End(xlUp).Row
        Range(Cells(2, 1), Cells(ult_linha, 3)).Clear
    End if

    For Each aba in ThisWorkbook.Sheets
        If aba.Name <> "Resumo" Then
            aba.Select
            ult_linha = Range("A100000").End(xlUp).Row
            Range(Cells(2,1), Cells(ult_linha, 3)).Copy

            Sheets("Resumo").Select
            prox_linha = Range("A100000").End(xlUp).Row +1
            Cells(prox_linha, 1).Select
            ActiveSheet.Paste
        End If
    Next

    Sheets("Resumo").Select
    Cells(1, 1).Select
End Sub
