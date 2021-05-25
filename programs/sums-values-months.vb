Sub temp()

linha = 6
ano = 1961 'Ano de inicio'
linhaResult = 6 'Linha da inicio da tabela de resultado'
colunaResult = 9 'Coluna iniciar, da tabela de resultado (Janeiro)'
flag = False 'Bandeira, para saber se celula está vazio, true = vazio ou false = número'

anoStop = 1962 'Ano de fim'

InicioWhile:

While Cells(linha, 1) = ano

    For Mes = 1 To 12
        flag = False

        While Cells(linha, 2) = Mes

            If Cells(linha, 6) = "null" Then
                flag = True
            Else
                Cells(linhaResult, colunaResult) = Cells(linhaResult, colunaResult) + Cells(linha, 6)
            End If

            linha = linha + 1
        Wend

        If flag Then
            Cells(linhaResult, colunaResult) = "null"
        End If

        colunaResult = colunaResult + 1
    Next
Wend

ano = ano + 1
linhaResult = linhaResult + 1
colunaResult = 9

If Cells(linha, 1) = anoStop Then
    Stop
End If

GoTo InicioWhile
    
End Sub