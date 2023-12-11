Sub calculo_calssificacao()


' Variaveis
Dim meta, qualidade, treinamento, compras, linha As Integer


linha = 9
meta = Cells(linha, 11)
qualidade = Cells(linha, 13)
compras = Cells(linha, 15)
treinamento = Cells(linha, 17)
' Fim das variaveis

' Inicio do calculo da classificação

While Cells(linha, 1) <> ""



' Verificação da categoria e calculos de classificações
If Cells(linha, 3) = "BDS" or Cells(linha, 3) = "BCS/BDS" or Cells(linha, 3) = "BTS" or Cells(linha, 3) = "BCS/BTS" Then

' INICIO da verificação BDS

        ' Calculo da classificação da meta BDS
        If Cells(linha, 10) < Cells(5, 4) Then
            Cells(linha, 11) = "C"
        ElseIf Cells(linha, 10) < Cells(5, 3) Then
            Cells(linha, 11) = "B"
            Else
            Cells(linha, 11) = "A"
        End If
        ' Calculo da classificação da compras BDS
        If Cells(linha, 14) < Cells(3, 4) Then
            Cells(linha, 15) = "C"
        ElseIf Cells(linha, 14) < Cells(3, 3) Then
            Cells(linha, 15) = "B"
            Else
            Cells(linha, 15) = "A"
        End If
        ' Calculo da classificação da qualidade BDS
        If Cells(linha, 12) < Cells(6, 4) Then
            Cells(linha, 13) = "C"
        ElseIf Cells(linha, 12) < Cells(6, 3) Then
            Cells(linha, 13) = "B"
            Else
            Cells(linha, 13) = "A"
        End If
        ' Calculo da classificação do treinamento BDS
        If Cells(linha, 16) < Cells(4, 4) Then
            Cells(linha, 17) = "C"
        ElseIf Cells(linha, 16) < Cells(4, 3) Then
            Cells(linha, 17) = "B"
            Else
            Cells(linha, 17) = "A"
        End If
        
' Final da verificação do BDS

'Caso seja BDC
ElseIf Cells(linha, 3) = "BDC" or Cells(linha,3) = "BCS/BDC" Then
    ' Inicio da verificação BDC
    ' Calculo da classificação da metas BDC
    If Cells(linha, 10) < Cells(5, 4) Then
        Cells(linha, 11) = "C"
        ElseIf Cells(linha, 10) < Cells(5, 3) Then
        Cells(linha, 11) = "B"
        ElseIf Cells(linha, 10) < Cells(5, 2) Then
        Cells(linha, 11) = "A"
        Else
        Cells(linha, 11) = "D"
    End If
    ' Calculo da classificação de compras BDC
    If Cells(linha, 14) < Cells(3, 4) Then
        Cells(linha, 15) = "C"
        ElseIf Cells(linha, 14) < Cells(3, 3) Then
        Cells(linha, 15) = "B"
        ElseIf Cells(linha, 14) < Cells(3, 2) Then
        Cells(linha, 15) = "A"
        Else
        Cells(linha, 15) = "D"
    End If
    ' Calculo da classificação da qualidade BDC
    If Cells(linha, 12) < Cells(6, 4) Then
        Cells(linha, 13) = "C"
        ElseIf Cells(linha, 12) < Cells(6, 3) Then
        Cells(linha, 13) = "B"
        ElseIf Cells(linha, 12) < Cells(6, 2) Then
        Cells(linha, 13) = "A"
        Else
        Cells(linha, 13) = "D"
    End If
    ' Calculo da classificação do treinamento BDC
    If Cells(linha, 16) < Cells(4, 4) Then
        Cells(linha, 17) = "C"
        ElseIf Cells(linha, 16) < Cells(4, 3) Then
        Cells(linha, 17) = "B"
        Else
        Cells(linha, 17) = "A"
    End If
    Else
    Cells(linha, 19) = "C"

    ' Final da verificação do BDC
End If
    ' Final das verificações de categoria e classificações


    ' Calculo da classificação final

    If Cells(linha, 11) = "C" Or Cells(linha, 13) = "C" Or Cells(linha, 15) = "C" Or Cells(linha, 17) = "C" Then
        Cells(linha, 19) = "C"
    ElseIf Cells(linha, 11) = "B" Or Cells(linha, 13) = "B" Or Cells(linha, 15) = "B" Or Cells(linha, 17) = "B" Then
        Cells(linha, 19) = "B"
    ElseIf Cells(linha, 11) = "D" And Cells(linha, 13) = "D" And Cells(linha, 15) = "D" And Cells(linha, 17) = "A" Then
        Cells(linha, 19) = "D"
    ElseIf Cells(linha, 11) = "A" Or Cells(linha, 13) = "A" Or Cells(linha, 15) = "A" Or Cells(linha, 17) = "A" Then
        Cells(linha, 19) = "A"
    Else
        Cells(linha, 19) = "O que deu errado?"

    End If
    linha = linha + 1
    ' Fim do calculo da classificação final

Wend

texto = "A classificação foi calculada com sucesso!"
msgbox (texto)

End Sub



