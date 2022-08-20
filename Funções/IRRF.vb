Public Function IRRF(Valor As Double, Optional Faixa = "", Optional Soma_Salario = False)
'Para mais conteúdos igual a este consulte meu GitHub ou canal no YouTube:
'» GitHub: https://github.com/t-goulart/VBA
'» YouTube: https://www.youtube.com/channel/UC2iemIZz25SIaucByX5FpSw

'1ª faixa: 7,5% para salário base de R$ 1.903,99 a R$ 2.826,65;
'2ª faixa: 15% para salário base de R$ 2.826,66 a R$ 3.751,05;
'3ª faixa: 22,5% para salário base de R$ 3.751,06 a R$ 4.664,68;
'4ª faixa: 27,5% para salário base de a partir de R$ 4.664,69.

'Valor: Valor do salário
'Faixa: Altera para a faixa desejada
'Soma_Salario: Soma o salário e o IRRF

'Se o campo faixa estiver vazio
If IsEmpty(Faixa) Or Faixa = "" Then
    If Soma_Salario = False Then
        If Valor <= 1000 Then
            QX_IRRF = Valor * 0.075
            Exit Function
        ElseIf Valor > 1000 And Valor <= 2203.48 Then
            QX_IRRF = Valor * 0.15
            Exit Function
        ElseIf Valor >= 2203.49 And Valor <= 3305.22 Then
            QX_IRRF = Valor * 0.225
            Exit Function
        ElseIf Valor >= 3305.23 And Valor <= 6433.57 Then
            QX_IRRF = Valor * 0.275
            Exit Function
        End If
    ElseIf Soma_Salario = True Then
        If Valor <= 1000 Then
            QX_IRRF = (Valor * 0.075) + Valor
            Exit Function
        ElseIf Valor > 1000 And Valor <= 2203.48 Then
            QX_IRRF = (Valor * 0.15) + Valor
            Exit Function
        ElseIf Valor >= 2203.49 And Valor <= 3305.22 Then
            QX_IRRF = (Valor * 0.225) + Valor
            Exit Function
        ElseIf Valor >= 3305.23 And Valor <= 6433.57 Then
            QX_IRRF = (Valor * 0.275) + Valor
            Exit Function
        End If
    End If
'Se o campo faixa não for nulo ou vazio
ElseIf Not (IsEmpty(Faixa)) Then
    If Soma_Salario = False Then
    Select Case Faixa
        Case 1
            QX_IRRF = (Valor * 0.075)
            Exit Function
        Case 2
            QX_IRRF = (Valor * 0.15)
            Exit Function
        Case 3
            QX_IRRF = (Valor * 0.225)
            Exit Function
        Case 4
            QX_IRRF = (Valor * 0.275)
            Exit Function
        Case Else
            QX_IRRF = "ESSA FAIXA Ñ EXISTE!"
            Exit Function
    End Select
    ElseIf Soma_Salario = True Then
    Select Case Faixa
        Case 1
            QX_IRRF = (Valor * 0.075) + Valor
            Exit Function
        Case 2
            QX_IRRF = (Valor * 0.15) + Valor
            Exit Function
        Case 3
            QX_IRRF = (Valor * 0.225) + Valor
            Exit Function
        Case 4
            QX_IRRF = (Valor * 0.275) + Valor
            Exit Function
        Case Else
            QX_IRRF = "ESSA FAIXA Ñ EXISTE!"
            Exit Function
    End Select
    End If
End If

End Function
