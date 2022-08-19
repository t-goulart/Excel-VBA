Public Function INSS(Valor As Double, Optional Faixa = "", Optional Soma_Salario = False)
'Função para calcular o INSS
'Faixa 1: 7,5% para salário de até 1.100,00
'Faixa 2: 9% para salário de até  de 1.100,01 até 2.203,48
'Faixa 3: 12% para salário de até de 2.203,49 até 3.305,22
'Faixa 4: 14% para salário de até de 3.305,23 até 6.433,57
'----------------------------------------
'Valor: Valor do salário
'Faixa: Altera para a faixa desejada
'Soma_Salario: Soma o salário e o INSS

'Se o campo faixa estiver vazio
If IsEmpty(Faixa) Or Faixa = "" Then
    If Soma_Salario = False Then
        If Valor <= 1000 Then
            QX_INSS = Valor * 0.075
            Exit Function
        ElseIf Valor > 1000 And Valor <= 2203.48 Then
            QX_INSS = Valor * 0.09
            Exit Function
        ElseIf Valor >= 2203.49 And Valor <= 3305.22 Then
            QX_INSS = Valor * 0.12
            Exit Function
        ElseIf Valor >= 3305.23 And Valor <= 6433.57 Then
            QX_INSS = Valor * 0.14
            Exit Function
        End If
    ElseIf Soma_Salario = True Then
        If Valor <= 1000 Then
            QX_INSS = (Valor * 0.075) + Valor
            Exit Function
        ElseIf Valor > 1000 And Valor <= 2203.48 Then
            QX_INSS = (Valor * 0.09) + Valor
            Exit Function
        ElseIf Valor >= 2203.49 And Valor <= 3305.22 Then
            QX_INSS = (Valor * 0.12) + Valor
            Exit Function
        ElseIf Valor >= 3305.23 And Valor <= 6433.57 Then
            QX_INSS = (Valor * 0.14) + Valor
            Exit Function
        End If
    End If
'Se o campo faixa não for nulo ou vazio
ElseIf Not (IsEmpty(Faixa)) Then
    If Soma_Salario = False Then
    Select Case Faixa
        Case 1
            QX_INSS = (Valor * 0.075)
            Exit Function
        Case 2
            QX_INSS = (Valor * 0.09)
            Exit Function
        Case 3
            QX_INSS = (Valor * 0.12)
            Exit Function
        Case 4
            QX_INSS = (Valor * 0.14)
            Exit Function
        Case Else
            QX_INSS = "ESSA FAIXA Ñ EXISTE!"
            Exit Function
    End Select
    ElseIf Soma_Salario = True Then
    Select Case Faixa
        Case 1
            QX_INSS = (Valor * 0.075) + Valor
            Exit Function
        Case 2
            QX_INSS = (Valor * 0.09) + Valor
            Exit Function
        Case 3
            QX_INSS = (Valor * 0.12) + Valor
            Exit Function
        Case 4
            QX_INSS = (Valor * 0.14) + Valor
            Exit Function
        Case Else
            QX_INSS = "ESSA FAIXA Ñ EXISTE!"
            Exit Function
    End Select
    End If
End If
    

End Function