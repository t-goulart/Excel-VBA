Public Function PIS(Celula, Optional Nova_Aliquota)
'Função que calcula o valor do PIS
'Valor: Valor do produto/mercadoria/NF
'Resultado: Calculo do PIS

If VarType(Nova_Aliquota) = vbError Then 'Se estiver vazio
    Valor_Aliquota = 1.65 'Usa a aliquota atual
Else
    Valor_Aliquota = Nova_Aliquota 'Isere a nova aliquota
End If

If Celula = vbNullString Then 'Se a célula estiver vazia retorna vazio
    PIS = vbNullString
Else
    PIS = (CLng(Celula) * Valor_Aliquota) / 100
End If

End Function