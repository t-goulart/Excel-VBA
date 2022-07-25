Public Function COFINS(Celula, Optional Nova_Aliquota)
'Função que calcula o valor do COFINS
'Valor: Valor do produto/mercadoria/NF
'Resultado: Calculo do COFINS

If VarType(Nova_Aliquota) = vbError Then 'Se estiver vazio
    Valor_Aliquota = 7.6 'Usa a aliquota atual
Else
    Valor_Aliquota = Nova_Aliquota 'Isere a nova aliquota
End If

If Celula = vbNullString Then 'Se a célula estiver vazia retorna vazio
    COFINS = 0
Else
    COFINS = (CLng(Celula) * Valor_Aliquota) / 100
End If

End Function