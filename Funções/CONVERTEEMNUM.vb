Public Function CONVERTEEMNUM(ByVal Celula As String) As Long
'_____ Essa função tem como objetivo converter uma célula com erro de tipo de dado em numero (Long)

If Celula = vbNullString Then 'Se o valor for vazio
    CONVERTEEMNUM = vbNullString 'Retorna vazio também
Else
    Numero_Convertido = CLng(Trim(Celula)) 'Remove espaços vazios e converte em tipo Long
    CONVERTEEMNUM = Numero_Convertido 'Atriu o valor a função
End If

End Function