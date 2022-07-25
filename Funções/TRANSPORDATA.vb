Public Function TRANSPORDATA(Celula, Optional Opcao As String) As Variant
'Função que separa data e hora de uma célula
'Resultado: Retorna a data, hora ou dia da semana

Lista = Split(Trim(Celula), " ") 'Lista que vai ter os dados contidos na célula
Opcao_Data = UCase(Opcao)

On Error GoTo Erro

If Opcao_Data = vbNullString Then 'Se ñ escolher uma das opções retorna data
    Resultado = CDate(Format(Lista(0), "dd/MM/yyyy"))
ElseIf Opcao_Data = "D" Then 'Retorna a posição 0 (Data)
    Resultado = CDate(Format(Lista(0), "dd/MM/yyyy"))
ElseIf Opcao_Data = "H" Then 'Retorna a posição 1 (Hora)
    Resultado = Lista(1)
ElseIf Opcao_Data = "S" Then 'Retorna o dia da semana da data
    Resultado = StrConv(Format(Lista(0), "dddd"), vbProperCase) 'Deixa apenas a primeira letra maiúscula
Else
    QX_TRANSPORDATA = "Segunda opção inválida"
End If

QX_TRANSPORDATA = Resultado

Erro:
    If Err.Number <> 0 Then QX_TRANSPORDATA = Err.Description


End Function