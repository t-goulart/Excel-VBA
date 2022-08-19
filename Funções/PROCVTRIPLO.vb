Public Function PROCVTRIPLO(Valores_Procurados As Variant, Primeira_Coluna As Range, Segunda_Coluna As Range, Terceira_Coluna As Range, Matriz_Tabela As Range, Coluna_Resultado)
'____ Funciona como INDICE + CORRESP limitado a 3 valores procurados
'Valores_Procurados: Recebe três valores concatenados formando um valor único
'Primeira_Coluna: Coluna que contém a primeira palavra das três contidas dentro de Valores_Procurados
'Segunda_Coluna: Coluna que contém a segunda palavra das três contidas dentro de Valores_Procurados
'Segunda_Coluna: Coluna que contém a terceira palavra das três contidas dentro de Valores_Procurados
'Matriz_Tabela: Tabela que contém as duas colunas procuradas e a coluna com o resultado procurado
'Coluna_Resultado: Coluna que contém o resultado da busca

'Inicia na primeira linha do Range de cada coluna (Primeira_Coluna, Segunda_Coluna e Terceira_Coluna) | 'Ñ é a primeira linha da planilha
Linha = 1

'Remove possíveis espaços vazios
Valores_Procurados = Trim(Valores_Procurados)

'Itera sobre a primeira coluna | Desta forma vai limitar a qtde de linhas que o laço vai ler
For Each Elemento In Primeira_Coluna
    
    '_____Valida se o dado informado por alguma das 3 colunas é tipo data | Quando datas são concatenas automaticamente se tornando numero (Ex.:01/02/2020 = 43862)
    'Se a linha informada da primeira coluna é tipo data
    If VarType(Primeira_Coluna.Rows(Linha)) = vbDate Then
        'Concatena o valor contido na linha de cada uma das duas colunas (ranges) informados | Converte o primeiro valor em tipo long (inteiro)
        Concatena_Colunas = WorksheetFunction.Concat(CLng(Primeira_Coluna.Rows(Linha)), Trim(Segunda_Coluna.Rows(Linha)), Trim(Terceira_Coluna.Rows(Linha)))
    'Se a linha informada da segunda coluna é tipo data
    ElseIf VarType(Segunda_Coluna.Rows(Linha)) = vbDate Then
        'Concatena o valor contido na linha de cada uma das duas colunas (ranges) informados | Converte o segundo valor em tipo long (inteiro)
        Concatena_Colunas = WorksheetFunction.Concat(Trim(Primeira_Coluna.Rows(Linha)), CLng(Segunda_Coluna.Rows(Linha)), Trim(Terceira_Coluna.Rows(Linha)))
    'Se a linha informada da terceira coluna é tipo data
    ElseIf VarType(Terceira_Coluna.Rows(Linha)) = vbDate Then
        'Concatena o valor contido na linha de cada uma das duas colunas (ranges) informados | Converte o terceiro valor em tipo long (inteiro)
        Concatena_Colunas = WorksheetFunction.Concat(Trim(Primeira_Coluna.Rows(Linha)), Trim(Segunda_Coluna.Rows(Linha)), CLng(Terceira_Coluna.Rows(Linha)))
    'Se nenhuma coluna é tipo data então...
    Else
        'Concatena o valor contido na linha de cada uma das duas colunas (ranges) informados | Ñ converte nenhum dos valores
        Concatena_Colunas = WorksheetFunction.Concat(Trim(Primeira_Coluna.Rows(Linha)), Trim(Segunda_Coluna.Rows(Linha)), Trim(Terceira_Coluna.Rows(Linha)))
    End If
    
    '____ Faz a remoção de pontos e virgulas das variáveis antes de comparar
    'Se tiver ponto na variável remove
    If InStr(1, Valores_Procurados, ".", vbBinaryCompare) > 0 Then
        Valores_Procurados = Replace(Valores_Procurados, ".", "")
    'Se tem virgula na variável remove
    ElseIf InStr(1, Valores_Procurados, ",", vbBinaryCompare) > 0 Then
        Valores_Procurados = Replace(Valores_Procurados, ",", "")
    End If
    'Se tiver ponto na variável remove
    If InStr(1, Concatena_Colunas, ".", vbBinaryCompare) > 0 Then
        Concatena_Colunas = Replace(Concatena_Colunas, ".", "")
    'Se tem virgula na variável remove
    ElseIf InStr(1, Concatena_Colunas, ",", vbBinaryCompare) > 0 Then
        Concatena_Colunas = Replace(Concatena_Colunas, ",", "")
    End If
    
    
    'Se os valores informados são iguais a concatenação da linha atual das três colunas informadas
    If Valores_Procurados = Concatena_Colunas Then
        'Procura na matriz informada e retorna a sua posição exata com base na linha e coluna (Coluna que contém o resultado desejado)
        PROCVTRIPLO = Matriz_Tabela.Cells(Linha, Coluna_Resultado)
        'Sai do laço
        Exit For
    End If
    'Soma o valor da variável +1
    Linha = Linha + 1
Next

'Se tiver dados duplicados retorna apenas a primeira ocorrência

End Function