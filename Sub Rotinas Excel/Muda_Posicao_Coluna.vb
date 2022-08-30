Sub Muda_Posicao_Coluna(Recorta_Coluna, Cola_Coluna, Optional Insere_Coluna)
' ____ Muda a posição de uma coluna recortando e colando em outro lugar
' Recorta_Coluna: Recorta a coluna informada
' Cola_Coluna: Nova posição onde a coluna recortada será colada
' Opcional Insere_Coluna: Opcionalmente pode inserir uma nova coluna

' Se o tipo de dado do Insere_Coluna é diferente de 10 | varType 10 significa que é um erro ou valor nulo
If VarType(Insere_Coluna) <> 10 Then 
    ' Insere uma nova coluna à direita da coluna selecionada
    Columns(Insere_Coluna).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove  
End If

' Recorta os dados da coluna e cola na coluna informada
Columns(Recorta_Coluna).Cut Destination:=Columns(Cola_Coluna) 

End Sub