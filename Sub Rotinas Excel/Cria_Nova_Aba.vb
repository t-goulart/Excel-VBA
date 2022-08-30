Public Sub Cria_Nova_Aba(Nome As String, Optional Cor_da_Aba)
' ____ Cria uma nova aba (Microsoft Excel)
' Nome: Nome da nova aba
' Opcional Cor_da_Aba: Pinta a cor da aba

' Declaração das variáveis
Dim Dict

' Transforma a variável em um dicionário
Set Dict = CreateObject("Scripting.Dictionary")
' Comparação do tipo texto
Dict.CompareMode = TextCompare

' Vai iterar entre todas as abas criadas dentro desta pasta de trabalho
For Aba = 1 To Sheets.Count 
    ' Salva no dicionário o nome e o número da aba atual
    Dict.Add Key:=Sheets(Aba).Name, Item:=Aba
' Vai para a próxima aba
Next Aba

' Se o Nome da aba que queremos criar existir dentro do dicionário
If Dict.Exists(Nome) = True Then  
    ' Mensagem informando que a aba já existe
    MsgBox "A aba " & Nome & " já existe!" & Chr(13) & "Apague ou renomeia o arquivo", vbExclamation, "ATENÇÃO" 
    ' Encerra a macro | Ñ encerra apenas essa sub rotina, mas toda a macro
    End 

Else ' Se a aba que queremos criar não existir dentro do dicionário

    ' Insere uma nova aba
    Sheets.Add After:=ActiveSheet
    ' Nomeia a aba
    ActiveSheet.Name = Nome
        
    ' Se cor for diferente de vazio pinta a aba
    If VarType(Cor_da_Aba) <> vbError Then 
        ' Pinta a cor da aba
        ActiveWorkbook.Sheets(Nome).Tab.Color = Cor_da_Aba
    ' Se a cor ñ for preenchida deixa transparente
    Else
        ' Cor automatica é transparente
        ActiveWorkbook.Sheets(Nome).Tab.Color = xlAutomatic
    End If
End If
   
End Sub