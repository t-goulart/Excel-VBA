Public Sub Cria_Nova_Aba(Nome As String, Optional Cor_da_Aba)
'____ Cria uma nova aba
'Nome: Nome da nova aba
'Opcional Cor_da_Aba: Pinta a cor da aba

Dim ADicionario 'Variável que será o dicionário de dados
Set ADicionario = CreateObject("Scripting.Dictionary") 'Transforma a variável em um dicionário
ADicionario.CompareMode = TextCompare 'Comparação do tipo texto

For Aba = 1 To Sheets.Count 'Vai iterar entre todas as abas criadas dentro desta pasta de trabalho
    ADicionario.Add Key:=Sheets(Aba).Name, Item:=Aba 'Salva no dicionário o nome e o número da aba atual
Next Aba 'Vai para a próxima aba

If ADicionario.Exists(Nome) = True Then  'Se o Nome da aba que queremos criar existir dentro do dicionário
    MsgBox "A aba " & Nome & " já existe!" & Chr(13) & "Apague ou renomeia o arquivo", vbExclamation, "ATENÇÃO" 'Mensagem informando que a aba já existe
    End 'Encerra a macro | Ñ encerra apenas essa sub rotina, mas toda a macro
Else 'Se a aba que queremos criar não existir dentro do dicionário
        Sheets.Add After:=ActiveSheet 'Insere uma nova aba
        ActiveSheet.Name = Nome  'Nomeia a aba
        
        If VarType(Cor_da_Aba) <> vbError Then 'Se cor for diferente de vazio pinta a aba
            ActiveWorkbook.Sheets(Nome).Tab.Color = Cor_da_Aba 'Pinta a cor da aba
        Else 'Se a cor ñ for preenchida deixa transparente
            ActiveWorkbook.Sheets(Nome).Tab.Color = xlAutomatic 'Cor automatica é transparente
        End If
End If
   
End Sub