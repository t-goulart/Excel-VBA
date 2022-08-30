Public Sub Renomeia_Aba(Novo_Nome As String)
' ____ Renomeia uma aba (Microsoft Excel)
' Novo_Nome: Novo nome para a aba existente

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
If Dict.Exists(Novo_Nome) = True Then

    ' Mensagem informando que a aba já existe
    MsgBox "A aba " & Novo_Nome & " já existe!" & Chr(13) & "Apague ou renomeia o arquivo", vbExclamation "ATENÇÃO" 
    ' Encerra a macro | Ñ encerra apenas essa sub rotina, mas toda a macro
    End 

Else ' Se a aba que queremos criar não existir dentro do dicionário

    ' Se a variável é diferente de vazio
    If Novo_Nome <> vbNullString Then
        ' Novo_Nome a aba atual 
        ActiveSheet.Name = Novo_Nome 
    End If

End If

End Sub
