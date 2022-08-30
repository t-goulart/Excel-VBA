Public Sub Abrir_Janela_Abrir_Arquivo()
'____ Abre uma janela no Excel para selecionar um arquivo

' Declaração das variáveis
Dim Arquivo As String 

' Captura o arquivo selecionado
Arquivo = Application.GetOpenFilename(, , "Abrir arquivo")
' Abre o arquivo capturado pela variável
Workbooks.Open Arquivo

End Sub