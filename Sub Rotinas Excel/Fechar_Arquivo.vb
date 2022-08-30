Public Sub Fechar_Arquivo(Nome_Arquivo)
' ____ Fecha uma pasta de trabalho (Ñ é uma aba ou planilha é uma pasta de trabalho)
' Nome_Arquivo: Nome do arquivo que será fechado

' Fecha o arquivo sem salvar qualquer alteração
Workbooks(Nome_Arquivo).Close SaveChanges:=False 

End Sub