Public Function MASCARAS(ByVal Celula As String, Optional ByVal Tipo_Doc As String)
'Para mais conteúdos igual a este consulte meu GitHub ou canal no YouTube:
'» GitHub: https://github.com/t-goulart/VBA
'» YouTube: https://www.youtube.com/channel/UC2iemIZz25SIaucByX5FpSw

'_____ Função que faz a formatação de alguns documentos com base no valor informado
'Célula: Valor que será editado com base na mascara informada
'Tipo_Doc: Descrição do documento que a mascara deve formatar

Dim Documento As Variant
Dim c As Integer 'Variavel de controle
Dim p As Integer 'Posição no array
Dim Separador As Boolean 'Variável que identifica se existe um separados (ponto, traço, virgula ou barra)

Select Case UCase(Trim(Tipo_Doc)) 'Selecione a mascara do tipo_doc
    Case Is = "RG": Mascara = "##.###.###-#" 'Máscara do RG
    Case Is = "PIS": Mascara = "##.#####.##-#" 'Máscara do PIS
    Case Is = "CPF": Mascara = "###.###.###-##" 'Máscara do CPF
    Case Is = "CTPS": Mascara = "###### ###-##" 'Máscara do CTPS
    Case Is = "CNPJ": Mascara = "##.###.###/####-##" 'Máscara do CNPJ
    Case Is = "TE": Mascara = "#### #### ####" 'Máscara do TITULO ELEITORAL
    Case Is = "RESERVISTA": Mascara = "###### #" 'Máscara do RESERVISTA
End Select

Documento = StrConv(Trim(Celula), vbUnicode) 'Converte os valores em Unicode

'Converte em um array, separa os valores, ignora os vazios e o ultimo caractere (vazio)
'vbNullChar identifica os valores nulos ou vazios que o Split vai ignorar
Documento = Split(Left(Documento, Len(Documento) - 1), vbNullChar)

For c = 1 To Len(Mascara) 'Itera sobre os caracteres da Mascara
    
    If InStr(Mid(Mascara, c, 1), "#") Then 'Se o valor da máscara na posição atual for igual a # (Diferente dos separados)
        Mid(Mascara, c, 1) = Documento(p) 'Mascara recebe o número do Doc que corresponde a posição do #
    Else
        Separador = True 'Se tiver ponto, barra, virgula e etc seja true
    End If
    
    'Se o Separador for True | Atribui False para Separador | Señ a variável P soma P+1 e atribui False para Separador
    If Separador Then Separador = False Else p = p + 1: Separador = False

Next c 'Vai para a próxima posição da Mascara

MASCARAS = Mascara 'Recebe o resultado da Máscara que foi substituido os # pelos valores

End Function