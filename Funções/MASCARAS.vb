Public Function MASCARAS(Celula As Variant, Optional ByVal Tipo_Doc As String) As String
    ' Função que aplica uma máscara a um valor, baseado no tipo de documento informado
    ' Celula: Valor que será formatado conforme a máscara
    ' Tipo_Doc: Tipo de documento que define a máscara a ser usada
    
    ' Declaração das variáveis
    Dim Documento As Variant ' Armazena o valor convertido em array
    Dim c As Integer ' Variável de controle para percorrer a máscara
    Dim p As Integer ' Variável de controle para percorrer os dígitos do documento
    Dim Separador As Boolean ' Identifica se o caractere atual na máscara é um separador (como "." ou "-")
    Dim Mascara As String ' Armazena a máscara correspondente ao tipo de documento
    Dim NumEsperado As Integer ' Define o número esperado de caracteres para o tipo de documento
    Dim NumFaltando As Integer ' Calcula a quantidade de caracteres faltantes quando insuficientes

    ' Seleciona a máscara com base no tipo de documento informado
    Select Case UCase(Trim(Tipo_Doc)) ' Converte o tipo de documento para maiúsculas e remove espaços extras
        Case Is = "RG": Mascara = "##.###.###-#": NumEsperado = 9 ' Máscara e número de caracteres esperados para RG
        Case Is = "PIS": Mascara = "##.#####.##-#": NumEsperado = 11 ' Máscara para PIS
        Case Is = "CPF": Mascara = "###.###.###-##": NumEsperado = 11 ' Máscara para CPF
        Case Is = "CTPS": Mascara = "###### ###-##": NumEsperado = 12 ' Máscara para CTPS
        Case Is = "CNPJ": Mascara = "##.###.###/####-##": NumEsperado = 14 ' Máscara para CNPJ
        Case Is = "TE": Mascara = "#### #### ####": NumEsperado = 12 ' Máscara para Título Eleitoral
        Case Is = "RESERVISTA": Mascara = "###### #": NumEsperado = 7 ' Máscara para Reservista
        Case Else ' Caso o tipo de documento informado seja inválido
            MASCARAS = "Tipo de documento inválido." ' Mensagem de erro para tipo não reconhecido
            Exit Function ' Interrompe a execução da função
    End Select

    ' Verifica se o número de caracteres do valor é insuficiente
    If Len(Celula) < NumEsperado Then
        NumFaltando = NumEsperado - Len(Celula) ' Calcula a quantidade de caracteres faltantes
        MASCARAS = "Quantidade de caracteres insuficientes. Eram esperados " & NumEsperado & " caracteres. Faltam " & NumFaltando & " caracteres." ' Mensagem de erro
        Exit Function ' Interrompe a execução da função
    End If

    ' Converte o valor da célula para Unicode e remove espaços extras
    Documento = StrConv(Trim(Celula), vbUnicode) ' Converte o valor para Unicode para facilitar o processamento
    Documento = Split(Left(Documento, Len(Documento) - 1), vbNullChar) ' Divide o valor em um array, ignorando caracteres nulos

    ' Percorre cada caractere da máscara
    For c = 1 To Len(Mascara)
        If InStr(Mid(Mascara, c, 1), "#") Then ' Verifica se o caractere atual na máscara é um marcador (#)
            Mid(Mascara, c, 1) = Documento(p) ' Substitui o marcador (#) pelo dígito correspondente do documento
            p = p + 1 ' Avança para o próximo dígito do documento
        End If
    Next c ' Avança para o próximo caractere da máscara

    ' Retorna o valor formatado com a máscara aplicada
    MASCARAS = Mascara
End Function
