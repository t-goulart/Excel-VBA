Public Function NUMNOTA(ByVal Celula As String, Optional Tipo As String) As String
'____ Extrai um numero de uma cadeia de caracteres usando como base uma palavra chave _
        'Simplificando: Extrair um número de uma nota fiscal de uma célula cheia de informações. Ex.: _
        'Célula = IPI S/ DEV NF 000032554 CF NF 000033235| RESULTADO = 000033235 (RETORNA SEMPRE A 2ª NF)
        
'Celula: Celula onde tem o número da NF
'Tipo: Palavra chave que será usada como parâmetro para busca do documento (NF, Nota, Num, DOC e etc)
        
Application.Volatile 'Atualiza o cálculo automaticamente

Dim Qtde_Caracteres As Long 'Variável que vai receber a qtde de caracteres contidos na célula
Dim Controle As Boolean 'Variável que será usada para entrar em uma condição específica

Celula = UCase(Trim(Celula)) 'Remove espaços vazios excedentes para evitar erro | Deixa todas as letras em maiúsculo para evitar problemas case sensitive
Qtde_Caract = Len(Celula) 'Conta a quantidade de caracteres
Controle = False

Tipo_Busca = UCase(Tipo) 'Por padrão será maiúsculo para evitar problemas com case sensitive
If Tipo = vbNullString Then Tipo_Busca = "NF" 'Se não informar a palavra chave, por padrão será NF

For Numero = 1 To Qtde_Caract 'Para cada caractere (um a um) identifica se é número ou texto
    If IsNumeric(Mid(Celula, Numero, 1)) Then 'Verifica se o caractere atual é um número | Se sim entra nessa condição
        If Controle = True And Retira_Numeros <> vbNullString Then  'Se a variável de controle é true então o caractere anterior era uma letra
            Retira_Numeros = Retira_Numeros + "/" 'Insere a barra para separar os numeros
        End If
        Controle = False 'Se o caractere é um número permanece como false
        Retira_Numeros = Retira_Numeros & Mid(Celula, Numero, 1) 'Vai concatenando os valores até formar o numero da NF
    Else
        Controle = True 'Se o caractere ñ é um número muda a variável para True
    End If
Next

Lista_Numeros = Split(Retira_Numeros, "/") 'Criar um array com cada uma das notas
Posicao_NF = InStrRev(Celula, Tipo_Busca) 'Retorna a posição da palavra NF da direita para a esquerda
Posicao_Elemento = 0 'Variável que inicia a posição dos numeros do array (Array sempre inicia pelo 0)

If Posicao_NF > 0 Then 'Se encontrar a palavra NF faz a função
    For Each Nota In Lista_Numeros 'Itera sobre os números do célula
        Posicao_Numero = InStr(Celula, Nota) 'Retorna a posição do número correspondente a NF
        If Posicao_NF < Posicao_Numero Then 'Se a posição da palavra NF é menor que o PRIMEIRO NÚMERO, então está na frente do primeiro número. EX.: NF 000123 P 123
            Nota_Fiscal = Lista_Numeros(Posicao_Elemento) 'Salva o primeiro número dentro da variavel
            Exit For 'Sai do laço | Só precisamos comparar a posição do primeiro número
        Else
            Posicao_Elemento = Posicao_Elemento + 1 'Vai para o próximo numero pois o atual está em uma posição inferir a palavra NF
        End If
    Next
Else
    Nota_Fiscal = 0 'Se não encontrar a palavra chave retorna 0
End If

NUMNOTA = Nota_Fiscal

End Function

