Sub RetornaChaveEValores()
'***** IMPORTANTE!!! Primeiro habilitar em Ferramentas > Referências > Microsoft Scripting Runtime

'---- Essa função vai usar algum dado como base para fazer a soma de valores e retornar para o usuário
'---- Esses valores de base podem ser Notas Fiscais, Fornecedores, Cliente e etc.
'---- Uma das grandes vantagens será que o usuário não precisa saber, por exemplo, o número das NFs.
'---- Essa rotina vai identificar sozinha e retornar sem duplicar as NFs e a soma dos seus respectivos valores


' Declaração das variáveis
Dim Dict As New Dictionary
Dim Matriz As Variant
Dim Nome_Aba As String
Dim Matriz_Linha_Inicial As Integer
Dim Matriz_Coluna_Inicial As Integer
Dim Numero_Coluna_Chave As Integer
Dim Numero_Coluna_Valores As Integer
Dim Destino_Numero_Linha As Integer
Dim Destino_Numero_Coluna As Integer
Dim Matriz_Celula_Inicial As Variant


'---- Coleta dos dados que serão usados para identificar as NFs
' » Nome_Aba = Aba que contém a base dos dados | Pode ser necessário consultar dados de uma aba diferente
Nome_Aba = Application.InputBox("Digite o nome da aba que contém os dados", "Origem dos dados")
' » Matriz_Linha_Inicial = Informar o número da linha que inicia a tabela
'Obs.: Type = 1 significa que a caixa de entrada só aceita números
Matriz_Linha_Inicial = Application.InputBox("Informe o número da célula que inicia a tabela", "Origem dos dados", "Digite apenas números", Type:=1)
' » Matriz_Coluna_Inicial = Informar o número da coluna que inicia a tabela
Matriz_Coluna_Inicial = Application.InputBox("Informe o número da coluna que inicia a tabela", "Origem dos dados", "Digite apenas números", Type:=1)
' » Numero_Coluna_Chave = Informar o número da coluna que contém a NF ou outro dado que possa ser usado como chave
Numero_Coluna_Chave = Application.InputBox("Informe o numero da coluna que contém os dados que serão usados como referência", "Origem dos dados", "Digite apenas números", Type:=1)
' » Numero_Coluna_Valores = Coluna que contém os valores que queremos somar para descobrir o valor total da NF
Numero_Coluna_Valores = Application.InputBox("Informe o numero da coluna que contém os valores que serão somados", "Origem dos dados", "Digite apenas números", Type:=1)


' » Matriz_Celula_Inicial = Célula que o Currentregion deve iniciar a varredura e salvar seus respectivos dados
Matriz_Celula_Inicial = Cells(Matriz_Linha_Inicial, Matriz_Coluna_Inicial).Address

'---- Identificando a fonte dos dados que vamos precisar para identificar o valor da NF
'A partir da célula informada faça uma varredura da região na aba informa e atribua os valores à variável matriz
'matriz = Sheets(Nome_Aba).Cells(Linha_Matriz, Coluna_Matriz).CurrentRegion.Value
Matriz = Sheets(Nome_Aba).Range(Matriz_Celula_Inicial).CurrentRegion.Value

'---- Laço para iterar sobre cada elemento (célula) da nossa matriz
'LBound indica o início de uma matriz | UBound indica o fim de uma matriz
' +1 vai pular a linha dos títulos
For i = LBound(Matriz) + 1 To UBound(Matriz)
    'Atribui ao dicionário a chave e o valor se não existir | Se já existir vai acumular o valor existente somado ao novo valor
    Dict(Matriz(i, Numero_Coluna_Chave)) = Dict(Matriz(i, Numero_Coluna_Chave)) + Matriz(i, Numero_Coluna_Valores)
Next

'---- Salvando os dados que foram coletados no local indicado pelo usuário
' As variáveis abaixo recebem o número da coluna e da linha que os dados devem ser salvos
' A partir da coluna + linha informados o laço salvará os dados
Destino_Numero_Coluna = Application.InputBox("Informe o número da coluna que os dados devem começar a ser salvos", "Destino dos dados", "Digite apenas números", Type:=1)
Destino_Numero_Linha = Application.InputBox("Informe o número da linha que os dados devem começar a ser salvos", "Destino dos dados", "Digite apenas números", Type:=1)

' Laço que vai iterar sobre cada uma das chaves do dicionário
For Each Chave In Dict
    'Salva na célula a chave
    Cells(Destino_Numero_Linha, Destino_Numero_Coluna) = Chave
    'Salva na coluna ao lado da célula o valor atribuido a chave informada
    Cells(Destino_Numero_Linha, Destino_Numero_Coluna).Offset(0, 1) = Dict(Chave)
    'Soma a linha atual +1 para salvar os dados na linha de baixo
    Destino_Numero_Linha = Destino_Numero_Linha + 1
Next


End Sub
