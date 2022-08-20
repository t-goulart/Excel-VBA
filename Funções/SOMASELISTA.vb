Public Function SOMASELISTA(Lista_Dados, Range_Busca_Dados As Range, Range_Busca_Soma As Range)
'Para mais conteúdos igual a este consulte meu GitHub ou canal no YouTube:
'» GitHub: https://github.com/t-goulart/VBA
'» YouTube: https://www.youtube.com/channel/UC2iemIZz25SIaucByX5FpSw

'_____ Essa função faz a soma de vários valores de uma vez usando como base uma lista de valores procurados
'Lista_Dados: Dados que serão procurados no Range_Busca_Dados para fazer a soma
'Range_Busca_Dados: Coluna que tem as informações que vamos usar para comparação
'Range_Busca_Soma: Coluna que contém os valores para soma

Total_Soma = 0
Soma_Com_Base_Nos_Dados = 0

'Itera sobre a lista de valores procurados
For Each dado In Lista_Dados
    'Faz a soma de cada um dos valores procurados e salva na variável
    Soma_Com_Base_Nos_Dados = WorksheetFunction.SumIfs(Range_Busca_Soma, Range_Busca_Dados, dado)
    'Vai somando os valores e inserindo na variável total
    Total_Soma = Total_Soma + Soma_Com_Base_Nos_Dados
Next

'Atribui a função o valor total da soma após encerrar o laço
SOMASELISTA = Total_Soma

End Function
