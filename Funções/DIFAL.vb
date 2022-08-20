Public Function DIFAL(Valor As Double, Estado_Origem As String, Estado_Destino As String)
'Para mais conteúdos igual a este consulte meu GitHub ou canal no YouTube:
'» GitHub: https://github.com/t-goulart/VBA
'» YouTube: https://www.youtube.com/channel/UC2iemIZz25SIaucByX5FpSw

'---- Função para calcular a diferença de ICMS entre estados
'---- Estados do Sul e Sudeste (Exceto Espirito Santo): 12%
'---- Demais estados (Incluindo Espirito Santo) 7%
' Valor = Valor do produto/mercadoria/NF
' Estado_Origem = Estado de Origem do DIFAL. Ex.: "SP" para São Paulo
' Estado_Destino = Estado de Destino do DIFAL. Ex.: "RJ" para Rio de Janeiro
' Resultado = Calculo da alíquota do Estado Destino - Calculo da alíquota do Estado Origem


'Variável que será o dicionário de dados do DIFAL
Dim Dict As New Dictionary
Dim Aliquota_Destino As Variant
Dim Aliquota_Origem As Double
Dim ICMS_Estado_Origem As Double
Dim ICMS_Estado_Destino As Double

'---- Identifica o estado de origem e captura o valor da alíquota
Select Case Estado_Origem
    Case "MG" 'Minas Gerais
        Aliquota_Origem = 0.12
    Case "PR" 'Paraná
        Aliquota_Origem = 0.12
    Case "RJ" 'Rio de Janeiro
        Aliquota_Origem = 0.12
    Case "RS" 'Rio Grande do Sul
        Aliquota_Origem = 0.12
    Case "SC" 'Santa Catarina
        Aliquota_Origem = 0.12
    Case "SP" 'São Paulo
        Aliquota_Origem = 0.12
    Case Else 'Todos os demais estados
        Aliquota_Origem = 0.07
End Select

'---- Cria o dicionário que vai conter o valor percentual da taxa do DIFAL por estado
With Dict
    .CompareMode = TextCompare
    'Adiciona elementos ao dicionário
    'Datas são as Keys | Descrição são os Items
    .Add Key:="AC", Item:="0.17" 'Acre 17%
    .Add Key:="AL", Item:="0.18" 'Alagoas 18%
    .Add Key:="AP", Item:="0.18" 'Amapá 18%
    .Add Key:="AM", Item:="0.18" 'Amazonas 18%
    .Add Key:="BA", Item:="0.18" 'Bahia 18%
    .Add Key:="CE", Item:="0.18" 'Ceará 18%
    .Add Key:="DF", Item:="0.18" 'Distrito Federal 18%
    .Add Key:="ES", Item:="0.17" 'Espirito Santo 17%
    .Add Key:="GO", Item:="0.17" 'Goiás 17%
    .Add Key:="MA", Item:="0.18" 'Maranhão  18%
    .Add Key:="MT", Item:="0.17" 'Mato Grosso 17%
    .Add Key:="MS", Item:="0.17" 'Mato Grosso do Sul 17%
    .Add Key:="MG", Item:="0.18" 'Minas Gerais 18%
    .Add Key:="PA", Item:="0.17" 'Pará 17%
    .Add Key:="PB", Item:="0.18" 'Paraíba 18%
    .Add Key:="PR", Item:="0.18" 'Paraná 18%
    .Add Key:="PE", Item:="0.18" 'Pernambuco 18%
    .Add Key:="PI", Item:="0.18" 'Piauí 18%
    .Add Key:="RR", Item:="0.175" 'Roraima 17,5%
    .Add Key:="RO", Item:="0.17" 'Rondônia 17%
    .Add Key:="RJ", Item:="0.20" 'Rio de Janeiro 20%
    .Add Key:="RN", Item:="0.18" 'Rio Grande do Norte 18%
    .Add Key:="RS", Item:="0.18" 'Rio Grande do Sul 18%
    .Add Key:="SC", Item:="0.17" 'Santa Catarina 17%
    .Add Key:="SP", Item:="0.18" 'São Paulo 18%
    .Add Key:="SE", Item:="0.18" 'Sergipe 18%
    .Add Key:="TO", Item:="0.18" 'Tocantins 18%
End With

'---- Identifica o estado de destino e captura o valor da alíquota
' Replace - Substitui o ponto pela virgula para fazer o calculo
' UCase - Deixa todas as letras maiúsculas
Aliquota_Destino = Replace(Dict.Item(UCase(Estado_Destino)), ".", ",")

' Valor multiplicado pela alíquota do Estado de Origem
ICMS_Estado_Origem = Valor * Aliquota_Origem
' Valor multiplicado pela alíquota do Estado de Destino
ICMS_Estado_Destino = Valor * Aliquota_Destino

'---- Calculo do DIFAL
' Resultado = Valor do Estado Destino - Valor do Estado Origem
DIFAL = Round(ICMS_Estado_Destino - ICMS_Estado_Origem, 2)

End Function