Public Function ICMS(Valor As Double, Estado As Variant) As Variant
'Função para calcular o valor do ICMS por Estado
'Valor: Valor do produto/mercadoria/NF
'Estado: Sigla do Estado de origem. Ex.: "SP" para São Paulo
'Resultado: Calculo da alíquota do ICMS com base no Estado

On Error Resume Next

Dim varTaxa As Double
'Variável que será o dicionário de dados do ICMS
Dim LDicionarioICMS

'_____ Criação do dicionário
'Transforma a variável em um dicionário
Set LDicionarioICMS = CreateObject("Scripting.Dictionary")
With LDicionarioICMS
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

'Retorna um item do dicionário baseado no estado
varTaxa = Replace(LDicionarioICMS.Item(UCase(Estado)), ".", ",")
ICMS = Round(CDbl(Valor) * varTaxa, 2)

If Err.Number = 13 Then
    ICMS = "ESTADO INVÁLIDO"
End If

End Function