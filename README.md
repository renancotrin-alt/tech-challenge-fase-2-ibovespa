# Tech Challenge Fase 2 - Previsao do Fechamento do IBOVESPA

Projeto desenvolvido para a Pos-Tech FIAP - Data Analytics.

[Abrir notebook no Google Colab](https://colab.research.google.com/github/renancotrin-alt/tech-challenge-fase-2-ibovespa/blob/V1_Ibovespa/notebooks/tech_challenge_fase_2_ibovespa.ipynb)

## Objetivo

Construir uma solucao de serie temporal para prever o fechamento do IBOVESPA nos proximos pregoes, com foco em regressao e interpretacao da assertividade por erro percentual.

Seguindo a orientacao recebida para a prova substitutiva, a metrica principal foi tratada como:

```text
assertividade = 100 * (1 - WMAPE)
```

O WMAPE mede o erro percentual ponderado. Portanto, quanto menor o WMAPE, maior a assertividade da previsao.

## Dados

A base utilizada contem dados historicos diarios do IBOVESPA obtidos no Investing.com.

- Periodo bruto: 04/01/2010 a 22/07/2026
- Quantidade bruta: 4.105 pregoes
- Arquivo bruto: `data/raw/dados_historicos_ibovespa.csv`
- Fonte: Investing.com - dados historicos do indice Bovespa

O arquivo bruto foi mantido no formato brasileiro original, com pontos para milhar e virgulas para decimais percentuais. A limpeza e feita no codigo para evitar perda de informacao em valores como `175.831`, que representam 175.831 pontos.

## Metodologia

O desenvolvimento seguiu as etapas abaixo:

1. Carregamento e limpeza dos dados historicos.
2. Analise exploratoria da serie do IBOVESPA.
3. Decomposicao da serie temporal em tendencia, sazonalidade e residuo.
4. Teste de estacionariedade com ADF.
5. Separacao temporal entre treino e teste, sem embaralhamento.
6. Comparacao entre baseline de persistencia, ARIMA estatico e ARIMA walk-forward.
7. Avaliacao por WMAPE, assertividade, MAE e acerto dentro de uma faixa de 2,15%.
8. Teste complementar de modelos com features tecnicas.
9. Previsao dos proximos 15 pregoes.

O conjunto de teste principal foi composto pelos ultimos 125 pregoes disponiveis com alvo conhecido, cobrindo 20/01/2026 a 21/07/2026.

## Resultado Principal

O modelo escolhido para a narrativa final foi o ARIMA(5,1,0) em avaliacao walk-forward. Essa configuracao respeita a ordem temporal dos dados e recalibra o modelo a cada novo pregao do teste, simulando melhor o uso pratico.

| Modelo | WMAPE | Assertividade | MAE | Acerto ate 2,15% |
| --- | ---: | ---: | ---: | ---: |
| Naive - persistencia | 0,95% | 99,05% | 1.722 pontos | 91,20% |
| ARIMA(5,1,0) walk-forward | 1,33% | 98,67% | 2.411 pontos | 82,40% |
| ARIMA(5,1,0) estatico | 8,84% | 91,16% | 15.979 pontos | 2,40% |

O baseline de persistencia ficou muito forte, como e comum em series financeiras de curto prazo. Mesmo assim, o ARIMA walk-forward tambem supera a meta de 80% quando a assertividade e calculada por `1 - WMAPE` e fica acima de 80% na metrica operacional de erro dentro da faixa de 2,15%.

## Analise Complementar

Tambem foram avaliados modelos supervisionados com features tecnicas e uma classificacao de direcao (alta ou baixa no proximo pregao). Essa parte foi mantida como complemento analitico, pois o enunciado atual e a orientacao recebida apontam para serie temporal/regressao e uso de metricas de distancia.

## Previsao dos Proximos 15 Pregoes

Com a base atualizada ate 22/07/2026, a previsao de 15 pregoes inicia em 23/07/2026 e segue ate 12/08/2026. O notebook apresenta a tabela completa com fechamento previsto e intervalo de confianca de 95%.

## Estrutura

```text
.
|-- data/
|   |-- raw/
|   `-- processed/
|-- notebooks/
|-- reports/
|   |-- apresentacao/
|   `-- figures/
|-- scripts/
|-- src/
|-- README.md
`-- requirements.txt
```

## Como Executar

Notebook principal:

`notebooks/tech_challenge_fase_2_ibovespa.ipynb`

Em ambiente local:

```bash
pip install -r requirements.txt
python -m src.data_prep
python -m src.forecasting
```

No Google Colab, abra o link do inicio deste README e execute as celulas em sequencia.

## Observacao

O modelo deve ser interpretado como apoio analitico para estudo de serie temporal, nao como recomendacao automatica de compra ou venda. Mercados financeiros sao sensiveis a eventos externos, revisoes de dados e mudancas de regime.
