# Tech Challenge Fase 2 - Previsao de Tendencia do IBOVESPA

Projeto desenvolvido para a Pos-Tech FIAP - Data Analytics.

[Abrir notebook no Google Colab](https://colab.research.google.com/github/renancotrin-alt/tech-challenge-fase-2-ibovespa/blob/V1_Ibovespa/notebooks/tech_challenge_fase_2_ibovespa.ipynb)

## Objetivo

Construir um modelo preditivo para estimar se o fechamento do IBOVESPA no proximo pregao sera maior ou menor que o fechamento do pregao atual.

O problema foi tratado como classificacao binaria:

- `1`: fechamento do proximo pregao maior que o fechamento atual.
- `0`: fechamento do proximo pregao menor ou igual ao fechamento atual.

## Dados

A base utilizada contem dados historicos diarios do IBOVESPA obtidos no Investing.com.

- Periodo original: 02/01/2023 a 31/07/2025
- Quantidade original: 644 pregoes
- Arquivo bruto: `data/raw/dados_historicos_ibovespa.csv`
- Base de modelagem apos janelas tecnicas: 594 registros

## Metodologia

O desenvolvimento seguiu as etapas abaixo:

1. Carregamento e limpeza dos dados no formato brasileiro.
2. Ordenacao cronologica dos pregoes.
3. Criacao da variavel alvo `alta_amanha`.
4. Engenharia de atributos usando apenas informacoes disponiveis ate o dia da previsao.
5. Separacao temporal entre treino e teste.
6. Comparacao entre baselines e modelos de classificacao.
7. Avaliacao final no periodo de teste.

O conjunto de teste foi composto pelos ultimos 30 pregoes disponiveis na base de modelagem, evitando embaralhamento aleatorio e respeitando a natureza temporal do problema.

## Features

Foram utilizadas features derivadas do proprio IBOVESPA:

- retornos acumulados e defasados;
- medias moveis e distancia em relacao as medias;
- volatilidade recente;
- amplitude intradiaria e corpo do candle;
- posicao do fechamento em janelas recentes;
- volume transformado;
- indicadores tecnicos simples, como RSI, MACD e Bollinger;
- variaveis de calendario.

## Resultado

O melhor modelo foi um SVC treinado com uma janela recente de 150 pregoes.

- Teste: ultimos 30 pregoes disponiveis
- Periodo de teste: 18/06/2025 a 30/07/2025
- Acuracia final: 76,67%
- Matriz de confusao: `[[16, 2], [5, 7]]`

Esse resultado supera a meta minima de 75% definida no enunciado.

## Estrutura

```text
.
|-- data/
|   |-- raw/
|   `-- processed/
|-- docs/
|-- notebooks/
|-- reports/
|   |-- apresentacao/
|   `-- figures/
|-- src/
|-- README.md
`-- requirements.txt
```

## Como executar

Notebook principal:

`notebooks/tech_challenge_fase_2_ibovespa.ipynb`

Documentacao em formato Word:

`reports/apresentacao/tech_challenge_fase_2_storytelling.docx`

No Google Colab, abra o link abaixo e execute as celulas em sequencia:

https://colab.research.google.com/github/renancotrin-alt/tech-challenge-fase-2-ibovespa/blob/V1_Ibovespa/notebooks/tech_challenge_fase_2_ibovespa.ipynb

Em ambiente local:

```bash
pip install -r requirements.txt
python -m src.data_prep
python -m src.modeling
```

## Observacao

O modelo deve ser interpretado como apoio analitico para leitura de tendencia, nao como recomendacao automatica de compra ou venda. Mercados financeiros sao sensiveis a eventos externos e exigem revisao periodica.
