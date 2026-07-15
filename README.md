# Tech Challenge Fase 2 - Previsao de Tendencia do IBOVESPA

Projeto desenvolvido para a Pos-Tech FIAP - Data Analytics.

## Objetivo

Construir um modelo preditivo para estimar se o fechamento do IBOVESPA no proximo pregao sera maior ou menor que o fechamento do pregao atual.

O problema sera tratado como uma classificacao binaria:

- `1`: fechamento do proximo pregao maior que o fechamento atual.
- `0`: fechamento do proximo pregao menor ou igual ao fechamento atual.

## Dados

A base utilizada contem dados historicos diarios do IBOVESPA obtidos no Investing.com.

- Periodo disponivel: 02/01/2023 a 31/07/2025
- Quantidade de registros: 644 pregoes
- Arquivo original no projeto: `data/raw/dados_historicos_ibovespa.csv`

## Estrategia

O desenvolvimento sera organizado em etapas:

1. Carregamento e limpeza dos dados.
2. Analise exploratoria.
3. Criacao da variavel alvo.
4. Engenharia de atributos com informacoes disponiveis ate o dia da previsao.
5. Separacao temporal entre treino e teste, mantendo os ultimos 30 pregoes como teste final.
6. Comparacao entre baseline e modelos de classificacao.
7. Analise das metricas e interpretacao gerencial dos resultados.

## Resultado atual

O melhor modelo testado ate o momento foi um SVC com janela recente de 150 pregoes para treino.

- Teste: ultimos 30 pregoes disponiveis na base
- Periodo de teste: 18/06/2025 a 30/07/2025
- Acuracia: 76,67%
- Matriz de confusao: `[[16, 2], [5, 7]]`

Esse resultado supera a meta minima de 75% definida no enunciado.

## Estrutura

```text
.
├── data/
│   ├── raw/
│   └── processed/
├── docs/
├── notebooks/
├── reports/
│   ├── apresentacao/
│   └── figures/
├── src/
├── README.md
└── requirements.txt
```

## Como executar

O notebook principal esta em:

`notebooks/tech_challenge_fase_2_ibovespa.ipynb`

Ele foi preparado para execucao no Google Colab e tambem em ambiente local.

As dependencias principais estao em `requirements.txt`.

Em ambiente local:

```bash
pip install -r requirements.txt
python -m src.data_prep
python -m src.modeling
```

No Google Colab, abra o notebook a partir do GitHub publico. A primeira celula clona o repositorio e ajusta o diretorio de trabalho automaticamente.
