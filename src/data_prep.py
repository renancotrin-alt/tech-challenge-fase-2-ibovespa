from __future__ import annotations

from math import log
from pathlib import Path

import pandas as pd


RAW_DATA_PATH = Path("data/raw/dados_historicos_ibovespa.csv")
PROCESSED_DATA_PATH = Path("data/processed/ibovespa_modelagem.csv")


def parse_brazilian_number(value: object) -> float:
    """Converte valores no formato brasileiro usado pelo Investing.com."""
    if pd.isna(value):
        return float("nan")

    text = str(value).strip()
    if text == "":
        return float("nan")

    return float(text.replace(".", "").replace(",", "."))


def parse_percentage(value: object) -> float:
    """Converte percentuais como '-0,69%' para decimal, exemplo: -0.0069."""
    if pd.isna(value):
        return float("nan")

    text = str(value).strip().replace("%", "")
    return parse_brazilian_number(text) / 100


def parse_volume(value: object) -> float:
    """Converte volumes como '9,20B' e '14,45M' para numero absoluto."""
    if pd.isna(value):
        return float("nan")

    text = str(value).strip()
    if text == "":
        return float("nan")

    suffix = text[-1].upper()
    multiplier = 1.0
    if suffix == "B":
        multiplier = 1_000_000_000
        text = text[:-1]
    elif suffix == "M":
        multiplier = 1_000_000
        text = text[:-1]
    elif suffix == "K":
        multiplier = 1_000
        text = text[:-1]

    return parse_brazilian_number(text) * multiplier


def calculate_rsi(series: pd.Series, window: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(window).mean()
    avg_loss = loss.rolling(window).mean()
    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))


def load_raw_data(path: Path = RAW_DATA_PATH) -> pd.DataFrame:
    # Lemos tudo como texto para preservar zeros finais em valores como
    # "133.990", que representam 133990 pontos no formato do Investing.com.
    return pd.read_csv(path, encoding="utf-8-sig", dtype=str)


def clean_ibovespa_data(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()

    cleaned = cleaned.rename(
        columns={
            "Data": "data",
            "Último": "fechamento",
            "Abertura": "abertura",
            "Máxima": "maxima",
            "Mínima": "minima",
            "Vol.": "volume",
            "Var%": "variacao_pct",
        }
    )

    cleaned["data"] = pd.to_datetime(cleaned["data"], format="%d.%m.%Y")

    for column in ["fechamento", "abertura", "maxima", "minima"]:
        cleaned[column] = cleaned[column].apply(parse_brazilian_number)

    cleaned["variacao_pct"] = cleaned["variacao_pct"].apply(parse_percentage)
    cleaned["volume_num"] = cleaned["volume"].apply(parse_volume)
    cleaned = cleaned.sort_values("data").reset_index(drop=True)

    return cleaned


def add_target(df: pd.DataFrame) -> pd.DataFrame:
    featured = df.copy()
    featured["fechamento_amanha"] = featured["fechamento"].shift(-1)
    featured["alta_amanha"] = (
        featured["fechamento_amanha"] > featured["fechamento"]
    ).astype(int)

    # A ultima linha nao tem o fechamento do proximo pregao, entao nao pode
    # entrar na modelagem supervisionada.
    featured = featured.dropna(subset=["fechamento_amanha"]).reset_index(drop=True)
    return featured


def add_basic_features(df: pd.DataFrame) -> pd.DataFrame:
    featured = df.copy()

    featured["retorno_1d"] = featured["fechamento"].pct_change()
    featured["retorno_2d"] = featured["fechamento"].pct_change(2)
    featured["retorno_5d"] = featured["fechamento"].pct_change(5)
    featured["retorno_10d"] = featured["fechamento"].pct_change(10)
    featured["retorno_21d"] = featured["fechamento"].pct_change(21)

    featured["retorno_lag_1"] = featured["retorno_1d"].shift(1)
    featured["retorno_lag_2"] = featured["retorno_1d"].shift(2)
    featured["retorno_lag_3"] = featured["retorno_1d"].shift(3)
    featured["retorno_lag_5"] = featured["retorno_1d"].shift(5)

    featured["media_movel_5"] = featured["fechamento"].rolling(5).mean()
    featured["media_movel_10"] = featured["fechamento"].rolling(10).mean()
    featured["media_movel_21"] = featured["fechamento"].rolling(21).mean()
    featured["media_movel_50"] = featured["fechamento"].rolling(50).mean()

    featured["distancia_mm5"] = featured["fechamento"] / featured["media_movel_5"] - 1
    featured["distancia_mm21"] = featured["fechamento"] / featured["media_movel_21"] - 1
    featured["distancia_mm50"] = featured["fechamento"] / featured["media_movel_50"] - 1

    featured["amplitude_intradia"] = featured["maxima"] / featured["minima"] - 1
    featured["corpo_candle"] = featured["fechamento"] / featured["abertura"] - 1
    featured["volatilidade_5d"] = featured["retorno_1d"].rolling(5).std()
    featured["volatilidade_21d"] = featured["retorno_1d"].rolling(21).std()
    featured["volume_log"] = (featured["volume_num"] + 1).apply(log)
    featured["volume_media_5"] = featured["volume_log"].rolling(5).mean()
    featured["volume_choque_5"] = featured["volume_log"] / featured["volume_media_5"] - 1

    featured["maxima_5d"] = featured["fechamento"].rolling(5).max()
    featured["minima_5d"] = featured["fechamento"].rolling(5).min()
    featured["posicao_5d"] = (
        (featured["fechamento"] - featured["minima_5d"])
        / (featured["maxima_5d"] - featured["minima_5d"])
    )
    featured["maxima_21d"] = featured["fechamento"].rolling(21).max()
    featured["minima_21d"] = featured["fechamento"].rolling(21).min()
    featured["posicao_21d"] = (
        (featured["fechamento"] - featured["minima_21d"])
        / (featured["maxima_21d"] - featured["minima_21d"])
    )

    featured["alta_hoje"] = (featured["retorno_1d"] > 0).astype(int)
    featured["alta_lag_1"] = featured["alta_hoje"].shift(1)
    featured["alta_lag_2"] = featured["alta_hoje"].shift(2)
    featured["alta_lag_3"] = featured["alta_hoje"].shift(3)
    featured["altas_5d"] = featured["alta_hoje"].rolling(5).sum()

    featured["rsi_14"] = calculate_rsi(featured["fechamento"], 14)
    ema_12 = featured["fechamento"].ewm(span=12, adjust=False).mean()
    ema_26 = featured["fechamento"].ewm(span=26, adjust=False).mean()
    featured["macd"] = ema_12 - ema_26
    featured["macd_sinal"] = featured["macd"].ewm(span=9, adjust=False).mean()
    featured["macd_hist"] = featured["macd"] - featured["macd_sinal"]

    bollinger_mean = featured["fechamento"].rolling(20).mean()
    bollinger_std = featured["fechamento"].rolling(20).std()
    featured["bollinger_pos"] = (
        (featured["fechamento"] - bollinger_mean) / (2 * bollinger_std)
    )

    featured["dia_semana"] = featured["data"].dt.dayofweek
    featured["mes"] = featured["data"].dt.month

    featured = featured.replace([float("inf"), float("-inf")], pd.NA)
    featured = featured.dropna().reset_index(drop=True)
    return featured


def build_modeling_dataset(raw_path: Path = RAW_DATA_PATH) -> pd.DataFrame:
    raw = load_raw_data(raw_path)
    cleaned = clean_ibovespa_data(raw)
    with_target = add_target(cleaned)
    return add_basic_features(with_target)


def temporal_train_test_split(
    df: pd.DataFrame, test_size: int = 30
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if len(df) <= test_size:
        raise ValueError("A base precisa ter mais registros que o tamanho do teste.")

    train = df.iloc[:-test_size].copy()
    test = df.iloc[-test_size:].copy()
    return train, test


def save_processed_dataset(
    df: pd.DataFrame, path: Path = PROCESSED_DATA_PATH
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def summarize_dataset(df: pd.DataFrame, train: pd.DataFrame, test: pd.DataFrame) -> dict:
    return {
        "linhas_modelagem": len(df),
        "data_inicial": df["data"].min().date().isoformat(),
        "data_final": df["data"].max().date().isoformat(),
        "treino_linhas": len(train),
        "treino_inicio": train["data"].min().date().isoformat(),
        "treino_fim": train["data"].max().date().isoformat(),
        "teste_linhas": len(test),
        "teste_inicio": test["data"].min().date().isoformat(),
        "teste_fim": test["data"].max().date().isoformat(),
        "classe_alta_pct_total": round(float(df["alta_amanha"].mean()), 4),
        "classe_alta_pct_teste": round(float(test["alta_amanha"].mean()), 4),
    }


if __name__ == "__main__":
    dataset = build_modeling_dataset()
    train_df, test_df = temporal_train_test_split(dataset, test_size=30)
    save_processed_dataset(dataset)

    summary = summarize_dataset(dataset, train_df, test_df)
    for key, value in summary.items():
        print(f"{key}: {value}")
