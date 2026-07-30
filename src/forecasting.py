from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.stattools import adfuller

from src.data_prep import clean_ibovespa_data, load_raw_data


REGRESSION_TARGET = "fechamento_amanha"
TARGET_DATE_COLUMN = "data_alvo"


@dataclass
class ForecastResult:
    name: str
    wmape: float
    assertiveness: float
    mae: float
    hit_rate_0215: float


def wmape(y_true: pd.Series | np.ndarray, y_pred: pd.Series | np.ndarray) -> float:
    """Weighted MAPE, adequado para serie de fechamento sem valores proximos de zero."""
    actual = np.asarray(y_true, dtype=float)
    predicted = np.asarray(y_pred, dtype=float)
    return float(np.sum(np.abs(actual - predicted)) / np.sum(np.abs(actual)))


def mae(y_true: pd.Series | np.ndarray, y_pred: pd.Series | np.ndarray) -> float:
    actual = np.asarray(y_true, dtype=float)
    predicted = np.asarray(y_pred, dtype=float)
    return float(np.mean(np.abs(actual - predicted)))


def hit_rate(
    y_true: pd.Series | np.ndarray,
    y_pred: pd.Series | np.ndarray,
    tolerance: float = 0.0215,
) -> float:
    actual = np.asarray(y_true, dtype=float)
    predicted = np.asarray(y_pred, dtype=float)
    error_pct = np.abs(actual - predicted) / actual
    return float(np.mean(error_pct <= tolerance))


def build_regression_dataset() -> pd.DataFrame:
    cleaned = clean_ibovespa_data(load_raw_data())
    dataset = cleaned.copy()
    dataset[REGRESSION_TARGET] = dataset["fechamento"].shift(-1)
    dataset[TARGET_DATE_COLUMN] = dataset["data"].shift(-1)
    dataset["direcao_amanha"] = (
        dataset[REGRESSION_TARGET] > dataset["fechamento"]
    ).astype(int)
    return dataset.dropna(subset=[REGRESSION_TARGET]).reset_index(drop=True)


def temporal_split(
    dataset: pd.DataFrame,
    test_size: int = 125,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if len(dataset) <= test_size:
        raise ValueError("A base precisa ter mais registros que o tamanho do teste.")
    return dataset.iloc[:-test_size].copy(), dataset.iloc[-test_size:].copy()


def evaluate_predictions(
    name: str,
    y_true: pd.Series | np.ndarray,
    y_pred: pd.Series | np.ndarray,
    tolerance: float = 0.0215,
) -> ForecastResult:
    error = wmape(y_true, y_pred)
    return ForecastResult(
        name=name,
        wmape=error,
        assertiveness=1 - error,
        mae=mae(y_true, y_pred),
        hit_rate_0215=hit_rate(y_true, y_pred, tolerance=tolerance),
    )


def arima_static_forecast(train: pd.DataFrame, steps: int) -> np.ndarray:
    series = train["fechamento"].to_numpy()
    model = ARIMA(series, order=(5, 1, 0)).fit()
    return np.asarray(model.forecast(steps=steps), dtype=float)


def arima_walk_forward_forecast(train: pd.DataFrame, test: pd.DataFrame) -> np.ndarray:
    history = list(train["fechamento"].values)
    predictions = []
    for actual_close in test["fechamento"].values:
        # Ao final do pregao D, seu fechamento ja e conhecido e deve entrar
        # no historico antes de gerar a previsao para D+1.
        history.append(actual_close)
        model = ARIMA(history, order=(5, 1, 0)).fit()
        predictions.append(float(model.forecast(steps=1)[0]))
    return np.asarray(predictions, dtype=float)


def run_regression_experiments(test_size: int = 125) -> list[ForecastResult]:
    dataset = build_regression_dataset()
    train, test = temporal_split(dataset, test_size=test_size)
    y_true = test[REGRESSION_TARGET]

    naive_pred = test["fechamento"].values
    # O horizonte do teste comeca no alvo da primeira linha de teste. Logo,
    # o fechamento corrente dessa linha pertence ao historico disponivel.
    static_history = pd.concat([train, test.iloc[[0]]], ignore_index=True)
    static_pred = arima_static_forecast(static_history, steps=len(test))
    walk_forward_pred = arima_walk_forward_forecast(train, test)

    results = [
        evaluate_predictions("Naive - persistencia", y_true, naive_pred),
        evaluate_predictions("ARIMA(5,1,0) estatico", y_true, static_pred),
        evaluate_predictions("ARIMA(5,1,0) walk-forward", y_true, walk_forward_pred),
    ]
    return sorted(results, key=lambda result: result.hit_rate_0215, reverse=True)


def adf_summary() -> dict[str, float]:
    cleaned = clean_ibovespa_data(load_raw_data())
    original = adfuller(cleaned["fechamento"])
    differenced = adfuller(cleaned["fechamento"].diff().dropna())
    return {
        "p_valor_original": float(original[1]),
        "p_valor_primeira_diferenca": float(differenced[1]),
    }


def forecast_next_15_business_days() -> pd.DataFrame:
    cleaned = clean_ibovespa_data(load_raw_data())
    series = cleaned["fechamento"].values
    model = ARIMA(series, order=(5, 1, 0)).fit()
    forecast = model.get_forecast(steps=15)
    confidence = forecast.conf_int(alpha=0.05)
    dates = pd.bdate_range(cleaned["data"].max() + pd.Timedelta(days=1), periods=15)
    return pd.DataFrame(
        {
            "data": dates,
            "fechamento_previsto": forecast.predicted_mean,
            "ic_inferior_95": confidence[:, 0],
            "ic_superior_95": confidence[:, 1],
        }
    )


if __name__ == "__main__":
    dataset = build_regression_dataset()
    train_df, test_df = temporal_split(dataset)
    print(f"Base: {dataset['data'].min().date()} a {dataset['data'].max().date()}")
    print(f"Treino: {train_df['data'].min().date()} a {train_df['data'].max().date()}")
    print(f"Teste: {test_df['data'].min().date()} a {test_df['data'].max().date()}")
    print()
    for result in run_regression_experiments():
        print("=" * 80)
        print(result.name)
        print(f"WMAPE: {result.wmape:.2%}")
        print(f"Assertividade (1 - WMAPE): {result.assertiveness:.2%}")
        print(f"MAE: {result.mae:,.0f} pontos")
        print(f"Acerto dentro de 2,15%: {result.hit_rate_0215:.2%}")
