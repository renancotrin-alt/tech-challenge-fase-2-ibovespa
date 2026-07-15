from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
from sklearn.ensemble import GradientBoostingClassifier, RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import accuracy_score, classification_report, confusion_matrix
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC

from src.data_prep import build_modeling_dataset, temporal_train_test_split


TARGET_COLUMN = "alta_amanha"
FEATURE_COLUMNS = [
    "fechamento",
    "abertura",
    "maxima",
    "minima",
    "variacao_pct",
    "retorno_1d",
    "retorno_2d",
    "retorno_5d",
    "retorno_10d",
    "retorno_21d",
    "retorno_lag_1",
    "retorno_lag_2",
    "retorno_lag_3",
    "retorno_lag_5",
    "distancia_mm5",
    "distancia_mm21",
    "distancia_mm50",
    "amplitude_intradia",
    "corpo_candle",
    "volatilidade_5d",
    "volatilidade_21d",
    "volume_log",
    "volume_choque_5",
    "posicao_5d",
    "posicao_21d",
    "alta_hoje",
    "alta_lag_1",
    "alta_lag_2",
    "alta_lag_3",
    "altas_5d",
    "rsi_14",
    "macd",
    "macd_sinal",
    "macd_hist",
    "bollinger_pos",
    "dia_semana",
    "mes",
]


@dataclass
class ModelResult:
    name: str
    train_window: int
    accuracy: float
    confusion: list[list[int]]
    report: str


class MajorityClassBaseline:
    def fit(self, _x: pd.DataFrame, y: pd.Series) -> "MajorityClassBaseline":
        self.majority_class_ = int(y.mode().iloc[0])
        return self

    def predict(self, x: pd.DataFrame) -> list[int]:
        return [self.majority_class_] * len(x)


class LastTrendBaseline:
    def fit(self, _x: pd.DataFrame, y: pd.Series) -> "LastTrendBaseline":
        self.last_class_ = int(y.iloc[-1])
        return self

    def predict(self, x: pd.DataFrame) -> list[int]:
        return [self.last_class_] * len(x)


def get_model_specs() -> list[tuple[str, object, int | None]]:
    return [
        ("Baseline - classe majoritaria", MajorityClassBaseline(), None),
        ("Baseline - ultima tendencia", LastTrendBaseline(), None),
        (
            "Regressao Logistica",
            Pipeline(
                steps=[
                    ("scaler", StandardScaler()),
                    (
                        "model",
                        LogisticRegression(
                            max_iter=1000,
                            class_weight="balanced",
                            C=0.1,
                            random_state=42,
                        ),
                    ),
                ]
            ),
            None,
        ),
        (
            "Random Forest",
            RandomForestClassifier(
                n_estimators=300,
                max_depth=4,
                min_samples_leaf=8,
                class_weight="balanced",
                random_state=42,
            ),
            None,
        ),
        (
            "Gradient Boosting",
            GradientBoostingClassifier(
                n_estimators=120,
                learning_rate=0.04,
                max_depth=2,
                min_samples_leaf=8,
                random_state=42,
            ),
            None,
        ),
        (
            "SVC - janela recente 150 pregoes",
            Pipeline(
                steps=[
                    ("scaler", StandardScaler()),
                    (
                        "model",
                        SVC(
                            C=1.0,
                            kernel="rbf",
                            class_weight="balanced",
                            random_state=42,
                        ),
                    ),
                ]
            ),
            150,
        ),
    ]


def evaluate_model(
    name: str,
    model: object,
    x_train: pd.DataFrame,
    y_train: pd.Series,
    x_test: pd.DataFrame,
    y_test: pd.Series,
    train_window: int,
) -> ModelResult:
    model.fit(x_train, y_train)
    y_pred = model.predict(x_test)

    return ModelResult(
        name=name,
        train_window=train_window,
        accuracy=float(accuracy_score(y_test, y_pred)),
        confusion=confusion_matrix(y_test, y_pred, labels=[0, 1]).tolist(),
        report=classification_report(
            y_test,
            y_pred,
            labels=[0, 1],
            target_names=["Baixa/igual", "Alta"],
            zero_division=0,
        ),
    )


def run_experiments() -> list[ModelResult]:
    dataset = build_modeling_dataset()
    train_df, test_df = temporal_train_test_split(dataset, test_size=30)

    x_test = test_df[FEATURE_COLUMNS]
    y_test = test_df[TARGET_COLUMN]

    results = []
    for name, model, window_size in get_model_specs():
        model_train_df = train_df if window_size is None else train_df.tail(window_size)
        x_train = model_train_df[FEATURE_COLUMNS]
        y_train = model_train_df[TARGET_COLUMN]
        results.append(
            evaluate_model(
                name,
                model,
                x_train,
                y_train,
                x_test,
                y_test,
                train_window=len(model_train_df),
            )
        )

    return sorted(results, key=lambda item: item.accuracy, reverse=True)


if __name__ == "__main__":
    for result in run_experiments():
        print("=" * 80)
        print(result.name)
        print(f"Janela de treino: {result.train_window} registros")
        print(f"Acuracia: {result.accuracy:.2%}")
        print(f"Matriz de confusao [[baixa_real, alta_prevista?], ...]: {result.confusion}")
        print(result.report)
