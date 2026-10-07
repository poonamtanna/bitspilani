"""Interchangeable model strategies used by the application service."""

from typing import Protocol

from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler


class ModelStrategy(Protocol):
    name: str

    def build(self):
        """Create an unfitted estimator."""


class LogisticRegressionStrategy:
    name = "Logistic regression"

    def build(self):
        return make_pipeline(
            StandardScaler(),
            LogisticRegression(class_weight="balanced", max_iter=1500, random_state=42),
        )


class RandomForestStrategy:
    name = "Random forest"

    def build(self):
        return RandomForestClassifier(
            n_estimators=160,
            class_weight="balanced",
            min_samples_leaf=10,
            n_jobs=1,
            random_state=42,
        )


STRATEGIES = {
    LogisticRegressionStrategy.name: LogisticRegressionStrategy,
    RandomForestStrategy.name: RandomForestStrategy,
}


def get_strategy(name):
    try:
        return STRATEGIES[name]()
    except KeyError as error:
        raise ValueError(f"Unknown model strategy: {name}") from error