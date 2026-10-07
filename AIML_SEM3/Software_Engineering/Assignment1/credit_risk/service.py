"""Application-level training, evaluation, and inference operations."""

import logging

from sklearn.metrics import accuracy_score, confusion_matrix, precision_score, recall_score

from .data import split_data
from .strategies import LogisticRegressionStrategy


DEFAULT_THRESHOLD = 0.5
LOGGER = logging.getLogger(__name__)


def default_probability(model, features):
    default_index = list(model.classes_).index(1)
    return model.predict_proba(features)[:, default_index]


def fit_model(strategy=None):
    strategy = strategy or LogisticRegressionStrategy()
    x_train, _, y_train, _ = split_data()
    model = strategy.build()
    model.fit(x_train, y_train)
    return model


def evaluate_model(model, threshold=DEFAULT_THRESHOLD, split=None):
    _, x_test, _, y_test = split if split is not None else split_data()
    probabilities = default_probability(model, x_test)
    predictions = (probabilities >= threshold).astype(int)
    true_negative, false_positive, false_negative, true_positive = confusion_matrix(
        y_test, predictions, labels=[0, 1]
    ).ravel()
    metrics = {
        "accuracy": accuracy_score(y_test, predictions),
        "recall": recall_score(y_test, predictions, zero_division=0),
        "specificity": true_negative / (true_negative + false_positive),
        "precision": precision_score(y_test, predictions, zero_division=0),
        "true_negative": int(true_negative),
        "false_positive": int(false_positive),
        "false_negative": int(false_negative),
        "true_positive": int(true_positive),
        "test_size": len(y_test),
    }
    LOGGER.info(
        "Model evaluation: estimator=%s threshold=%.2f metrics=%s",
        type(model).__name__, threshold, metrics,
    )
    return metrics


def evaluate(strategy=None, threshold=DEFAULT_THRESHOLD):
    model = fit_model(strategy)
    metrics = evaluate_model(model, threshold)
    return model, metrics


def predict(model, features, threshold=DEFAULT_THRESHOLD):
    probability = float(default_probability(model, features)[0])
    return {
        "default_probability": probability,
        "flag_for_review": probability >= threshold,
        "threshold": threshold,
    }