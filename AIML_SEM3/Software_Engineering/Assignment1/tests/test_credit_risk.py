import numpy as np
import pandas as pd
import pytest
from fastapi.testclient import TestClient

from api import app
from credit_risk.data import DATA_PATH, EXCLUDED_FEATURES, load_data, split_data
from credit_risk.fairness import audit_recall_by_group
from credit_risk.service import fit_model, predict
from credit_risk.strategies import STRATEGIES, get_strategy


client = TestClient(app)


def test_dataset_schema_target_and_demographic_exclusions():
    features, target = load_data()
    raw = pd.read_csv(DATA_PATH)
    assert len(features) == 30_000
    assert features.shape[1] == 19
    assert not set(EXCLUDED_FEATURES).intersection(features.columns)
    assert not {"SEX", "EDUCATION", "MARRIAGE", "AGE"}.intersection(features.columns)
    assert "default" not in features.columns
    assert set(target.unique()) == {0, 1}
    assert int(target.sum()) == 6_636
    assert len(raw) == 30_000


def test_split_sizes_are_stratified_and_reproducible():
    first = split_data()
    second = split_data()
    x_train, x_test, y_train, y_test = first
    assert (len(x_train), len(x_test)) == (24_000, 6_000)
    for actual, repeated in zip(first, second):
        pd.testing.assert_frame_equal(actual, repeated) if isinstance(actual, pd.DataFrame) else pd.testing.assert_series_equal(actual, repeated)
    full_default_rate = (int(y_train.sum()) + int(y_test.sum())) / 30_000
    assert abs(y_train.mean() - full_default_rate) < 0.001
    assert abs(y_test.mean() - full_default_rate) < 0.001
    assert set(x_train.index).isdisjoint(x_test.index)


@pytest.mark.parametrize("strategy_name", list(STRATEGIES))
def test_strategy_probability_and_threshold_boundary(strategy_name):
    features, _ = load_data()
    model = fit_model(get_strategy(strategy_name))
    example = features.iloc[[0]]
    result = predict(model, example, threshold=0.5)
    assert 0.0 <= result["default_probability"] <= 1.0
    boundary = predict(model, example, threshold=result["default_probability"])
    assert boundary["flag_for_review"] is True
    below = predict(model, example, threshold=min(1.0, result["default_probability"] + 0.001))
    assert below["flag_for_review"] is (result["default_probability"] >= below["threshold"])


def test_unknown_strategy_raises_value_error():
    with pytest.raises(ValueError, match="Unknown model strategy"):
        get_strategy("not-a-model")


def test_fairness_audit_uses_demographics_only_for_audit():
    features, _ = load_data()
    model = fit_model(get_strategy("Logistic regression"))
    audit = audit_recall_by_group(model)
    assert not set(audit["recall_by_group"]).intersection(features.columns)
    assert set(audit["recall_by_group"]) == {"SEX", "AGE band"}
    assert 0.0 <= audit["max_gap_pp"] <= 100.0


def test_api_health_metrics_and_prediction():
    assert client.get("/health").status_code == 200
    metrics = client.get("/metrics", params={"strategy": "Logistic regression", "threshold": 0.5})
    assert metrics.status_code == 200
    assert metrics.json()["test_size"] == 6_000
    features, _ = load_data()
    request = {
        "strategy": "Logistic regression",
        "features": features.median().to_dict(),
        "threshold": 0.5,
    }
    prediction = client.post("/predict", json=request)
    assert prediction.status_code == 200
    body = prediction.json()
    assert 0.0 <= body["default_probability"] <= 1.0
    assert body["flag_for_review"] is (body["default_probability"] >= 0.5)
    bad_strategy = client.get("/metrics", params={"strategy": "unknown"})
    assert bad_strategy.status_code == 422