"""Coursework subgroup recall audit using demographics excluded from modeling."""

import pandas as pd
from sklearn.metrics import recall_score

from .data import DATA_PATH, split_data
from .service import default_probability


def audit_recall_by_group(model, threshold=0.5):
    """Return holdout recall by sex and age band; audit fields never enter the model."""
    _, x_test, _, y_test = split_data()
    audit = pd.read_csv(DATA_PATH, usecols=["SEX", "AGE"]).loc[x_test.index]
    predictions = (default_probability(model, x_test) >= threshold).astype(int)
    rows = pd.DataFrame({"target": y_test, "prediction": predictions}, index=x_test.index)
    rows["SEX"] = audit["SEX"]
    rows["AGE band"] = pd.cut(
        audit["AGE"],
        bins=[0, 29, 39, 49, 59, float("inf")],
        labels=["<=29", "30-39", "40-49", "50-59", "60+"],
    )

    recall_by_group = {}
    for column in ("SEX", "AGE band"):
        group_scores = {}
        for group, values in rows.groupby(column, observed=True):
            group_scores[str(group)] = float(
                recall_score(values["target"], values["prediction"], zero_division=0)
            )
        recall_by_group[column] = group_scores

    gaps = [
        max(scores.values()) - min(scores.values())
        for scores in recall_by_group.values()
        if scores
    ]
    return {
        "recall_by_group": recall_by_group,
        "max_gap_pp": max(gaps, default=0.0) * 100,
        "target_pp": 5.0,
        "meets_target": max(gaps, default=0.0) * 100 <= 5.0,
    }