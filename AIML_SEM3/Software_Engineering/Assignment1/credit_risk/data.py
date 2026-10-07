"""Credit-default dataset access and deterministic evaluation split."""

from pathlib import Path

import pandas as pd
from sklearn.model_selection import train_test_split


RANDOM_STATE = 42
DATA_PATH = Path(__file__).resolve().parents[1] / "submission_assets" / "credit_default.csv"
EXCLUDED_FEATURES = ["SEX", "EDUCATION", "MARRIAGE", "AGE"]


def load_data():
    """Return credit-account features and a target where 1 means default next month."""
    dataset = pd.read_csv(DATA_PATH)
    target = dataset.pop("default").astype(int).rename("default")
    features = dataset.drop(columns=EXCLUDED_FEATURES)
    return features, target


def split_data(test_size=0.2):
    features, target = load_data()
    return train_test_split(
        features,
        target,
        test_size=test_size,
        random_state=RANDOM_STATE,
        stratify=target,
    )