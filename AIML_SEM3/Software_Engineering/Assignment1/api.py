"""FastAPI client/server interface for the shared credit-risk service."""

from functools import lru_cache

import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from pydantic import BaseModel, Field

from credit_risk.data import load_data
from credit_risk.service import evaluate_model, fit_model
from credit_risk.strategies import STRATEGIES, get_strategy


app = FastAPI(
    title="Group 54 Credit Default Risk API",
    description="Educational prototype only; not for real credit decisions.",
)


class PredictionRequest(BaseModel):
    strategy: str = "Logistic regression"
    features: dict[str, float]
    threshold: float = Field(default=0.5, ge=0.0, le=1.0)


@lru_cache(maxsize=None)
def get_fitted_model(strategy_name):
    return fit_model(get_strategy(strategy_name))


def checked_model(strategy_name):
    try:
        return get_fitted_model(strategy_name)
    except ValueError as error:
        raise HTTPException(status_code=422, detail=str(error)) from error


@app.post("/predict")
def predict_endpoint(request: PredictionRequest):
    model = checked_model(request.strategy)
    columns = load_data()[0].columns
    missing = set(columns) - set(request.features)
    extra = set(request.features) - set(columns)
    if missing or extra:
        raise HTTPException(
            status_code=422,
            detail={"missing_features": sorted(missing), "unknown_features": sorted(extra)},
        )
    values = pd.DataFrame([[request.features[name] for name in columns]], columns=columns)
    probability = float(model.predict_proba(values)[0][list(model.classes_).index(1)])
    return {
        "default_probability": probability,
        "flag_for_review": probability >= request.threshold,
    }


@app.get("/metrics")
def metrics_endpoint(
    strategy: str = Query(default="Logistic regression"),
    threshold: float = Query(default=0.5, ge=0.0, le=1.0),
):
    model = checked_model(strategy)
    return {"strategy": strategy, "threshold": threshold, **evaluate_model(model, threshold)}


@app.get("/health")
def health_endpoint():
    return {"status": "ok", "strategies": list(STRATEGIES)}