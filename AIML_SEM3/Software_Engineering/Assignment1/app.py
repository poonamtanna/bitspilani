"""Streamlit user interface for credit-default risk review."""

import streamlit as st

from credit_risk.data import load_data, split_data
from credit_risk.service import evaluate_model, fit_model as train_model, predict
from credit_risk.strategies import STRATEGIES, get_strategy


st.set_page_config(page_title="Credit Default Risk Review | Group 54", page_icon="$", layout="wide")
st.title("Credit Default Risk Review")
st.caption("A financial-risk coursework prototype for analyst review, not an automated credit decision.")


@st.cache_resource
def fit_model(model_name):
    return train_model(get_strategy(model_name))


@st.cache_data
def load_app_data():
    return load_data()


@st.cache_data
def load_evaluation_split():
    return split_data()


features, _ = load_app_data()
evaluation_split = load_evaluation_split()
model_name = st.sidebar.selectbox("Model strategy", list(STRATEGIES))
threshold = st.sidebar.slider("Review threshold", 0.10, 0.90, 0.50, 0.01)
model = fit_model(model_name)
metrics = evaluate_model(model, threshold, evaluation_split)

st.subheader("Held-out evaluation")
metric_columns = st.columns(4)
for column, label, key in zip(
    metric_columns,
    ["Default recall", "Specificity", "Precision", "Accuracy"],
    ["recall", "specificity", "precision", "accuracy"],
):
    column.metric(label, f"{metrics[key]:.1%}")
st.caption(
    f"Stratified 80/20 split; {metrics['test_size']} held-out records. "
    "Historical retrospective results; not evidence of future portfolio or customer-level performance."
)

st.subheader("Account review")
st.write("Adjust six account indicators. The other 13 model inputs are held at dataset medians.")
selected_features = [
    "LIMIT_BAL",
    "PAY_0",
    "PAY_2",
    "BILL_AMT1",
    "PAY_AMT1",
    "BILL_AMT2",
]
case = features.median().to_frame().T.copy()
input_columns = st.columns(3)
for index, feature_name in enumerate(selected_features):
    with input_columns[index % 3]:
        low = float(features[feature_name].quantile(0.05))
        high = float(features[feature_name].quantile(0.95))
        default = float(features[feature_name].median())
        if feature_name in {"PAY_0", "PAY_2"}:
            low, high, default = int(low), int(high), int(round(default))
            case.loc[0, feature_name] = st.slider(
                feature_name, low, high, default, 1
            )
        else:
            step = float(max(100, round((high - low) / 1000 / 100) * 100))
            low = float(round(low / step) * step)
            high = float(round(high / step) * step)
            default = float(min(max(round(default / step) * step, low), high))
            case.loc[0, feature_name] = st.slider(
                feature_name, low, high, default, step, format="%.0f"
            )

result = predict(model, case, threshold)
left, right = st.columns([2, 1])
with left:
    st.metric("Model-estimated next-month default risk", f"{result['default_probability']:.1%}")
    st.progress(result["default_probability"])
with right:
    if result["flag_for_review"]:
        st.error("Flagged for credit-risk analyst review")
    else:
        st.success("Below the selected review threshold")

st.warning(
    "Educational demonstration only. Do not use this score for credit approval, pricing, "
    "collections, or other customer-level decisions. Inputs are illustrative controls; "
    "independent validation, fairness review, and qualified oversight are required."
)