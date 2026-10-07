# Group 54: Credit Default Risk Review Prototype

Course: AIMLZG546 - Software Engineering for Machine Learning, Assignment I

This educational prototype uses the UCI **Default of Credit Card Clients** dataset to estimate whether an existing credit-card account will default on payment next month. It is intended for analyst review and software-engineering demonstration only. It must not be used for credit approval, pricing, collections, or other customer-level decisions.

## Run

```powershell
python -m pip install -r requirements-dev.txt
streamlit run app.py
```

The optional API can be started separately with `uvicorn api:app --reload`; its interactive schema is available at `/docs`. Run the checks with `pytest`.

The bundled dataset makes application runs independent of network access. The app evaluates a stratified 80/20 split, allows logistic-regression or random-forest strategies, adjusts a review threshold, and exposes six exploratory account indicators. Other model inputs are held at their dataset medians. Streamlit caches each fitted strategy in memory, so threshold and account-control changes reuse the fitted model while metrics are recomputed without retraining.

## Group Details

| Sl. No | BITS ID | Name | Contribution (Qualitative) | Percentage (out of 100) |
|---|---|---|---|---|
| 1 | 2025AE05017 | POONAM TANNA | Contributed across most project areas, collaborating on problem formulation, GR4ML views, system architecture, pattern implementation, Streamlit UI, and report generation. | 40 |
| 2 | 2025AE05791 | Mohammed Asif S | Requirements formulation: domain and problem statement, users and stakeholders, GR4ML measurable goals and indicators table, functional and non-functional requirements with acceptance criteria, and justification of the top three quality requirements. | 20 |
| 3 | 2025AF05145 | NAVEEN.K.B | GR4ML modelling: Business View, Analytics Design View and Data Preparation View (tables and diagrams), cross-view traceability, softgoal contribution links and decomposition correctness. | 20 |
| 4 | 2025AF05136 | SIMRAN SURI | Evaluation and reporting: model evaluation experiments (threshold sweep, fairness audit by subgroup), application testing and screenshots with captions, notebook execution and saved outputs, report compilation, formatting and consistency review between docx and notebook. | 20 |

## Files

- `54.ipynb`: assignment notebook covering requirements, GR4ML views, architecture, and implementation.
- `54.docx`: submission report, including the architecture diagram and application screenshot.
- `app.py`: Streamlit presentation layer.
- `api.py`: optional FastAPI client/server interface for prediction and holdout metrics.
- `credit_risk/`: dataset adapter, model strategies, application service, and fairness audit.
- `submission_assets/credit_default.csv`: bundled UCI source data with the row identifier removed.
- `tests/`: pytest checks for data contracts, model strategies, thresholds, fairness inputs, and API endpoints.
- `requirements-dev.txt`: pytest, Playwright, and notebook execution tools.
- `build_submission.py`: regenerates the report, notebook, and diagrams.
- `54_submission.zip`: submission bundle with report, notebook, source, assets, requirements, and tests.

Run verification with `python -m pip install -r requirements-dev.txt` followed by `pytest`.

## Dataset and safeguards

The source dataset contains 30,000 Taiwan credit-card accounts, 23 predictors, and a binary next-month default outcome. The local adapter removes the record identifier and excludes sex, education, marital status, and age from model inputs; 19 predictors remain. The dataset is imbalanced (6,636 defaults). It is a historical benchmark, not a representative current portfolio. The fixed holdout demonstrates an evaluation workflow only and does not establish calibration, fairness, generalization, or suitability for financial decisions. A real deployment would require independent temporal and population validation, fairness review, governance, security controls, and human oversight.

Source: Yeh, I.-C. and Lien, C.-H. (2009), “The comparisons of data mining techniques for the predictive accuracy of probability of default of credit card clients,” *Expert Systems with Applications*, 36(2), 2473-2480. UCI Machine Learning Repository, [Default of Credit Card Clients](https://archive.ics.uci.edu/dataset/350/default+of+credit+card+clients).