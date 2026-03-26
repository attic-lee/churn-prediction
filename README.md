# Customer Churn Prediction — UK Telecoms

End-to-end churn prediction pipeline built as a consulting analytics portfolio project.

🔗 **[Live Dashboard](https://churn-prediction-qvtngptey9n7gz2wfkz6yf.streamlit.app/)**

## Overview
Predicts customer churn for a UK telecoms dataset (1,000 customers, 23 features) 
using XGBoost, achieving AUC 0.784. Results delivered via an interactive Streamlit 
dashboard and structured consulting recommendations.

## Key Findings
- Customers in their first 6 months churn at **60.3%** — nearly 3× the rate of established customers
- Detractors (NPS 0–3) churn at **54.8%** — 2.4× the rate of Promoters
- Month-to-Month contracts churn at **35.6%** vs 15.5% for 2-Year contracts
- High-risk segment (low NPS + complaint + Month-to-Month) churns at **80%**, 
  putting £690/mo revenue at immediate risk

## Tech Stack
Python · XGBoost · SQL (SQLite) · Streamlit · Pandas · Scikit-learn · Plotly · Excel

## Project Structure
```
churn-prediction/
├── data/         # UK Telecom Churn Dataset
├── sql/          # SQL queries for segment analysis
├── notebooks/    # EDA and modelling notebooks
├── app/          # Streamlit dashboard
└── outputs/      # Charts, scored dataset, Excel report
```

## Results
| Model | AUC |
|-------|-----|
| XGBoost | 0.784 |
| Logistic Regression | 0.753 |
| Random Baseline | 0.500 |