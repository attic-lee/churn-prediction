-- Churn rate by contract type
-- Source: UK Telecom Churn Dataset, customers table
-- Author: Attic Lee


SELECT
    "Contract Type" AS contract_type,
    COUNT(*) AS total_customers,
    SUM("Churned (1=Yes)") AS churned,
    ROUND(AVG("Churned (1=Yes)") * 100, 1) AS churn_rate_pct,
    ROUND(AVG("Tenure (Months)"), 1) AS avg_tenure_months
FROM customers
GROUP BY "Contract Type"
ORDER BY churn_rate_pct DESC