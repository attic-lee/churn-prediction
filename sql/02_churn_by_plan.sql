-- Churn rate and average monthly fee by subscription plan
-- Source: UK Telecom Churn Dataset, customers table
-- Author: Attic Lee

SELECT 
    plan,
    COUNT(*) AS total_customers,
    SUM("Churned (1=Yes)") AS churned,
    ROUND(AVG("Churned (1=Yes)") * 100, 1) AS churn_rate_pct,
    ROUND(AVG("Monthly Fee (£)"), 2) AS avg_monthly_fee
FROM customers
GROUP BY plan
ORDER BY churn_rate_pct DESC