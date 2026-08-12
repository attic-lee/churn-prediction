-- High-risk customer segment analysis
-- Source: UK Telecom Churn Dataset, customers table

SELECT
    COUNT(*) AS high_risk_customers,
    ROUND(AVG("Churned (1=Yes)") * 100, 1) AS churn_rate_pct,
    ROUND(SUM("Monthly Fee (£)"), 2) AS monthly_revenue_at_risk
FROM customers
WHERE "NPS Score (0-10)" <= 3
    AND "Contract Type" = 'Month-to-Month'
    AND "Complaints Raised" >= 1
