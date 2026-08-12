-- Churn rate by NPS score band
-- Source: UK Telecom Churn Dataset, customers table    
SELECT
    CASE 
        WHEN "NPS Score (0-10)" BETWEEN 0 AND 3 THEN '0-3 Detractors'
        WHEN "NPS Score (0-10)" BETWEEN 4 AND 6 THEN '4-6 Passives'
        WHEN "NPS Score (0-10)" BETWEEN 7 AND 10 THEN '7-10 Promoters'
    END AS nps_band,
    COUNT(*) AS total_customers,
    SUM("Churned (1=Yes)") AS churned,
    ROUND(AVG("Churned (1=Yes)") * 100, 1) AS churn_rate_pct
FROM customers
GROUP BY nps_band
ORDER BY churn_rate_pct DESC
