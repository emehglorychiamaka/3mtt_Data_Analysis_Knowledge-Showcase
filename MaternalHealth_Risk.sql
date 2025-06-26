-- View All Records
SELECT * 
FROM MaternalHealth;


-- Get Basic Summary Count
SELECT COUNT(*) AS TotalRecords 
FROM MaternalHealth;


-- Count by Risk Level
SELECT RiskLevel, COUNT(*) AS Count
FROM MaternalHealth
GROUP BY RiskLevel
ORDER BY Count DESC;

-- Average Vitals by Risk Level
SELECT RiskLevel,
       AVG(SystolicBP) AS Avg_SystolicBP,
       AVG(DiastolicBP) AS Avg_DiastolicBP,
       AVG(BS) AS Avg_BloodSugar,
       AVG(HeartRate) AS Avg_HeartRate
FROM MaternalHealth
GROUP BY RiskLevel;


-- Get High-Risk Patients Only
SELECT *
FROM MaternalHealth
WHERE RiskLevel = 'high risk';


-- Group by Age Range (Bucketed)
SELECT
  CASE 
    WHEN Age < 20 THEN 'Below 20'
    WHEN Age BETWEEN 20 AND 29 THEN '20-29'
    WHEN Age BETWEEN 30 AND 39 THEN '30-39'
    WHEN Age BETWEEN 40 AND 49 THEN '40-49'
    ELSE '50+'
  END AS AgeGroup,
  COUNT(*) AS Count
FROM MaternalHealth
GROUP BY 
  CASE 
    WHEN Age < 20 THEN 'Below 20'
    WHEN Age BETWEEN 20 AND 29 THEN '20-29'
    WHEN Age BETWEEN 30 AND 39 THEN '30-39'
    WHEN Age BETWEEN 40 AND 49 THEN '40-49'
    ELSE '50+'
  END;


-- Top 10 Highest Risk Patients by BP
SELECT TOP 10 *
FROM MaternalHealth
WHERE RiskLevel = 'high risk'
ORDER BY SystolicBP DESC, DiastolicBP DESC;


-- Add a RiskScore
SELECT *,
       (SystolicBP * 0.3 + DiastolicBP * 0.3 + BS * 0.4) AS RiskScore
FROM MaternalHealth
ORDER BY RiskScore DESC;

