-- First identify all distinct part/site combinations from both databases
WITH DistinctPartSites AS (
    SELECT DISTINCT tr_part, tr_site
    FROM (
        SELECT tr_part, tr_site FROM [QADEE2798].[dbo].[tr_hist]
        WHERE tr_type = 'iss-wo'
          AND tr_effdate >= DATEADD(WEEK, -4, GETDATE())
        
        UNION
        
        SELECT tr_part, tr_site FROM [QADEE].[dbo].[tr_hist]
        WHERE tr_type = 'iss-wo'
          AND tr_effdate >= DATEADD(WEEK, -4, GETDATE())
    ) AS CombinedSources
),

-- Daily ISS-WO data (last 4 weeks only)
DailyISS_WO AS (
    SELECT
        tr_part,
        tr_site,
        tr_effdate,
        SUM(tr_qty_loc) AS Daily_Qty
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE tr_type = 'iss-wo'
      AND tr_effdate >= DATEADD(WEEK, -4, GETDATE())
    GROUP BY
        tr_part,
        tr_site,
        tr_effdate
    
    UNION ALL
    
    SELECT
        tr_part,
        tr_site,
        tr_effdate,
        SUM(tr_qty_loc) AS Daily_Qty
    FROM [QADEE].[dbo].[tr_hist]
    WHERE tr_type = 'iss-wo'
      AND tr_effdate >= DATEADD(WEEK, -4, GETDATE())
    GROUP BY
        tr_part,
        tr_site,
        tr_effdate
),

-- Get the current week number and year
CurrentWeekInfo AS (
    SELECT 
        DATEPART(YEAR, GETDATE()) AS CurrentYear,
        DATEPART(WEEK, GETDATE()) AS CurrentWeek
),

-- Weekly ISS-WO data with relative week calculations (last 4 weeks only)
WeeklyISS_WO AS (
    SELECT
        d.tr_part,
        d.tr_site,
        DATEPART(YEAR, d.tr_effdate) AS YearNumber,
        DATEPART(WEEK, d.tr_effdate) AS WeekNumber,
        cw.CurrentYear,
        cw.CurrentWeek,
        -- Calculate relative week position
        CASE
            WHEN DATEPART(YEAR, d.tr_effdate) = cw.CurrentYear THEN
                cw.CurrentWeek - DATEPART(WEEK, d.tr_effdate)
            ELSE
                cw.CurrentWeek + (52 - DATEPART(WEEK, d.tr_effdate))
        END AS WeeksAgo,
        d.Daily_Qty
    FROM DailyISS_WO d
    CROSS JOIN CurrentWeekInfo cw
    WHERE d.tr_effdate >= DATEADD(WEEK, -4, GETDATE())
),

-- Calculate weekly averages for each part/site
WeeklyAverages AS (
    SELECT
        dps.tr_part,
        dps.tr_site,
        -- Average daily iss-wo for previous week (Week -1)
        ISNULL((
            SELECT AVG(Daily_Qty)*(-1)
            FROM WeeklyISS_WO w
            WHERE w.tr_part = dps.tr_part
              AND w.tr_site = dps.tr_site
              AND w.WeeksAgo = 1
        ), 0) AS [avg_ISS-WO_CW_-1],
        
        -- Average daily iss-wo for 2 weeks ago (Week -2)
        ISNULL((
            SELECT AVG(Daily_Qty)*(-1)
            FROM WeeklyISS_WO w
            WHERE w.tr_part = dps.tr_part
              AND w.tr_site = dps.tr_site
              AND w.WeeksAgo = 2
        ), 0) AS [avg_ISS-WO_CW_-2],
        
        -- Average daily iss-wo for 3 weeks ago (Week -3)
        ISNULL((
            SELECT AVG(Daily_Qty)*(-1)
            FROM WeeklyISS_WO w
            WHERE w.tr_part = dps.tr_part
              AND w.tr_site = dps.tr_site
              AND w.WeeksAgo = 3
        ), 0) AS [avg_ISS-WO_CW_-3],
        
        -- Average daily iss-wo for 4 weeks ago (Week -4)
        ISNULL((
            SELECT AVG(Daily_Qty)*(-1)
            FROM WeeklyISS_WO w
            WHERE w.tr_part = dps.tr_part
              AND w.tr_site = dps.tr_site
              AND w.WeeksAgo = 4
        ), 0) AS [avg_ISS-WO_CW_-4]
    FROM DistinctPartSites dps
)

-- Final result set with WIP_minimum calculation
SELECT
    wa.tr_part,
    wa.tr_site,
    wa.[avg_ISS-WO_CW_-1],
    wa.[avg_ISS-WO_CW_-2],
    wa.[avg_ISS-WO_CW_-3],
    wa.[avg_ISS-WO_CW_-4],
    -- Calculate WIP_minimum as 3x average of last 4 weeks
    (wa.[avg_ISS-WO_CW_-1] + wa.[avg_ISS-WO_CW_-2] + 
     wa.[avg_ISS-WO_CW_-3] + wa.[avg_ISS-WO_CW_-4]) / 4 * 3 AS [WIP_minimum]
FROM WeeklyAverages wa
ORDER BY wa.tr_site, wa.tr_part;