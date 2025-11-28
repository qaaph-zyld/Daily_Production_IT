SET DATEFIRST 1; -- Set Monday as the first day of the week (ISO standard)

WITH BaseData AS (
    SELECT 
        sod.[sod_part] AS [Item Number],
        schd.[schd_date] AS [Date],
        CAST(schd.[schd_discr_qty] AS INT) AS [Discrete Qty],  -- Cast to integer
        CASE 
            WHEN schd.[schd_date] < CAST(GETDATE() AS DATE) THEN 'Past Due'
            WHEN DATEDIFF(WEEK, GETDATE(), schd.[schd_date]) BETWEEN 1 AND 8 THEN 
                'Week ' + CAST(DATEDIFF(WEEK, GETDATE(), schd.[schd_date]) AS VARCHAR(10))
            WHEN DATEDIFF(WEEK, GETDATE(), schd.[schd_date]) >= 9 THEN 'Long Term'
            ELSE 'Other'
        END AS [Time Period]
    FROM [QADEE2798].[dbo].[sod_det] sod
    LEFT JOIN [QADEE2798].[dbo].[active_schd_det] schd
        ON sod.[sod_nbr] = schd.[schd_nbr] 
        AND sod.[sod_line] = schd.[schd_line]
        AND sod.[sod_curr_rlse_id[3]]] = schd.[schd_rlse_id]
    WHERE sod.[sod_status] IS NULL
        AND (sod.[sod_end_eff[1]]] IS NULL OR sod.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE))
        AND schd.[schd_date] IS NOT NULL
)
SELECT 
    [Item Number],
    SUM(CAST(CASE WHEN [Time Period] = 'Past Due' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Total Past Due],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 1' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 1],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 2' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 2],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 3' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 3],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 4' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 4],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 5' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 5],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 6' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 6],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 7' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 7],
    SUM(CAST(CASE WHEN [Time Period] = 'Week 8' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 8],
    SUM(CAST(CASE WHEN [Time Period] = 'Long Term' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Total Long Term]
FROM BaseData
GROUP BY [Item Number]
ORDER BY [Item Number];