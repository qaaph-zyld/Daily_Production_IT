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
),
SalesData AS (
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
),
CombinedPS AS (
    SELECT 
        ps_par,
        ps_comp,
        ps_qty_per
    FROM [QADEE2798].[dbo].[ps_mstr]  -- Only QADEE2798 source
    WHERE [ps_end] IS NULL
),
SFG_Identification AS (
    SELECT 
        ps_par,
        ps_comp,
        ps_qty_per,
        CASE WHEN EXISTS (
            SELECT 1 
            FROM CombinedPS c2 
            WHERE c2.ps_par = c1.ps_comp
        ) THEN 'SFG' END AS [Structure_Type]
    FROM CombinedPS c1
),
BOMHierarchy AS (
    SELECT 
        ps_par AS root_parent,
        ps_par AS current_parent,
        ps_comp AS component,
        ps_qty_per,
        [Structure_Type],
        0 AS LEVEL
    FROM SFG_Identification
    WHERE ps_par NOT IN (
        SELECT ps_comp 
        FROM CombinedPS 
        WHERE ps_comp IS NOT NULL
    )
    
    UNION ALL
    
    SELECT 
        h.root_parent,
        m.ps_par AS current_parent,
        m.ps_comp AS component,
        m.ps_qty_per,
        m.[Structure_Type],
        h.LEVEL + 1
    FROM SFG_Identification m
    INNER JOIN BOMHierarchy h
        ON m.ps_par = h.component
    WHERE LEVEL < 10
),
BOMData AS (
    SELECT 
        h.root_parent AS [Parent_Item],
        h.current_parent AS [ps_par],
        h.component AS [ps_comp],
        h.ps_qty_per AS [Quantity_Per]
    FROM BOMHierarchy h
    WHERE h.[Structure_Type] <> 'SFG' OR h.[Structure_Type] IS NULL
)
-- Join the two datasets and group by SO Project and ps_comp
SELECT 
    bd.[ps_comp] AS [Component],
    SUM(sd.[Total Past Due] * bd.[Quantity_Per]) AS [Total Past Due],
    SUM(sd.[Week 1] * bd.[Quantity_Per]) AS [Week 1],
    SUM(sd.[Week 2] * bd.[Quantity_Per]) AS [Week 2],
    SUM(sd.[Week 3] * bd.[Quantity_Per]) AS [Week 3],
    SUM(sd.[Week 4] * bd.[Quantity_Per]) AS [Week 4],
    SUM(sd.[Week 5] * bd.[Quantity_Per]) AS [Week 5],
    SUM(sd.[Week 6] * bd.[Quantity_Per]) AS [Week 6],
    SUM(sd.[Week 7] * bd.[Quantity_Per]) AS [Week 7],
    SUM(sd.[Week 8] * bd.[Quantity_Per]) AS [Week 8],
    SUM(sd.[Total Long Term] * bd.[Quantity_Per]) AS [Total Long Term]
FROM SalesData sd
LEFT JOIN BOMData bd ON sd.[Item Number] = bd.[Parent_Item]
GROUP BY bd.[ps_comp]
ORDER BY 
    bd.[ps_comp];