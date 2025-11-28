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
-- NEW CTE: Calculate pure demand for parent items by subtracting inventory
ParentPureDemand AS (
    SELECT 
        sd.[Item Number],
        -- Calculate pure demand by subtracting MRP Qty from total demand
        -- Ensure we don't get negative values
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] THEN 0
            ELSE sd.[Total Past Due] - inv.[MRP Qty]
        END AS [Pure Past Due],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] THEN sd.[Week 1] - (inv.[MRP Qty] - sd.[Total Past Due])
            ELSE sd.[Week 1]
        END AS [Pure Week 1],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] THEN sd.[Week 2] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1])
            ELSE sd.[Week 2]
        END AS [Pure Week 2],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] THEN sd.[Week 3] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2])
            ELSE sd.[Week 3]
        END AS [Pure Week 3],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] THEN sd.[Week 4] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2] - sd.[Week 3])
            ELSE sd.[Week 4]
        END AS [Pure Week 4],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] THEN sd.[Week 5] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2] - sd.[Week 3] - sd.[Week 4])
            ELSE sd.[Week 5]
        END AS [Pure Week 5],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] THEN sd.[Week 6] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2] - sd.[Week 3] - sd.[Week 4] - sd.[Week 5])
            ELSE sd.[Week 6]
        END AS [Pure Week 6],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] + sd.[Week 7] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] THEN sd.[Week 7] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2] - sd.[Week 3] - sd.[Week 4] - sd.[Week 5] - sd.[Week 6])
            ELSE sd.[Week 7]
        END AS [Pure Week 7],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] + sd.[Week 7] + sd.[Week 8] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] + sd.[Week 7] THEN sd.[Week 8] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2] - sd.[Week 3] - sd.[Week 4] - sd.[Week 5] - sd.[Week 6] - sd.[Week 7])
            ELSE sd.[Week 8]
        END AS [Pure Week 8],
        CASE 
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] + sd.[Week 7] + sd.[Week 8] + sd.[Total Long Term] THEN 0
            WHEN inv.[MRP Qty] >= sd.[Total Past Due] + sd.[Week 1] + sd.[Week 2] + sd.[Week 3] + sd.[Week 4] + sd.[Week 5] + sd.[Week 6] + sd.[Week 7] + sd.[Week 8] THEN sd.[Total Long Term] - (inv.[MRP Qty] - sd.[Total Past Due] - sd.[Week 1] - sd.[Week 2] - sd.[Week 3] - sd.[Week 4] - sd.[Week 5] - sd.[Week 6] - sd.[Week 7] - sd.[Week 8])
            ELSE sd.[Total Long Term]
        END AS [Pure Long Term],
        -- Keep original values for reference
        sd.[Total Past Due],
        sd.[Week 1],
        sd.[Week 2],
        sd.[Week 3],
        sd.[Week 4],
        sd.[Week 5],
        sd.[Week 6],
        sd.[Week 7],
        sd.[Week 8],
        sd.[Total Long Term],
        inv.[MRP Qty]
    FROM SalesData sd
    LEFT JOIN (
        SELECT 
            [in_part] AS [Item Number],
            [in_qty_oh] AS [MRP Qty],
            [in_qty_nonet] AS [Non Nettable],
            [in_qty_oh] + [in_qty_nonet] AS [Total Inv]
        FROM [QADEE2798].[dbo].[15]
    ) inv ON sd.[Item Number] = inv.[Item Number]
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
),
-- MODIFIED CTE: Calculate component demand based on pure parent demand
ComponentData AS (
    SELECT 
        bd.[ps_comp] AS [Item Number],
        SUM(pd.[Pure Past Due] * bd.[Quantity_Per]) AS [Total Past Due],
        SUM(pd.[Pure Week 1] * bd.[Quantity_Per]) AS [Week 1],
        SUM(pd.[Pure Week 2] * bd.[Quantity_Per]) AS [Week 2],
        SUM(pd.[Pure Week 3] * bd.[Quantity_Per]) AS [Week 3],
        SUM(pd.[Pure Week 4] * bd.[Quantity_Per]) AS [Week 4],
        SUM(pd.[Pure Week 5] * bd.[Quantity_Per]) AS [Week 5],
        SUM(pd.[Pure Week 6] * bd.[Quantity_Per]) AS [Week 6],
        SUM(pd.[Pure Week 7] * bd.[Quantity_Per]) AS [Week 7],
        SUM(pd.[Pure Week 8] * bd.[Quantity_Per]) AS [Week 8],
        SUM(pd.[Pure Long Term] * bd.[Quantity_Per]) AS [Total Long Term]
    FROM ParentPureDemand pd
    LEFT JOIN BOMData bd ON pd.[Item Number] = bd.[Parent_Item]
    GROUP BY bd.[ps_comp]
),
CombinedData AS (
    -- Sales data rows (using original demand values)
    SELECT 
        [Item Number],
        [Total Past Due],
        [Week 1],
        [Week 2],
        [Week 3],
        [Week 4],
        [Week 5],
        [Week 6],
        [Week 7],
        [Week 8],
        [Total Long Term]
    FROM SalesData
    
    UNION
    
    -- Component data rows (using calculated component demand)
    SELECT 
        [Item Number],
        [Total Past Due],
        [Week 1],
        [Week 2],
        [Week 3],
        [Week 4],
        [Week 5],
        [Week 6],
        [Week 7],
        [Week 8],
        [Total Long Term]
    FROM ComponentData
),
ItemCount AS (
    SELECT 
        [Item Number],
        COUNT(*) AS [Item Number Count]
    FROM CombinedData
    GROUP BY [Item Number]
),
InventoryData AS (
    SELECT 
        [in_part] AS [Item Number],
        [in_qty_oh] AS [MRP Qty],
        [in_qty_nonet] AS [Non Nettable],
        [Total Inv] = [in_qty_oh] + [in_qty_nonet]
    FROM [QADEE2798].[dbo].[15]
),
LastDates AS (
    SELECT 
        tr_part,
        MAX(CASE WHEN tr_type = 'iss-wo' THEN tr_effdate END) AS LastIssueDate,
        MAX(CASE WHEN tr_type = 'iss-so' THEN tr_effdate END) AS LastSaleDate,
        MAX(CASE WHEN tr_type = 'rct-wo' THEN tr_effdate END) AS LastProdDate,
        MAX(CASE WHEN tr_type = 'rct-po' THEN tr_effdate END) AS LastReceiptDate
    FROM 
        [QADEE2798].[dbo].[tr_hist]
    WHERE 
        tr_type IN ('iss-wo','rct-po','iss-so','rct-wo')
        AND tr_userid <> 'ajelacn'
    GROUP BY 
        tr_part
),
MergedDates AS (
    SELECT 
        tr_part,
        -- Calculate merged issue date (most recent between iss-wo and iss-so)
        CASE 
            WHEN LastIssueDate IS NULL AND LastSaleDate IS NULL THEN NULL
            WHEN LastIssueDate IS NULL THEN LastSaleDate
            WHEN LastSaleDate IS NULL THEN LastIssueDate
            WHEN LastIssueDate > LastSaleDate THEN LastIssueDate
            ELSE LastSaleDate
        END AS MergedIssueDate,
        
        -- Calculate merged receipt date (most recent between rct-wo and rct-po)
        CASE 
            WHEN LastProdDate IS NULL AND LastReceiptDate IS NULL THEN NULL
            WHEN LastProdDate IS NULL THEN LastReceiptDate
            WHEN LastReceiptDate IS NULL THEN LastProdDate
            WHEN LastProdDate > LastReceiptDate THEN LastProdDate
            ELSE LastReceiptDate
        END AS MergedReceiptDate
    FROM 
        LastDates
),
-- Cost Data
CostData AS (
    SELECT 
        [sct_part] as [Item Number],
        [sct_cst_tot] as [Standard Cost], 
        ([sct_mtl_tl] + [sct_mtl_ll]) AS [CMAT],
        ([sct_cst_tot] - ([sct_mtl_tl] + [sct_mtl_ll])) AS [LBO],
        CASE 
            WHEN [sct_cst_tot] = 0 THEN ''
            WHEN [sct_cst_tot] = ([sct_mtl_tl] + [sct_mtl_ll]) THEN 'RM'
            ELSE 'SFG/FG'
        END AS [Prod/Mfg]
    FROM [QADEE2798].[dbo].[sct_det]
    WHERE [sct_sim] = 'standard'
)
SELECT 
    cd.[Item Number],
    cd.[Total Past Due],
    cd.[Week 1],
    cd.[Week 2],
    cd.[Week 3],
    cd.[Week 4],
    cd.[Week 5],
    cd.[Week 6],
    cd.[Week 7],
    cd.[Week 8],
    cd.[Total Long Term],
    inv.[MRP Qty],
    inv.[Non Nettable],
    inv.[Total Inv],
    CASE 
        WHEN inv.[Total Inv] = 0 THEN 0
        ELSE (inv.[Non Nettable] * 100.0 / inv.[Total Inv])
    END AS [NN %],
    CASE 
        WHEN inv.[MRP Qty] < cd.[Total Past Due] THEN 'Past Due'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] THEN 'Week 1'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] THEN 'Week 2'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] THEN 'Week 3'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] + cd.[Week 4] THEN 'Week 4'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] + cd.[Week 4] + cd.[Week 5] THEN 'Week 5'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] + cd.[Week 4] + cd.[Week 5] + cd.[Week 6] THEN 'Week 6'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] + cd.[Week 4] + cd.[Week 5] + cd.[Week 6] + cd.[Week 7] THEN 'Week 7'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] + cd.[Week 4] + cd.[Week 5] + cd.[Week 6] + cd.[Week 7] + cd.[Week 8] THEN 'Week 8'
        WHEN inv.[MRP Qty] < cd.[Total Past Due] + cd.[Week 1] + cd.[Week 2] + cd.[Week 3] + cd.[Week 4] + cd.[Week 5] + cd.[Week 6] + cd.[Week 7] + cd.[Week 8] + cd.[Total Long Term] THEN 'Total Long Term'
        ELSE 'Total Long Term'
    END AS [Coverage],
    -- Categorize merged issue date
    CASE 
        WHEN md.MergedIssueDate IS NULL THEN 'Active'
        WHEN DATEDIFF(day, md.MergedIssueDate, GETDATE()) < 90 THEN 'active'
        WHEN DATEDIFF(day, md.MergedIssueDate, GETDATE()) BETWEEN 91 AND 180 THEN '3 months'
        WHEN DATEDIFF(day, md.MergedIssueDate, GETDATE()) BETWEEN 181 AND 365 THEN '6 months'
        ELSE 'obsolete'
    END AS [Last Issue],
    
    -- Categorize merged receipt date
    CASE 
        WHEN md.MergedReceiptDate IS NULL THEN 'Active'
        WHEN DATEDIFF(day, md.MergedReceiptDate, GETDATE()) < 90 THEN 'active'
        WHEN DATEDIFF(day, md.MergedReceiptDate, GETDATE()) BETWEEN 91 AND 180 THEN '3 months'
        WHEN DATEDIFF(day, md.MergedReceiptDate, GETDATE()) BETWEEN 181 AND 365 THEN '6 months'
        ELSE 'obsolete'
    END AS [Last Receipt],
    -- Cost columns
    cost.[Standard Cost],
    cost.[CMAT],
    cost.[LBO],
    cost.[Prod/Mfg]
FROM CombinedData cd
JOIN ItemCount ic ON cd.[Item Number] = ic.[Item Number]
LEFT JOIN InventoryData inv ON cd.[Item Number] = inv.[Item Number]
LEFT JOIN MergedDates md ON cd.[Item Number] = md.tr_part
LEFT JOIN CostData cost ON cd.[Item Number] = cost.[Item Number]
ORDER BY  
    cd.[Item Number];