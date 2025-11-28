-- Set DATEFIRST at the session level (must be before any CTEs)
SET DATEFIRST 1;

-- Define all CTEs at the same level (no nesting)
WITH 
-- First query CTEs
ld_data AS (
    SELECT 
        xz.[xxwezoned_area_id],
        xz.[xxwezoned_zone_id],
        ld.[ld_loc],
        ld.[ld_part],
        SUM(ld.[ld_qty_oh]) as [ld_qty_oh],
        MAX(ld.[ld_status]) as [ld_status],
        sc.[sct_cst_tot],
        (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [mat_cost],
        (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
        SUM(ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],
        SUM(ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT]
    FROM 
        [QADEE2798].[dbo].[ld_det] ld
    JOIN 
        [QADEE2798].[dbo].[xxwezoned_det] xz
    ON 
        ld.[ld_loc] = xz.[xxwezoned_loc]
    JOIN 
        (
            SELECT
                [sct_site],
                [sct_part],
                [sct_cst_tot],
                [sct_mtl_tl],
                [sct_mtl_ll]
            FROM 
                [QADEE2798].[dbo].[sct_det]
            WHERE 
                [sct_sim] = 'standard'
        ) sc
    ON 
        ld.[ld_part] = sc.[sct_part] 
        AND ld.[ld_site] = sc.[sct_site]
    GROUP BY
        xz.[xxwezoned_area_id],
        xz.[xxwezoned_zone_id],
        ld.[ld_loc],
        ld.[ld_part],
        sc.[sct_cst_tot],
        sc.[sct_mtl_tl],
        sc.[sct_mtl_ll]
),

ser_data AS (
    SELECT 
        [ser_loc] as [Location],
        [ser_part] as [Item Number],
        SUM([ser_qty_avail]) AS [Serial Qty],
        COUNT([ser_serial_id]) AS count_ser_serial_id
    FROM 
        [QADEE2798].[dbo].[ser_active_picked]
    GROUP BY 
        [ser_part],
        [ser_loc]
),

InventoryData AS (
    SELECT 
        pt.[pt_part] as [Item Number],
        pt.[pt_desc1] as [Description],
        pt.[pt_prod_line] as [Prod Line],
        pt.[pt_group] as [Group],
        pt.[pt_status] as [Item Number Status],
        pt.[pt_sfty_stk] as [Safety Stock],
        pt.[pt_dsgn_grp] as [Project],
        pt.[pt_buyer] as [Planner],
        pt.[pt_vend] as [Supplier/Customer],
        pt.[pt_routing] as [Routing],
        pt.[pt__chr02] as [Item Type],
        ld_data.[xxwezoned_area_id] as [Area],
        ld_data.[xxwezoned_zone_id] as [Zone],
        ld_data.[ld_loc] as [Location],
        ld_data.[ld_qty_oh] as [Inventory per location],
        ld_data.[ld_status] as [Inventory Status],
        ld_data.[sct_cst_tot] as [Standard Cost],
        ld_data.[mat_cost] as [Material cost],
        ld_data.[LBO],
        ld_data.[COGS],
        ld_data.[CMAT],
        ser_data.[Serial Qty],
        ser_data.[count_ser_serial_id],
        ld_data.[ld_qty_oh] - ISNULL(ser_data.[Serial Qty], 0) AS [Loose Inv],
        CASE 
            WHEN ld_data.[mat_cost] IS NOT NULL 
            THEN (ld_data.[ld_qty_oh] - ISNULL(ser_data.[Serial Qty], 0)) * ld_data.[mat_cost]
            ELSE 0 
        END AS [CMAT_Loose]
    FROM 
        [QADEE2798].[dbo].[pt_mstr] pt
    LEFT JOIN ld_data ON pt.[pt_part] = ld_data.[ld_part]
    LEFT JOIN ser_data ON pt.[pt_part] = ser_data.[Item Number] 
        AND ld_data.[ld_loc] = ser_data.[Location]
    WHERE 
        ld_data.[ld_qty_oh] IS NOT NULL 
        AND ld_data.[ld_qty_oh] <> 0
),

-- Second query CTEs
BaseData AS (
    SELECT 
        sod.[sod_part] AS [Item Number],
        schd.[schd_date] AS [Date],
        CAST(schd.[schd_discr_qty] AS INT) AS [Discrete Qty],
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
    FROM [QADEE2798].[dbo].[ps_mstr]
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

ComponentData AS (
    SELECT 
        bd.[ps_comp] AS [Item Number],
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
),

CombinedData AS (
    -- Sales data rows
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
    
    -- Component data rows
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

InventoryData2 AS (
    SELECT 
        [in_part] AS [Item Number],
        [in_qty_oh] AS [MRP Qty],
        [in_qty_nonet] AS [Non Nettable],
        [in_qty_oh] + [in_qty_nonet] AS [Total Inv]
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

SalesForecastData AS (
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
        END AS [Last Receipt]
    FROM CombinedData cd
    JOIN ItemCount ic ON cd.[Item Number] = ic.[Item Number]
    LEFT JOIN InventoryData2 inv ON cd.[Item Number] = inv.[Item Number]
    LEFT JOIN MergedDates md ON cd.[Item Number] = md.tr_part
)

-- Final query joining both CTEs
SELECT 
    inv.[Item Number],
    inv.[Description],
    inv.[Prod Line],
    inv.[Group],
    inv.[Item Number Status],
    inv.[Safety Stock],
    inv.[Project],
    inv.[Planner],
    inv.[Supplier/Customer],
    inv.[Routing],
    inv.[Item Type],
    inv.[Area],
    inv.[Zone],
    inv.[Location],
    inv.[Inventory per location],
    inv.[Inventory Status],
    inv.[Standard Cost],
    inv.[Material cost],
    inv.[LBO],
    inv.[COGS],
    inv.[CMAT],
    inv.[Serial Qty],
    inv.[count_ser_serial_id],
    inv.[Loose Inv],
    inv.[CMAT_Loose],
    sales.[Total Past Due],
    sales.[Week 1],
    sales.[Week 2],
    sales.[Week 3],
    sales.[Week 4],
    sales.[Week 5],
    sales.[Week 6],
    sales.[Week 7],
    sales.[Week 8],
    sales.[Total Long Term],
    sales.[MRP Qty],
    sales.[Non Nettable],
    sales.[Total Inv],
    sales.[NN %],
    sales.[Coverage],
    sales.[Last Issue],
    sales.[Last Receipt]
FROM 
    InventoryData inv
LEFT JOIN 
    SalesForecastData sales ON inv.[Item Number] = sales.[Item Number]
ORDER BY 
    inv.[Item Number];