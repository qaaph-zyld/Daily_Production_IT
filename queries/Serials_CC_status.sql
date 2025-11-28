-- Step 0: Create a CTE for the item cost data
WITH ItemCost AS (
    SELECT
        [sct_part] AS [Item Number],
        [sct_cst_tot] AS [Standard Cost], 
        ([sct_mtl_tl] + [sct_mtl_ll]) AS [CMAT],
        ([sct_cst_tot] - ([sct_mtl_tl] + [sct_mtl_ll])) AS [LBO],
        CASE 
            WHEN [sct_cst_tot] = 0 THEN ''
            WHEN [sct_cst_tot] = ([sct_mtl_tl] + [sct_mtl_ll]) THEN 'RM'
            ELSE 'SFG/FG'
        END AS [Prod/Mfg]
    FROM 
        [QADEE2798].[dbo].[sct_det]
    WHERE 
        [sct_sim] = 'standard'
),
-- Step 1: Create a CTE for the last receipt data
LastReceipt AS (
    SELECT
        [serh_serial_id] AS [Serial ID],
        [serh_part] AS [Receipt Item Number],
        [serh_qty_chg] AS [Receipt Qty],
        [serh_trans_date] AS [Receipt Date],
        ROW_NUMBER() OVER(
            PARTITION BY [serh_serial_id] 
            ORDER BY [serh_trans_date] DESC, [serh_trans_nbr] DESC
        ) AS rn
    FROM 
        [QADEE2798].[dbo].[serh_hist]
    WHERE 
        [serh_trans_type] IN ('pck-rct', 'pck-bld') 
        AND [serh_stage] <> 'pending'
),
-- Step 2: Create a CTE that gathers all base data and calculates the first dependent column, [WH_age]
BaseData AS (
    SELECT
        -- Columns from the first table (ser_active_picked)
        sap.[ser_serial_id] AS [Serial ID],
        sap.[ser_stage] AS [Stage],
        sap.[ser_pack_code] AS [Pack Code],
        sap.[ser_loc] AS [Location],
        sap.[ser_part] AS [Item Number],
        sap.[ser_lotser] AS [Lot],
        sap.[ser_qty_pck] AS [Standard Pack],
        sap.[ser_qty_avail] AS [Qty],
        sap.[ser_mod_date] AS [Last Modified],

        -- Columns from the second table (serh_hist)
        sh.[serh_pack_code] AS [History Pack Code],
        sh.[serh_trans_date] AS [Last CC date],
        sh.[serh_mod_userid] AS [Modified by],
        sh.[serh_cnt_qty] AS [Counted Qty],
        sh.[serh_trans_nbr],

        -- Columns from the item master table (pt_mstr)
        pt.[pt__chr02] AS [Item Type],
        pt.[pt_desc1] AS [Description],
        pt.[pt_prod_line] AS [Prod Line],
        pt.[pt_group] AS [Group],
        pt.[pt_dsgn_grp] AS [Project],

        -- Columns from the ItemCost CTE
        cost.[Standard Cost],
        cost.[CMAT],
        cost.[Prod/Mfg],

        -- Columns from the LastReceipt CTE
        lr.[Receipt Date],
        lr.[Receipt Qty],

        -- Calculated Column 1: [COGS]
        cost.[Standard Cost] * sap.[ser_qty_avail] AS [COGS],

        -- NEW Calculated Column 2: [WH_age]
        CASE
            WHEN lr.[Receipt Date] IS NULL THEN NULL -- Cannot calculate age without a receipt date
            WHEN DATEDIFF(day, lr.[Receipt Date], GETDATE()) < 60 THEN 'New'
			WHEN DATEDIFF(day, lr.[Receipt Date], GETDATE()) < 180 THEN '< 2-6 months'
			WHEN DATEDIFF(day, lr.[Receipt Date], GETDATE()) < = 365 THEN '< 6-12 months'
            WHEN DATEDIFF(day, lr.[Receipt Date], GETDATE()) > 365 THEN 'Over 1 year'
            ELSE NULL
        END AS [WH_age]
    FROM 
        [QADEE2798].[dbo].[ser_active_picked] AS sap
    LEFT JOIN 
        [QADEE2798].[dbo].[serh_hist] AS sh 
        ON sap.[ser_serial_id] = sh.[serh_master_id] 
        AND sh.[serh_cnt_qty] <> 0
    LEFT JOIN 
        [QADEE2798].[dbo].[pt_mstr] AS pt
        ON sap.[ser_part] = pt.[pt_part]
    LEFT JOIN 
        ItemCost AS cost
        ON sap.[ser_part] = cost.[Item Number]
    LEFT JOIN 
        LastReceipt AS lr
        ON sap.[ser_serial_id] = lr.[Serial ID] 
        AND lr.rn = 1
),
-- Step 3: Create the final CTE that calculates the [CC class] using the [WH_age] column
MainReport AS (
    SELECT
        *,
        -- NEW Calculated Column 3: [CC class] (now depends on [WH_age])
        CASE
            WHEN [Last CC date] IS NULL AND [WH_age] = 'New' THEN 'New'
            WHEN [Last CC date] IS NULL AND [WH_age] <> 'New' THEN 'no CC'
			WHEN DATEDIFF(day, [Last CC date], GETDATE()) < 30 THEN 'OK'
            WHEN DATEDIFF(day, [Last CC date], GETDATE()) < 60 THEN '< 1-2 months'
            WHEN DATEDIFF(day, [Last CC date], GETDATE()) <= 90 THEN '2-3 months'
            WHEN DATEDIFF(day, [Last CC date], GETDATE()) <= 120 THEN '3-4 months'
            WHEN DATEDIFF(day, [Last CC date], GETDATE()) <= 180 THEN '4-6 months'
            WHEN DATEDIFF(day, [Last CC date], GETDATE()) <= 365 THEN '6-12 months'
            ELSE 'more than 1 year'
        END AS [CC class],
        -- The ranking function must be in the final CTE before the SELECT
        ROW_NUMBER() OVER(
            PARTITION BY [Serial ID]         
            ORDER BY [serh_trans_nbr] DESC        
        ) AS rn
    FROM 
        BaseData
)
-- Step 4: Final SELECT to produce the report
SELECT
    [Serial ID],
    [Stage],
    [Location],
    [Item Number],
    [Lot],
    [Qty],
    [Last Modified],
    [Last CC date],
    [Modified by],
    [Counted Qty],
    [Item Type],
    [Description],
    [Prod Line],
    [Group],
    [Project],
    [Standard Cost],
    [CMAT],
    [Prod/Mfg],
    [COGS],
    [Receipt Date],
    [Receipt Qty],
    [WH_age],
    [CC class]
FROM 
    MainReport
WHERE 
    rn = 1
ORDER BY
    [Serial ID];