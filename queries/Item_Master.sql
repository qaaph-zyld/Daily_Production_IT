SET DATEFIRST 1; -- Set Monday as the first day of the week (ISO standard)
-- First query with all the CTEs
WITH ItemMaster AS (
    SELECT 
        [pt__chr02] as [Item Type],
        [pt_part] as [Item Number],
        [pt_desc1] as [Description],
        [pt_prod_line] as [Prod Line],
        [pt_group] as [Group],
        [pt_status] as [Item Status],
        [pt_sfty_stk] as [Safety Stock],
        [pt_dsgn_grp] as [Project],
        [pt_buyer] as [Planner],
        [pt_vend] as [Supplier/Customer],
        [pt_routing] as [Routing],
        CASE 
            WHEN [pt_net_wt_um] = 'kg' THEN [pt_net_wt]
            WHEN [pt_net_wt_um] = 'g' THEN [pt_net_wt] / 1000.0
            ELSE NULL  -- For any other unit of measure
        END as [Net weight in KG],
        CASE 
            WHEN DATEDIFF(day, [pt_added], GETDATE()) < 90 THEN 'New Item'
            ELSE ''
        END as [New Item]
    FROM [QADEE2798].[dbo].[pt_mstr]
),
-- Inventory Data
InventoryData AS (
    SELECT 
        [in_part] AS [Item Number],
        [in_qty_oh] AS [MRP Qty],
        [in_qty_nonet] AS [Non Nettable],
        [in_qty_oh] + [in_qty_nonet] AS [Total Inv]
    FROM 
        [QADEE2798].[dbo].[15] WITH (NOLOCK)
    WHERE 
        [in_qty_oh] + [in_qty_nonet] <> 0
),
-- Standard Cost Calculation
StandardCost AS (
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
),
-- BOM Item Classification Query
BasePS AS (
    SELECT 
        ps_par,
        ps_comp,
        ps_qty_per,
        ps_rmks,
        ps_op,
        ps_ref
    FROM [QADEE2798].[dbo].[ps_mstr] WITH (NOLOCK)
    WHERE ps_end IS NULL
        AND ps_par IS NOT NULL 
        AND ps_comp IS NOT NULL
),
ComponentLookup AS (
    SELECT DISTINCT ps_comp
    FROM BasePS
),
SFG_Classification AS (
    SELECT 
        b.ps_par,
        b.ps_comp,
        b.ps_qty_per,
        b.ps_rmks,
        b.ps_op,
        b.ps_ref,
        CASE 
            WHEN cl.ps_comp IS NOT NULL THEN 'SFG'
            ELSE NULL 
        END AS Structure_Type
    FROM BasePS b
    LEFT JOIN ComponentLookup cl ON b.ps_par = cl.ps_comp
),
RootParents AS (
    SELECT DISTINCT ps_par AS root_item
    FROM SFG_Classification s1
    WHERE NOT EXISTS (
        SELECT 1 FROM BasePS s2 WHERE s2.ps_comp = s1.ps_par
    )
),
BOMTraversal AS (
    SELECT 
        r.root_item AS root_parent,
        s.ps_par AS current_parent,
        s.ps_comp AS component,
        s.Structure_Type,
        1 AS level_depth
    FROM RootParents r
    INNER JOIN SFG_Classification s ON r.root_item = s.ps_par
    
    UNION ALL
    
    SELECT 
        h.root_parent,
        s.ps_par AS current_parent,
        s.ps_comp AS component,
        s.Structure_Type,
        h.level_depth + 1
    FROM SFG_Classification s
    INNER JOIN BOMTraversal h ON s.ps_par = h.component
    WHERE h.level_depth < 8
),
ItemClassification AS (
    SELECT item_code, item_type, priority_rank
    FROM (
        SELECT DISTINCT 
            root_parent AS item_code,
            'FG' AS item_type,
            1 AS priority_rank
        FROM BOMTraversal
        
        UNION ALL
        
        SELECT DISTINCT 
            current_parent AS item_code,
            'SFG' AS item_type,
            2 AS priority_rank
        FROM BOMTraversal
        WHERE Structure_Type = 'SFG'
        
        UNION ALL
        
        SELECT DISTINCT 
            component AS item_code,
            'RM' AS item_type,
            3 AS priority_rank
        FROM BOMTraversal
        WHERE component IS NOT NULL
    ) combined_items
),
BOMItems AS (
    SELECT 
        item_code AS [Item Number],
        item_type AS [Item_Type_per_BOM]
    FROM (
        SELECT 
            item_code,
            item_type,
            ROW_NUMBER() OVER (
                PARTITION BY item_code 
                ORDER BY priority_rank ASC
            ) AS rn
        FROM ItemClassification
        WHERE item_code IS NOT NULL
    ) ranked_items
    WHERE rn = 1
),
-- Transaction History Data
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
        AND tr_effdate < DATEADD(day, -365, GETDATE())
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
TransactionHistory AS (
    SELECT 
        tr_part as [Item Number],
        -- Categorize merged issue date
        CASE 
            WHEN MergedIssueDate IS NULL THEN 'No transactions'
            WHEN DATEDIFF(day, MergedIssueDate, GETDATE()) < 90 THEN 'active'
            WHEN DATEDIFF(day, MergedIssueDate, GETDATE()) BETWEEN 91 AND 180 THEN '3 months'
            WHEN DATEDIFF(day, MergedIssueDate, GETDATE()) BETWEEN 181 AND 365 THEN '6 months'
            ELSE 'obsolete'
        END AS [Last Issue],
        
        -- Categorize merged receipt date
        CASE 
            WHEN MergedReceiptDate IS NULL THEN 'No transactions'
            WHEN DATEDIFF(day, MergedReceiptDate, GETDATE()) < 90 THEN 'active'
            WHEN DATEDIFF(day, MergedReceiptDate, GETDATE()) BETWEEN 91 AND 180 THEN '3 months'
            WHEN DATEDIFF(day, MergedReceiptDate, GETDATE()) BETWEEN 181 AND 365 THEN '6 months'
            ELSE 'obsolete'
        END AS [Last Receipt]
    FROM 
        MergedDates
),
-- PO Details
PODetails AS (
    SELECT 
        pod.[pod_nbr] as [PO],
        pod.[pod_line] as [PO line],
        pod.[pod_part] as [Item Number],
        pod.[pod_status] as [PO Line Closed],
        pod.[pod__chr08] as [PO Line Buyer],
        pod.[pod_cum_qty[1]]],
        pod.[pod_plan_weeks],
        pod.[pod_curr_rlse_id[1]]],
        pod.[pod_ord_mult],
        pod.[pod_translt_days],
        pod.[pod_sd_pat],
        pod.[pod_plan_mths],
        pod.[pod_firm_days],
        pod.[pod_sftylt_days],
        po.[po_sd_pat] as [po_sd_pat_mstr],
        ad.[ad_addr] as [Supplier],
        ad.[ad_name] as [Supplier Name],
        ad.[ad_country] as [Supplier Country]
    FROM [QADEE2798].[dbo].[pod_det] pod
    LEFT JOIN [QADEE2798].[dbo].[po_mstr] po
        ON pod.[pod_nbr] = po.[po_nbr]
    LEFT JOIN [QADEE2798].[dbo].[ad_mstr] ad
        ON po.[po_vend] = ad.[ad_addr]
    WHERE pod.[pod_end_eff[1]]] > CAST(GETDATE() AS DATE)
),
-- Supplier Transport Days
SupplierTransportDays AS (
    SELECT 
        [Supplier],
        COUNT(DISTINCT [pod_translt_days]) AS DistinctTransportDaysCount
    FROM PODetails
    GROUP BY [Supplier]
),
-- COGS by Area (new CTE from the second query)
COGSByArea AS (
    SELECT 
        pt.[pt_part] as [Item Number],
        SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WH' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WH],
        SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WH-FG' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WH_FG],
        SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WIP' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WIP],
        SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'EXLPICK' THEN ld_data.[COGS] ELSE 0 END) as [COGS_EXLPICK],
        SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WH-FG-E' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WH_FG_E]
    FROM 
        [QADEE2798].[dbo].[pt_mstr] pt
    LEFT JOIN 
        (
            SELECT 
                xz.[xxwezoned_area_id],
                ld.[ld_part],
                SUM(ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS]
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
                        [sct_cst_tot]
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
                ld.[ld_part]
        ) ld_data
    ON 
        pt.[pt_part] = ld_data.[ld_part]
    LEFT JOIN
        (
            SELECT 
                [ser_part] as [Item Number]
            FROM 
                [QADEE2798].[dbo].[ser_active_picked]
            GROUP BY 
                [ser_part]
        ) ser_data
    ON 
        pt.[pt_part] = ser_data.[Item Number]
    WHERE 
        ld_data.[COGS] IS NOT NULL 
        AND ld_data.[COGS] <> 0
        AND ld_data.[xxwezoned_area_id] IN ('WH', 'WH-FG', 'WIP', 'EXLPICK', 'WH-FG-E')
    GROUP BY
        pt.[pt_part]
),
-- Main Result Set
MainResultSet AS (
    SELECT 
        im.[Item Type],
        im.[Item Number],
        im.[Description],
        im.[Prod Line],
        im.[Group],
        im.[Item Status],
        im.[Safety Stock],
        im.[Project],
        im.[Planner],
        im.[Supplier/Customer],
        im.[Routing],
        im.[Net weight in KG],
        im.[New Item],
        bi.[Item_Type_per_BOM],
        CASE 
            WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM'
            ELSE 'No BOM'
        END AS [BOM],
        ISNULL(inv.[MRP Qty], 0) AS [MRP Qty],
        ISNULL(inv.[Non Nettable], 0) AS [Non Nettable],
        ISNULL(inv.[Total Inv], 0) AS [Total Inv],
        -- Updated Item Type Check with PACK/RM exception
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND (
                    (im.[Item Type] IS NULL AND bi.[Item_Type_per_BOM] IS NOT NULL AND ISNULL(inv.[Total Inv], 0) <> 0)
                    OR 
                    (
                        im.[Item Type] IS NOT NULL 
                        AND im.[Item Type] <> bi.[Item_Type_per_BOM]
                        AND NOT (im.[Item Type] = 'PACK' AND bi.[Item_Type_per_BOM] = 'RM')
                    )
                )
            THEN 'Error'
            ELSE ''
        END AS [Item Type Check],
        CASE 
            WHEN ISNULL(inv.[Total Inv], 0) = 0 THEN 'No Inventory'
            ELSE ''
        END AS [Inventory Check],
        -- Updated Inventory_BOM Check to exclude false positive
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'No BOM') 
                AND ISNULL(inv.[Total Inv], 0) > 0 
                AND im.[Prod Line] <> 'OBS'  -- Exclude the false positive case
            THEN 'Check'
            ELSE ''
        END AS [Inventory_BOM Check],
        CASE 
            WHEN (im.[Item Type] = 'FG' OR im.[Item Type] = 'SFG') 
                AND im.[Item Number] <> im.[Routing] THEN 'Error'
            ELSE ''
        END AS [Routing Check],
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND im.[Project] IS NULL 
            THEN 'Error'
            ELSE ''
        END AS [Project Check],
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND im.[Prod Line] = '0000' 
            THEN 'Error'
            ELSE ''
        END AS [Prod Line Check],
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND im.[Group] = 'F000' 
            THEN 'Error'
            ELSE ''
        END AS [Group Check],
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND im.[Planner] IS NULL 
            THEN 'Error'
            ELSE ''
        END AS [Planner Check (BOM)],
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND bi.[Item_Type_per_BOM] <> 'FG' 
                AND im.[Supplier/Customer] IS NULL 
            THEN 'Error'
            ELSE ''
        END AS [Supplier/Customer Check (BOM)],
        -- Add Standard Cost columns here
        sc.[Standard Cost],
        sc.[CMAT],
        sc.[LBO],
        sc.[Prod/Mfg],
        -- New column: Obs - Active Check
        CASE 
            WHEN 
                (CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM') 
                AND im.[Prod Line] = 'OBS' 
            THEN 'Check'
            ELSE ''
        END AS [Obs - Active Check]
    FROM ItemMaster im
    LEFT JOIN BOMItems bi ON im.[Item Number] = bi.[Item Number]
    LEFT JOIN InventoryData inv ON im.[Item Number] = inv.[Item Number]
    LEFT JOIN StandardCost sc ON im.[Item Number] = sc.[Item Number]
),
-- Final combined result from first query
FinalMainResultSet AS (
    SELECT 
        mrs.*,  -- All columns from the main result set including Standard Cost columns
        th.[Last Issue],
        th.[Last Receipt],
        pod.[PO],
        pod.[PO line],
        pod.[PO Line Closed],
        pod.[PO Line Buyer],
        pod.[pod_cum_qty[1]]],
        pod.[pod_plan_weeks],
        pod.[pod_curr_rlse_id[1]]],
        pod.[pod_ord_mult],
        pod.[pod_translt_days],
        pod.[pod_sd_pat],
        pod.[pod_plan_mths],
        pod.[pod_firm_days],
        pod.[pod_sftylt_days],
        pod.[po_sd_pat_mstr],
        pod.[Supplier],
        pod.[Supplier Name],
        pod.[Supplier Country],
        -- Updated check columns with RM caveat
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' THEN
                CASE 
                    WHEN mrs.[Supplier/Customer] = pod.[Supplier] THEN ''
                    ELSE 'error'
                END
            ELSE ''
        END AS [Supplier Check (PO)],
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' THEN
                CASE 
                    WHEN mrs.[Planner] = pod.[PO Line Buyer] THEN ''
                    ELSE 'error'
                END
            ELSE ''
        END AS [Planner Check (PO)],
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' THEN
                CASE 
                    WHEN pod.[pod_sd_pat] = pod.[po_sd_pat_mstr] THEN ''
                    ELSE 'error'
                END
            ELSE ''
        END AS [SDP Check],
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' THEN
                CASE 
                    WHEN pod.[pod_plan_weeks] = 26 THEN ''
                    ELSE 'error'
                END
            ELSE ''
        END AS [Weeks Check],
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' and pod.[PO] is not null THEN
                CASE 
                    WHEN pod.[pod_plan_mths] = 6 THEN ''
                    ELSE 'error'
                END
            ELSE ''
        END AS [Months Check],
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' and pod.[PO] is not null THEN
                CASE 
                    WHEN pod.[pod_firm_days] = 0 THEN ''
                    ELSE 'error'
                END
            ELSE ''
        END AS [Firm Days Check],
        CASE 
            WHEN mrs.[Item_Type_per_BOM] = 'RM' THEN
                CASE 
                    WHEN std.DistinctTransportDaysCount > 1 THEN 'error'
                    ELSE ''
                END
            ELSE ''
        END AS [Transport Days Check],
        -- New Unused/Obsolete column
        CASE 
            WHEN mrs.[BOM] = 'No BOM' 
                AND mrs.[Inventory Check] = 'No Inventory' 
                AND mrs.[New Item] = '' 
                AND (th.[Last Issue] = 'obsolete' OR th.[Last Issue] = 'No transactions')
                AND (th.[Last Receipt] = 'obsolete' OR th.[Last Receipt] = 'No transactions')
            THEN 'Unused/Obsolete'
            ELSE ''
        END AS [Unused],
        -- Cost Calculations (now using mrs columns)
        ISNULL(mrs.[MRP Qty], 0) * ISNULL(mrs.[Standard Cost], 0) AS [MRP COGS],
        ISNULL(mrs.[Non Nettable], 0) * ISNULL(mrs.[Standard Cost], 0) AS [Non Nettable Cogs],
        ISNULL(mrs.[Total Inv], 0) * ISNULL(mrs.[Standard Cost], 0) AS [Total Inv Cogs],
        ISNULL(mrs.[MRP Qty], 0) * ISNULL(mrs.[CMAT], 0) AS [MRP CMAT],
        ISNULL(mrs.[Non Nettable], 0) * ISNULL(mrs.[CMAT], 0) AS [Non Netttable CMAT],
        ISNULL(mrs.[Total Inv], 0) * ISNULL(mrs.[CMAT], 0) AS [Total INV CMAT]
    FROM MainResultSet mrs
    LEFT JOIN TransactionHistory th ON mrs.[Item Number] = th.[Item Number]
    LEFT JOIN PODetails pod ON mrs.[Item Number] = pod.[Item Number]
    LEFT JOIN SupplierTransportDays std ON pod.[Supplier] = std.[Supplier]
)
-- Now we need to define the demand data CTEs separately
-- Second query with demand data
, BaseData AS (
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
DemandInventoryData AS (
    SELECT 
        [in_part] AS [Item Number],
        [in_qty_oh] AS [MRP Qty],
        [in_qty_nonet] AS [Non Nettable],
        [Total Inv] = [in_qty_oh] + [in_qty_nonet]
    FROM [QADEE2798].[dbo].[15]
),
DemandLastDates AS (
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
DemandMergedDates AS (
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
        DemandLastDates
),
DemandData AS (
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
    LEFT JOIN DemandInventoryData inv ON cd.[Item Number] = inv.[Item Number]
    LEFT JOIN DemandMergedDates md ON cd.[Item Number] = md.tr_part
)
-- Final result joining both queries
SELECT 
    fmrs.*,
    -- COGS by Area columns
    ISNULL(cba.COGS_WH, 0) AS [COGS_WH],
    ISNULL(cba.COGS_WH_FG, 0) AS [COGS_WH_FG],
    ISNULL(cba.COGS_WIP, 0) AS [COGS_WIP],
    ISNULL(cba.COGS_EXLPICK, 0) AS [COGS_EXLPICK],
    ISNULL(cba.COGS_WH_FG_E, 0) AS [COGS_WH_FG_E],
    -- Demand columns with "No Demand" for null values
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Total Past Due] AS VARCHAR(50))
    END AS [Demand_Total Past Due],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 1] AS VARCHAR(50))
    END AS [Demand_Week 1],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 2] AS VARCHAR(50))
    END AS [Demand_Week 2],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 3] AS VARCHAR(50))
    END AS [Demand_Week 3],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 4] AS VARCHAR(50))
    END AS [Demand_Week 4],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 5] AS VARCHAR(50))
    END AS [Demand_Week 5],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 6] AS VARCHAR(50))
    END AS [Demand_Week 6],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 7] AS VARCHAR(50))
    END AS [Demand_Week 7],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Week 8] AS VARCHAR(50))
    END AS [Demand_Week 8],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE CAST(dd.[Total Long Term] AS VARCHAR(50))
    END AS [Demand_Total Long Term],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE dd.[Coverage]
    END AS [Demand_Coverage],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE dd.[Last Issue]
    END AS [Demand_Last Issue],
    CASE 
        WHEN dd.[Item Number] IS NULL THEN 'No Demand'
        ELSE dd.[Last Receipt]
    END AS [Demand_Last Receipt]
FROM FinalMainResultSet fmrs
LEFT JOIN COGSByArea cba ON fmrs.[Item Number] = cba.[Item Number]
LEFT JOIN DemandData dd ON fmrs.[Item Number] = dd.[Item Number]
ORDER BY 
    CASE 
        WHEN fmrs.[Item_Type_per_BOM] = 'FG' THEN 1 
        WHEN fmrs.[Item_Type_per_BOM] = 'SFG' THEN 2 
        WHEN fmrs.[Item_Type_per_BOM] = 'RM' THEN 3 
        ELSE 4
    END,
    fmrs.[Item Number],
    fmrs.[PO],
    fmrs.[PO line];