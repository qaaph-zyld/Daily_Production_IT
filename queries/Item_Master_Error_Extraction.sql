-- Enhanced Error Detection Query - Production SQL Server Implementation
-- Optimized for performance with proper bracket escaping and error handling
WITH ItemMaster AS (
    SELECT 
        pt__chr02 AS [Item Type],
        pt_part AS [Item Number],
        pt_desc1 AS [Description],
        pt_prod_line AS [Prod Line],
        pt_group AS [Group],
        pt_status AS [Item Status],
        pt_sfty_stk AS [Safety Stock],
        pt_dsgn_grp AS [Project],
        pt_buyer AS [Planner],
        pt_vend AS [Supplier/Customer],
        pt_routing AS [Routing],
        CASE 
            WHEN pt_net_wt_um = 'kg' THEN pt_net_wt
            WHEN pt_net_wt_um = 'g' THEN pt_net_wt / 1000.0
            ELSE NULL
        END AS [Net weight in KG],
        CASE 
            WHEN DATEDIFF(day, pt_added, GETDATE()) < 90 THEN 'New Item'
            ELSE ''
        END AS [New Item]
    FROM QADEE2798.dbo.pt_mstr
),
InventoryData AS (
    SELECT 
        in_part AS [Item Number],
        in_qty_oh AS [MRP Qty],
        in_qty_nonet AS [Non Nettable],
        in_qty_oh + in_qty_nonet AS [Total Inv]
    FROM QADEE2798.dbo.[15] WITH (NOLOCK)
    WHERE in_qty_oh + in_qty_nonet <> 0
),
BasePS AS (
    SELECT ps_par, ps_comp, ps_qty_per, ps_rmks, ps_op, ps_ref
    FROM QADEE2798.dbo.ps_mstr WITH (NOLOCK)
    WHERE ps_end IS NULL AND ps_par IS NOT NULL AND ps_comp IS NOT NULL
),
ComponentLookup AS (
    SELECT DISTINCT ps_comp FROM BasePS
),
SFG_Classification AS (
    SELECT 
        b.ps_par, b.ps_comp, b.ps_qty_per, b.ps_rmks, b.ps_op, b.ps_ref,
        CASE WHEN cl.ps_comp IS NOT NULL THEN 'SFG' ELSE NULL END AS Structure_Type
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
    SELECT item_code, item_type, priority_rank FROM (
        SELECT DISTINCT root_parent AS item_code, 'FG' AS item_type, 1 AS priority_rank
        FROM BOMTraversal
        UNION ALL
        SELECT DISTINCT current_parent AS item_code, 'SFG' AS item_type, 2 AS priority_rank
        FROM BOMTraversal WHERE Structure_Type = 'SFG'
        UNION ALL
        SELECT DISTINCT component AS item_code, 'RM' AS item_type, 3 AS priority_rank
        FROM BOMTraversal WHERE component IS NOT NULL
    ) combined_items
),
BOMItems AS (
    SELECT item_code AS [Item Number], item_type AS [Item_Type_per_BOM] FROM (
        SELECT item_code, item_type,
            ROW_NUMBER() OVER (PARTITION BY item_code ORDER BY priority_rank ASC) AS rn
        FROM ItemClassification WHERE item_code IS NOT NULL
    ) ranked_items WHERE rn = 1
),
-- PO Details from the second query
-- PO Details from the second query
PODetails AS (
    SELECT 
        pod.[pod_nbr] as [PO],
        pod.[pod_line] as [PO line],
        pod.[pod_part] as [Item Number],
        pod.[pod_status] as [PO Line Closed],
        pod.[pod__chr08] as [PO Line Buyer],
        pod.[pod_cum_qty[[1]]]]] as [pod_cum_qty[1]],  -- Correctly escaped
        pod.[pod_plan_weeks],
        pod.[pod_curr_rlse_id[[1]]]]] as [pod_curr_rlse_id[1]],  -- Correctly escaped
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
    WHERE pod.[pod_end_eff[[1]]]]] > CAST(GETDATE() AS DATE)  -- Correctly escaped
),
SupplierTransportDays AS (
    SELECT 
        [Supplier],
        COUNT(DISTINCT pod_translt_days) AS DistinctTransportDaysCount
    FROM PODetails 
    GROUP BY [Supplier]
),
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
        CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END AS [BOM],
        ISNULL(inv.[MRP Qty], 0) AS [MRP Qty],
        ISNULL(inv.[Non Nettable], 0) AS [Non Nettable],
        ISNULL(inv.[Total Inv], 0) AS [Total Inv],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM' 
                AND im.[Item Type] <> bi.[Item_Type_per_BOM] 
            THEN CASE WHEN im.[Item Type] = 'PACK' AND bi.[Item_Type_per_BOM] = 'RM' 
                     THEN '' ELSE 'Error' END
            ELSE ''
        END AS [Item Type Check],
        CASE WHEN ISNULL(inv.[Total Inv], 0) = 0 THEN 'No Inventory' ELSE '' END AS [Inventory Check],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'No BOM' 
                AND ISNULL(inv.[Total Inv], 0) > 0 THEN 'Check' ELSE ''
        END AS [Inventory_BOM Check],
        CASE 
            WHEN (im.[Item Type] = 'FG' OR im.[Item Type] = 'SFG') 
                AND im.[Item Number] <> im.[Routing] THEN 'Error' ELSE ''
        END AS [Routing Check],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM' 
                AND im.[Project] IS NULL THEN 'Error' ELSE ''
        END AS [Project Check],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM' 
                AND im.[Prod Line] = '0000' THEN 'Error' ELSE ''
        END AS [Prod Line Check],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM' 
                AND im.[Group] = 'F000' THEN 'Error' ELSE ''
        END AS [Group Check],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM' 
                AND im.[Planner] IS NULL THEN 'Error' ELSE ''
        END AS [Planner Check Master],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'in BOM' 
                AND bi.[Item_Type_per_BOM] <> 'FG' 
                AND im.[Supplier/Customer] IS NULL THEN 'Error' ELSE ''
        END AS [Supplier/Customer Check],
        pod.[PO],
        pod.[PO line],
        pod.[PO Line Closed],
        pod.[PO Line Buyer],
        pod.[pod_cum_qty_1],
        pod.pod_plan_weeks,
        pod.[pod_curr_rlse_id_1],
        pod.pod_ord_mult,
        pod.pod_translt_days,
        pod.pod_sd_pat,
        pod.pod_plan_mths,
        pod.pod_firm_days,
        pod.pod_sftylt_days,
        pod.po_sd_pat_mstr,
        pod.[Supplier],
        pod.[Supplier Name],
        pod.[Supplier Country],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN im.[Supplier/Customer] = pod.[Supplier] THEN '' ELSE 'error' END
            ELSE ''
        END AS [Supplier Check],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN im.[Planner] = pod.[PO Line Buyer] THEN '' ELSE 'error' END
            ELSE ''
        END AS [Planner Check PO],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN pod.pod_sd_pat = pod.po_sd_pat_mstr THEN '' ELSE 'error' END
            ELSE ''
        END AS [SDP Check],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN pod.pod_plan_weeks = 26 THEN '' ELSE 'error' END
            ELSE ''
        END AS [Weeks Check],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN pod.pod_plan_mths = 6 THEN '' ELSE 'error' END
            ELSE ''
        END AS [Months Check],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN pod.pod_firm_days = 0 THEN '' ELSE 'error' END
            ELSE ''
        END AS [Firm Days Check],
        CASE 
            WHEN bi.[Item_Type_per_BOM] = 'RM' THEN
                CASE WHEN std.DistinctTransportDaysCount > 1 THEN 'error' ELSE '' END
            ELSE ''
        END AS [Transport Days Check],
        CASE 
            WHEN CASE WHEN bi.[Item Number] IS NOT NULL THEN 'in BOM' ELSE 'No BOM' END = 'No BOM' 
                AND ISNULL(inv.[Total Inv], 0) = 0 
                AND im.[New Item] = '' 
            THEN 'Unused/Obsolete' ELSE ''
        END AS [Unused]
    FROM ItemMaster im
    LEFT JOIN BOMItems bi ON im.[Item Number] = bi.[Item Number]
    LEFT JOIN InventoryData inv ON im.[Item Number] = inv.[Item Number]
    LEFT JOIN PODetails pod ON im.[Item Number] = pod.[Item Number]
    LEFT JOIN SupplierTransportDays std ON pod.[Supplier] = std.[Supplier]
),
ErrorRows AS (
    SELECT * FROM MainResultSet
    WHERE [Item Type Check] IN ('Error', 'error') OR
          [Inventory Check] IN ('Error', 'error', 'No Inventory') OR
          [Inventory_BOM Check] IN ('Error', 'error', 'Check') OR
          [Routing Check] IN ('Error', 'error') OR
          [Project Check] IN ('Error', 'error') OR
          [Prod Line Check] IN ('Error', 'error') OR
          [Group Check] IN ('Error', 'error') OR
          [Planner Check Master] IN ('Error', 'error') OR
          [Supplier/Customer Check] IN ('Error', 'error') OR
          [Supplier Check] IN ('Error', 'error') OR
          [Planner Check PO] IN ('Error', 'error') OR
          [SDP Check] IN ('Error', 'error') OR
          [Weeks Check] IN ('Error', 'error') OR
          [Months Check] IN ('Error', 'error') OR
          [Firm Days Check] IN ('Error', 'error') OR
          [Transport Days Check] IN ('Error', 'error') OR
          [Unused] IN ('Unused/Obsolete')
),
UnpivotPrep AS (
    SELECT 
        [Item Number], [Description], [Item Type], [Item_Type_per_BOM], [BOM],
        [Prod Line], [Group], [Project], [Planner], [Supplier/Customer], [Routing],
        [PO], [PO line], [Supplier], [Supplier Name],
        CAST([Item Type Check] AS VARCHAR(100)) AS [Item Type Check],
        CAST([Inventory Check] AS VARCHAR(100)) AS [Inventory Check],
        CAST([Inventory_BOM Check] AS VARCHAR(100)) AS [Inventory_BOM Check],
        CAST([Routing Check] AS VARCHAR(100)) AS [Routing Check],
        CAST([Project Check] AS VARCHAR(100)) AS [Project Check],
        CAST([Prod Line Check] AS VARCHAR(100)) AS [Prod Line Check],
        CAST([Group Check] AS VARCHAR(100)) AS [Group Check],
        CAST([Planner Check Master] AS VARCHAR(100)) AS [Planner Check Master],
        CAST([Supplier/Customer Check] AS VARCHAR(100)) AS [Supplier/Customer Check],
        CAST([Supplier Check] AS VARCHAR(100)) AS [Supplier Check],
        CAST([Planner Check PO] AS VARCHAR(100)) AS [Planner Check PO],
        CAST([SDP Check] AS VARCHAR(100)) AS [SDP Check],
        CAST([Weeks Check] AS VARCHAR(100)) AS [Weeks Check],
        CAST([Months Check] AS VARCHAR(100)) AS [Months Check],
        CAST([Firm Days Check] AS VARCHAR(100)) AS [Firm Days Check],
        CAST([Transport Days Check] AS VARCHAR(100)) AS [Transport Days Check],
        CAST([Unused] AS VARCHAR(100)) AS [Unused]
    FROM ErrorRows
),
UnpivotedErrors AS (
    SELECT 
        [Item Number], [Description], [Item Type], [Item_Type_per_BOM], [BOM],
        [Prod Line], [Group], [Project], [Planner], [Supplier/Customer], [Routing],
        [PO], [PO line], [Supplier], [Supplier Name],
        CheckType, ErrorValue,
        CASE CheckType
            WHEN 'Item Type Check' THEN 1
            WHEN 'Routing Check' THEN 2
            WHEN 'Project Check' THEN 3
            WHEN 'Prod Line Check' THEN 4
            WHEN 'Group Check' THEN 5
            WHEN 'Planner Check Master' THEN 6
            WHEN 'Supplier/Customer Check' THEN 7
            WHEN 'Supplier Check' THEN 8
            WHEN 'Planner Check PO' THEN 9
            WHEN 'SDP Check' THEN 10
            WHEN 'Weeks Check' THEN 11
            WHEN 'Months Check' THEN 12
            WHEN 'Firm Days Check' THEN 13
            WHEN 'Transport Days Check' THEN 14
            WHEN 'Inventory_BOM Check' THEN 15
            WHEN 'Inventory Check' THEN 16
            WHEN 'Unused' THEN 17
            ELSE 99
        END AS ErrorPriority
    FROM UnpivotPrep
    UNPIVOT (
        ErrorValue FOR CheckType IN (
            [Item Type Check], [Inventory Check], [Inventory_BOM Check],
            [Routing Check], [Project Check], [Prod Line Check], [Group Check],
            [Planner Check Master], [Supplier/Customer Check], [Supplier Check],
            [Planner Check PO], [SDP Check], [Weeks Check], [Months Check],
            [Firm Days Check], [Transport Days Check], [Unused]
        )
    ) AS unpvt
    WHERE ErrorValue IN ('Error', 'error', 'Check', 'No Inventory', 'Unused/Obsolete')
)
SELECT 
    [Item Number],
    [Description],
    [Item Type],
    [Item_Type_per_BOM],
    [BOM],
    CheckType AS [Error Type],
    ErrorValue AS [Error Status],
    CASE 
        WHEN CheckType LIKE '%Supplier%' THEN [Supplier]
        WHEN CheckType LIKE '%PO%' OR CheckType IN (
            'SDP Check', 'Weeks Check', 'Months Check', 
            'Firm Days Check', 'Transport Days Check'
        ) THEN CONCAT('PO: ', [PO], ' Line: ', [PO line])
        ELSE NULL
    END AS [Related Info],
    [Prod Line],
    [Group],
    [Project],
    [Planner],
    [Supplier/Customer],
    [Routing],
    ErrorPriority
FROM UnpivotedErrors
ORDER BY 
    ErrorPriority ASC,
    [Item Number],
    CASE WHEN [PO] IS NOT NULL THEN [PO] ELSE '0' END,
    CASE WHEN [PO line] IS NOT NULL THEN [PO line] ELSE 0 END;