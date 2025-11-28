WITH CombinedPS AS (
    SELECT 
        '2798' AS [Plant],
        ps_par,
        ps_comp,
        ps_qty_per,
        ps_rmks,
        ps_op,
        ps_ref
    FROM [QADEE2798].[dbo].[ps_mstr]
    WHERE [ps_end] IS NULL
),
SFG_Identification AS (
    SELECT 
        ps_par,
        ps_comp,
        [Plant],
        ps_qty_per,
        ps_rmks,
        ps_op,
        ps_ref,
        CASE WHEN EXISTS (
            SELECT 1 
            FROM CombinedPS c2 
            WHERE c2.ps_par = c1.ps_comp
            AND c2.[Plant] = c1.[Plant]
        ) THEN 'SFG' END AS [Structure_Type]
    FROM CombinedPS c1
),
BOMHierarchy AS (
    SELECT 
        ps_par AS root_parent,
        ps_par AS current_parent,
        ps_comp AS component,
        [Plant],
        ps_qty_per,
        ps_rmks,
        ps_op,
        ps_ref,
        [Structure_Type],
        0 AS LEVEL
    FROM SFG_Identification
    WHERE ps_par NOT IN (
        SELECT ps_comp 
        FROM CombinedPS 
        WHERE ps_comp IS NOT NULL
        AND [Plant] = SFG_Identification.[Plant]
    )
    
    UNION ALL
    
    SELECT 
        h.root_parent,
        m.ps_par AS current_parent,
        m.ps_comp AS component,
        h.[Plant],
        m.ps_qty_per,
        m.ps_rmks,
        m.ps_op,
        m.ps_ref,
        m.[Structure_Type],
        h.LEVEL + 1
    FROM SFG_Identification m
    INNER JOIN BOMHierarchy h
        ON m.ps_par = h.component
        AND m.[Plant] = h.[Plant]
    WHERE LEVEL < 10
),
ItemMaster AS (
    SELECT 
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
            ELSE NULL
        END as [Net weight in KG],
        CASE 
            WHEN DATEDIFF(day, [pt_added], GETDATE()) < 90 THEN 'New Item'
            ELSE ''
        END as [New Item]
    FROM [QADEE2798].[dbo].[pt_mstr]
),
ItemOperationCount AS (
    SELECT 
        [Item Number],
        COUNT(DISTINCT ps_op) AS OperationCount
    FROM BOMHierarchy h
    LEFT JOIN ItemMaster im ON h.component = im.[Item Number]
    GROUP BY [Item Number]
),
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
SupplierTransportDays AS (
    SELECT 
        [Supplier],
        COUNT(DISTINCT [pod_translt_days]) AS DistinctTransportDaysCount
    FROM PODetails
    GROUP BY [Supplier]
)
SELECT distinct
    pod.[Item Number],
    im.[Item Number],
    h.ps_op,
    h.[Structure_Type],
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
    CASE 
        WHEN h.component IS NOT NULL THEN 'in BOM'
        ELSE 'NO BOM'
    END AS [BOM Checker],
    CASE 
        WHEN ioc.OperationCount > 1 THEN 'Check Operation'
        ELSE ''
    END AS [Operation check],
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
    CASE 
        WHEN im.[Supplier/Customer] = pod.[Supplier] THEN ''
        ELSE 'error'
    END AS [Supplier Check],
    CASE 
        WHEN im.[Planner] = pod.[PO Line Buyer] THEN ''
        ELSE 'error'
    END AS [Planner Check],
    CASE 
        WHEN pod.[pod_sd_pat] = pod.[po_sd_pat_mstr] THEN ''
        ELSE 'error'
    END AS [SDP Check],
    CASE 
        WHEN pod.[pod_plan_weeks] = 26 THEN ''
        ELSE 'error'
    END AS [Weeks Check],
    CASE 
        WHEN pod.[pod_plan_mths] = 6 THEN ''
        ELSE 'error'
    END AS [Months Check],
    CASE 
        WHEN pod.[pod_firm_days] = 0 THEN ''
        ELSE 'error'
    END AS [Firm Days Check],
    CASE 
        WHEN std.DistinctTransportDaysCount > 1 THEN 'error'
        ELSE ''
    END AS [Transport Days Check]
FROM BOMHierarchy h
LEFT JOIN ItemMaster im ON h.component = im.[Item Number]
LEFT JOIN ItemOperationCount ioc ON im.[Item Number] = ioc.[Item Number]
RIGHT JOIN PODetails pod ON im.[Item Number] = pod.[Item Number]
LEFT JOIN SupplierTransportDays std ON pod.[Supplier] = std.[Supplier]
ORDER BY
    pod.[Item Number];