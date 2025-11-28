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
)
SELECT distinct
     h.root_parent AS [root_parent],
    h.ps_op,
    h.[Structure_Type],
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
    CASE 
        WHEN h.component IS NOT NULL THEN 'in BOM'
        ELSE 'NO BOM'
    END AS [BOM Checker]
FROM BOMHierarchy h
LEFT JOIN ItemMaster im ON  h.root_parent = im.[Item Number]
ORDER BY
    h.root_parent;