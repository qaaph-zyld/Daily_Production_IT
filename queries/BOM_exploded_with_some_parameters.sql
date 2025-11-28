WITH CombinedPS AS (
    SELECT 
        '2798' AS [Plant],  -- Single plant value now
        ps_par,
        ps_comp,
        ps_qty_per,
        ps_rmks,
        ps_op,
        ps_ref
    FROM [QADEE2798].[dbo].[ps_mstr]  -- Only QADEE2798 source
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
ParentItemMaster AS (
    SELECT 
        [pt__chr02] as [Item Type],
        [pt_part] as [Item Number],
        [pt_desc1] as [Description],
        [pt_desc2] as [Description 2],
        [pt_prod_line] as [Prod Line],
        [pt_group] as [Group],
        [pt_dsgn_grp] as [Project],
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
ComponentItemMaster AS (
    SELECT 
        [pt_part] as [Item Number],
        [pt_desc1] as [Description],
        CASE 
            WHEN [pt_net_wt_um] = 'kg' THEN [pt_net_wt]
            WHEN [pt_net_wt_um] = 'g' THEN [pt_net_wt] / 1000.0
            ELSE NULL  -- For any other unit of measure
        END as [Net weight in KG]
    FROM [QADEE2798].[dbo].[pt_mstr]
)
SELECT 
    h.root_parent AS [Parent_Item],
    h.current_parent AS [ps_par],
    h.component AS [ps_comp],
    h.ps_qty_per AS [Quantity_Per],
    h.ps_rmks,
    h.ps_op,
    h.[Structure_Type],
    -- Parent item columns
    pim.[Item Type] AS [Parent Item Type],
    pim.[Description] AS [Parent Description],
    pim.[Description 2] AS [Parent Description 2],
    pim.[Prod Line] AS [Parent Prod Line],
    pim.[Group] AS [Parent Group],
    pim.[Project] AS [Parent Project],
    pim.[Net weight in KG] AS [Parent Net weight in KG],
    -- Component item columns
    cim.[Description] AS [Component Description]
FROM BOMHierarchy h
LEFT JOIN ParentItemMaster pim ON h.root_parent = pim.[Item Number]
LEFT JOIN ComponentItemMaster cim ON h.component = cim.[Item Number]
ORDER BY 
    h.[Plant],
    h.root_parent,
    h.current_parent,
    h.component;