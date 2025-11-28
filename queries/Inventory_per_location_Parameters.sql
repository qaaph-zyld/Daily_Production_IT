WITH ParentCTE AS (
    SELECT DISTINCT  
        '2674' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]  -- Adding the Parent column  
    FROM   
        [QADEE].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  

    UNION ALL  

    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]  -- Adding the Parent column  
    FROM   
        [QADEE2798].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  
),
ChildCTE AS (
    SELECT DISTINCT  
        '2674' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]  -- Adding the Child column  
    FROM   
        [QADEE].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  

    UNION ALL  

    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]  -- Adding the Child column  
    FROM   
        [QADEE2798].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  
),
BOMStatusCTE AS (
    SELECT 
        COALESCE(p.[Plant], c.[Plant]) AS [Plant],  
        COALESCE(p.[Item Number], c.[Item Number]) AS [Item Number],  
        ISNULL(p.[Parent], 'No') AS [Parent],  
        ISNULL(c.[Child], 'No') AS [Child],  
        CASE 
            WHEN p.[Parent] = 'Yes' AND c.[Child] = 'Yes' THEN 'Yes' 
            ELSE 'No' 
        END AS [SFG]  -- Determine if the item is SFG
    FROM 
        ParentCTE p
    FULL OUTER JOIN 
        ChildCTE c
    ON 
        p.[Plant] = c.[Plant] 
        AND p.[Item Number] = c.[Item Number]
)
SELECT 
    ld.[ld_site],
    xz.[xxwezoned_area_id],
    xz.[xxwezoned_zone_id],
    ld.[ld_loc],
    ld.[ld_part],
    ld.[ld_qty_oh],
    ld.[ld_status],
    sc.[sct_cst_tot],
    (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [mat_cost],
    (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
    (ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],  -- Calculate COGS
    (ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT],  -- Calculate CMAT
    pt.[pt_desc1], 
    pt.[pt_desc2], 
    pt.[pt_prod_line], 
    pt.[pt_group], 
    pt.[pt_part_type], 
    pt.[pt_status], 
    pt.[pt_abc], 
    pt.[pt_cyc_int], 
    pt.[pt_sfty_stk], 
    pt.[pt_sfty_time], 
    pt.[pt_buyer], 
    pt.[pt_vend], 
    pt.[pt_routing], 
    pt.[pt_net_wt], 
    pt.[pt_net_wt_um], 
    pt.[pt__chr02], 
    pt.[pt_dsgn_grp],
    ISNULL(b.[Parent], 'No') AS [Parent],  -- Add Parent column
    ISNULL(b.[Child], 'No') AS [Child],   -- Add Child column
    ISNULL(b.[SFG], 'No') AS [SFG]        -- Add SFG column
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
        UNION ALL
        SELECT
            [sct_site],
            [sct_part],
            [sct_cst_tot],
            [sct_mtl_tl],
            [sct_mtl_ll]
        FROM 
            [QADEE].[dbo].[sct_det]
        WHERE 
            [sct_sim] = 'standard'
    ) sc
ON 
    ld.[ld_part] = sc.[sct_part] 
    AND ld.[ld_site] = sc.[sct_site]
JOIN 
    (
        SELECT 
            [pt_site], 
            [pt_part], 
            [pt_desc1], 
            [pt_desc2], 
            [pt_prod_line], 
            [pt_group], 
            [pt_part_type], 
            [pt_status], 
            [pt_abc], 
            [pt_cyc_int], 
            [pt_sfty_stk], 
            [pt_sfty_time], 
            [pt_buyer], 
            [pt_vend], 
            [pt_routing], 
            [pt_net_wt], 
            [pt_net_wt_um], 
            [pt__chr02], 
            [pt_dsgn_grp]
        FROM [QADEE2798].[dbo].[pt_mstr]
        WHERE [pt_part_type] NOT IN ('xc', 'rc')  -- Filter out 'xc' and 'rc'

        UNION ALL

        SELECT 
            [pt_site], 
            [pt_part], 
            [pt_desc1], 
            [pt_desc2], 
            [pt_prod_line], 
            [pt_group], 
            [pt_part_type], 
            [pt_status], 
            [pt_abc], 
            [pt_cyc_int], 
            [pt_sfty_stk], 
            [pt_sfty_time], 
            [pt_buyer], 
            [pt_vend], 
            [pt_routing], 
            [pt_net_wt], 
            [pt_net_wt_um], 
            [pt__chr02], 
            [pt_dsgn_grp]
        FROM [QADEE].[dbo].[pt_mstr]
        WHERE [pt_part_type] NOT IN ('xc', 'rc')  -- Filter out 'xc' and 'rc'
    ) pt
ON 
    ld.[ld_site] = pt.[pt_site] 
    AND ld.[ld_part] = pt.[pt_part]
LEFT JOIN 
    BOMStatusCTE b
ON 
    ld.[ld_site] = b.[Plant] 
    AND ld.[ld_part] = b.[Item Number]

UNION ALL

SELECT 
    ld.[ld_site],
    xz.[xxwezoned_area_id],
    xz.[xxwezoned_zone_id],
    ld.[ld_loc],
    ld.[ld_part],
    ld.[ld_qty_oh],
    ld.[ld_status],
    sc.[sct_cst_tot],
    (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [mat_cost],
    (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
    (ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],  -- Calculate COGS
    (ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT],  -- Calculate CMAT
    pt.[pt_desc1], 
    pt.[pt_desc2], 
    pt.[pt_prod_line], 
    pt.[pt_group], 
    pt.[pt_part_type], 
    pt.[pt_status], 
    pt.[pt_abc], 
    pt.[pt_cyc_int], 
    pt.[pt_sfty_stk], 
    pt.[pt_sfty_time], 
    pt.[pt_buyer], 
    pt.[pt_vend], 
    pt.[pt_routing], 
    pt.[pt_net_wt], 
    pt.[pt_net_wt_um], 
    pt.[pt__chr02], 
    pt.[pt_dsgn_grp],
    ISNULL(b.[Parent], 'No') AS [Parent],  -- Add Parent column
    ISNULL(b.[Child], 'No') AS [Child],   -- Add Child column
    ISNULL(b.[SFG], 'No') AS [SFG]        -- Add SFG column
FROM 
    [QADEE].[dbo].[ld_det] ld
JOIN 
    [QADEE].[dbo].[xxwezoned_det] xz
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
            [QADEE].[dbo].[sct_det]
        WHERE 
            [sct_sim] = 'standard'
        UNION ALL
        SELECT
            [sct_site],
            [sct_part],
            [sct_cst_tot],
            [sct_mtl_tl],
            [sct_mtl_ll]
        FROM 
            [QADEE].[dbo].[sct_det]
        WHERE 
            [sct_sim] = 'standard'
    ) sc
ON 
    ld.[ld_part] = sc.[sct_part] 
    AND ld.[ld_site] = sc.[sct_site]
JOIN 
    (
        SELECT 
            [pt_site], 
            [pt_part], 
            [pt_desc1], 
            [pt_desc2], 
            [pt_prod_line], 
            [pt_group], 
            [pt_part_type], 
            [pt_status], 
            [pt_abc], 
            [pt_cyc_int], 
            [pt_sfty_stk], 
            [pt_sfty_time], 
            [pt_buyer], 
            [pt_vend], 
            [pt_routing], 
            [pt_net_wt], 
            [pt_net_wt_um], 
            [pt__chr02], 
            [pt_dsgn_grp]
        FROM [QADEE2798].[dbo].[pt_mstr]
        WHERE [pt_part_type] NOT IN ('xc', 'rc')  -- Filter out 'xc' and 'rc'

        UNION ALL

        SELECT 
            [pt_site], 
            [pt_part], 
            [pt_desc1], 
            [pt_desc2], 
            [pt_prod_line], 
            [pt_group], 
            [pt_part_type], 
            [pt_status], 
            [pt_abc], 
            [pt_cyc_int], 
            [pt_sfty_stk], 
            [pt_sfty_time], 
            [pt_buyer], 
            [pt_vend], 
            [pt_routing], 
            [pt_net_wt], 
            [pt_net_wt_um], 
            [pt__chr02], 
            [pt_dsgn_grp]
        FROM [QADEE].[dbo].[pt_mstr]
        WHERE [pt_part_type] NOT IN ('xc', 'rc')  -- Filter out 'xc' and 'rc'
    ) pt
ON 
    ld.[ld_site] = pt.[pt_site] 
    AND ld.[ld_part] = pt.[pt_part]
LEFT JOIN 
    BOMStatusCTE b
ON 
    ld.[ld_site] = b.[Plant] 
    AND ld.[ld_part] = b.[Item Number];