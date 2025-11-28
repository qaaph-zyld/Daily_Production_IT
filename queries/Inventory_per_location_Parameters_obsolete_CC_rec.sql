WITH ParentCTE AS (
    SELECT DISTINCT  
        '2674' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]
    FROM   
        [QADEE].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL

    UNION ALL  

    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]
    FROM   
        [QADEE2798].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL
),
ChildCTE AS (
    SELECT DISTINCT  
        '2674' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]
    FROM   
        [QADEE].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  

    UNION ALL  

    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]
    FROM   
        [QADEE2798].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  
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
        END AS [SFG]
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
    (ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],
    (ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT],
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
    ISNULL(b.[Parent], 'No') AS [Parent],
    ISNULL(b.[Child], 'No') AS [Child],
    ISNULL(b.[SFG], 'No') AS [SFG],
    CASE 
        WHEN ISNULL(b.[Parent], 'No') = 'No' 
             AND ISNULL(b.[Child], 'No') = 'No' 
             AND ISNULL(b.[SFG], 'No') = 'No' 
        THEN 'No' 
        ELSE 'Yes' 
    END AS [BOM],
    -- Add inventory columns from the second query
    inv.[in_qty_oh],
    inv.[in_iss_date],
    inv.[in_rec_date],
    inv.[in_cnt_date],
    inv.[in_qty_nonet],
    inv.[total_inv],
    inv.[Last_ISSUE],
    inv.[Last_REC],
    inv.[Last_CC],
    inv.[3 months],
    inv.[6 months],
    inv.[12 months]
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
        WHERE [pt_part_type] NOT IN ('xc', 'rc')

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
        WHERE [pt_part_type] NOT IN ('xc', 'rc')
    ) pt
ON 
    ld.[ld_site] = pt.[pt_site] 
    AND ld.[ld_part] = pt.[pt_part]
LEFT JOIN 
    BOMStatusCTE b
ON 
    ld.[ld_site] = b.[Plant] 
    AND ld.[ld_part] = b.[Item Number]
LEFT JOIN 
    (
        SELECT 
            [in_part],
            [in_site],
            [in_qty_oh],
            [in_iss_date],
            [in_rec_date],
            [in_cnt_date],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            DATEDIFF(day, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(day, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(day, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 91 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 180 THEN 'yes'
                ELSE 'no'
            END AS [3 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 180 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 365 THEN 'yes'
                ELSE 'no'
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 365 THEN 'yes'
                ELSE 'no'
            END AS [12 months]
        FROM 
            [QADEE].[dbo].[in_mstr]
        WHERE 
            [in_qty_oh] + [in_qty_nonet] <> 0

        UNION ALL

        SELECT 
            [in_part],
            [in_site],
            [in_qty_oh],
            [in_iss_date],
            [in_rec_date],
            [in_cnt_date],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            DATEDIFF(day, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(day, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(day, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 91 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 180 THEN 'yes'
                ELSE 'no'
            END AS [3 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 180 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 365 THEN 'yes'
                ELSE 'no'
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 365 THEN 'yes'
                ELSE 'no'
            END AS [12 months]
        FROM 
            [QADEE2798].[dbo].[15]
        WHERE 
            [in_qty_oh] + [in_qty_nonet] <> 0
    ) inv
ON 
    ld.[ld_part] = inv.[in_part] 
    AND ld.[ld_site] = inv.[in_site]

UNION ALL

-- Repeat the entire query for the second database (QADEE)
-- (Identical to the above query, just replacing QADEE2798 with QADEE)
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
    (ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],
    (ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT],
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
    ISNULL(b.[Parent], 'No') AS [Parent],
    ISNULL(b.[Child], 'No') AS [Child],
    ISNULL(b.[SFG], 'No') AS [SFG],
    CASE 
        WHEN ISNULL(b.[Parent], 'No') = 'No' 
             AND ISNULL(b.[Child], 'No') = 'No' 
             AND ISNULL(b.[SFG], 'No') = 'No' 
        THEN 'No' 
        ELSE 'Yes' 
    END AS [BOM],
    -- Add inventory columns from the second query
    inv.[in_qty_oh],
    inv.[in_iss_date],
    inv.[in_rec_date],
    inv.[in_cnt_date],
    inv.[in_qty_nonet],
    inv.[total_inv],
   -- Continuing the previous query
    inv.[Last_ISSUE],
    inv.[Last_REC],
    inv.[Last_CC],
    inv.[3 months],
    inv.[6 months],
    inv.[12 months]
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
        WHERE [pt_part_type] NOT IN ('xc', 'rc')

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
        WHERE [pt_part_type] NOT IN ('xc', 'rc')
    ) pt
ON 
    ld.[ld_site] = pt.[pt_site] 
    AND ld.[ld_part] = pt.[pt_part]
LEFT JOIN 
    BOMStatusCTE b
ON 
    ld.[ld_site] = b.[Plant] 
    AND ld.[ld_part] = b.[Item Number]
LEFT JOIN 
    (
        SELECT 
            [in_part],
            [in_site],
            [in_qty_oh],
            [in_iss_date],
            [in_rec_date],
            [in_cnt_date],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            DATEDIFF(day, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(day, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(day, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 91 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 180 THEN 'yes'
                ELSE 'no'
            END AS [3 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 180 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 365 THEN 'yes'
                ELSE 'no'
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 365 THEN 'yes'
                ELSE 'no'
            END AS [12 months]
        FROM 
            [QADEE].[dbo].[in_mstr]
        WHERE 
            [in_qty_oh] + [in_qty_nonet] <> 0

        UNION ALL

        SELECT 
            [in_part],
            [in_site],
            [in_qty_oh],
            [in_iss_date],
            [in_rec_date],
            [in_cnt_date],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            DATEDIFF(day, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(day, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(day, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 91 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 180 THEN 'yes'
                ELSE 'no'
            END AS [3 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 180 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 365 THEN 'yes'
                ELSE 'no'
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 365 THEN 'yes'
                ELSE 'no'
            END AS [12 months]
        FROM 
            [QADEE2798].[dbo].[15]
        WHERE 
            [in_qty_oh] + [in_qty_nonet] <> 0
    ) inv
ON 
    ld.[ld_part] = inv.[in_part] 
    AND ld.[ld_site] = inv.[in_site]
ORDER BY 
    ld.[ld_part];  -- Optional: order the results by part number