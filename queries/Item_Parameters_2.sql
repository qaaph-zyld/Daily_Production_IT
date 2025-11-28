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
),
PsOpCTE AS (
    SELECT DISTINCT
        '2674' AS [Plant],
        [ps_par] AS [Item Number],
        [ps_op] AS [Operation]
    FROM
        [QADEE].[dbo].[ps_mstr]
    WHERE
        [ps_end] IS NULL

    UNION ALL

    SELECT DISTINCT
        '2798' AS [Plant],
        [ps_par] AS [Item Number],
        [ps_op] AS [Operation]
    FROM
        [QADEE2798].[dbo].[ps_mstr]
    WHERE
        [ps_end] IS NULL
),
InventoryByArea AS (
    SELECT
        ld.[ld_site] AS [Plant],
        ld.[ld_part] AS [Item Number],
        SUM(CASE WHEN xz.[xxwezoned_area_id] = 'WH' THEN ld.[ld_qty_oh] ELSE 0 END) AS [WH_Qty],
        SUM(CASE WHEN xz.[xxwezoned_area_id] = 'WIP' THEN ld.[ld_qty_oh] ELSE 0 END) AS [WIP_Qty],
        SUM(CASE WHEN xz.[xxwezoned_area_id] = 'EXLPICK' THEN ld.[ld_qty_oh] ELSE 0 END) AS [EXLPICK_Qty]
    FROM 
        [QADEE2798].[dbo].[ld_det] ld
    JOIN 
        [QADEE2798].[dbo].[xxwezoned_det] xz
    ON 
        ld.[ld_loc] = xz.[xxwezoned_loc]
    GROUP BY
        ld.[ld_site],
        ld.[ld_part]
)
SELECT DISTINCT
    ld.[ld_site] AS [Plant],
    ld.[ld_part] AS [Item Number],
    sc.[sct_cst_tot] AS [Standard Cost],
    (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [Material Cost],
    (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
    pt.[pt_desc1] AS [Item Description], 
    pt.[pt_desc2] AS [pt_desc2], 
    pt.[pt_prod_line] AS [Prod Line], 
    pt.[pt_group] AS [Group], 
    pt.[pt_part_type] AS [pt_part_type], 
    pt.[pt_status] AS [Item Number Status],
    pt.[pt_added] AS [Date added],
    pt.[pt_abc] AS [ABC], 
    pt.[pt_cyc_int] AS [pt_cyc_int], 
    pt.[pt_sfty_stk] AS [Safety Stock], 
    pt.[pt_sfty_time] AS [Safety Time], 
    pt.[pt_buyer] AS [Item Planner], 
    pt.[pt_vend] AS [Item Supplier], 
    pt.[pt_routing] AS [Routing], 
    pt.[pt_net_wt] AS [Net weight], 
    pt.[pt_net_wt_um] AS [pt_net_wt_um], 
    pt.[pt__chr02] AS [Item Type], 
    pt.[pt_dsgn_grp] AS [Project],
    CASE
        WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
        WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
        WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
        ELSE 'No BOM'
    END AS [FG/SFG/RM],
    psop.[Operation] AS [Operation],
    inv.[in_qty_oh] AS [Total Qty Nettable],
    inv.[in_qty_nonet] AS [Total Qty Nonet],
    inv.[total_inv] AS [Total Inv],
    inv.[Last_ISSUE] AS [Last_ISSUE],
    inv.[Last_REC] AS [Last_REC],
    inv.[Last_CC] AS [Last_CC],
    CASE
        WHEN inv.[in_iss_date] IS NULL THEN 'No transactions'
        WHEN inv.[12 months] = 'yes' THEN '12 months'
        WHEN inv.[6 months] = 'yes' THEN '6 months'
        WHEN inv.[3 months] = 'yes' THEN '3 months'
        ELSE 'Active'
    END AS [Obsolete],
    (inv.[total_inv] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [Total CMAT],
    (inv.[total_inv] * sc.[sct_cst_tot]) AS [Total COGS],
    '' AS [New],
    ia.[WH_Qty],
    ia.[WIP_Qty],
    ia.[EXLPICK_Qty],
    ia.[WH_Qty] * sc.[sct_cst_tot] AS [WH_Value],
    ia.[WIP_Qty] * sc.[sct_cst_tot] AS [WIP_Value],
    ia.[EXLPICK_Qty] * sc.[sct_cst_tot] AS [EXLPICK_Value]
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
        UNION
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
            [pt_added],
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
        UNION
        SELECT 
            [pt_site], 
            [pt_part], 
            [pt_desc1], 
            [pt_desc2], 
            [pt_prod_line], 
            [pt_group], 
            [pt_part_type], 
            [pt_status], 
            [pt_added],
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
    ) pt
ON 
    ld.[ld_part] = pt.[pt_part] 
    AND ld.[ld_site] = pt.[pt_site]
LEFT JOIN 
    BOMStatusCTE b
ON 
    ld.[ld_site] = b.[Plant] 
    AND ld.[ld_part] = b.[Item Number]
LEFT JOIN 
    PsOpCTE psop
ON
    ld.[ld_site] = psop.[Plant]
    AND ld.[ld_part] = psop.[Item Number]
LEFT JOIN 
    (
        SELECT 
            [in_site],
            [in_part],
            [in_qty_oh],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            [in_iss_date],
            DATEDIFF(DAY, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(DAY, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(DAY, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 12 THEN 'yes' 
                ELSE 'no' 
            END AS [12 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 6 THEN 'yes' 
                ELSE 'no' 
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 3 THEN 'yes' 
                ELSE 'no' 
            END AS [3 months]
        FROM 
            [QADEE].[dbo].[in_mstr]
        UNION
        SELECT 
            [in_site],
            [in_part],
            [in_qty_oh],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            [in_iss_date],
            DATEDIFF(DAY, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(DAY, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(DAY, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 12 THEN 'yes' 
                ELSE 'no' 
            END AS [12 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 6 THEN 'yes' 
                ELSE 'no' 
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 3 THEN 'yes' 
                ELSE 'no' 
            END AS [3 months]
        FROM 
            [QADEE2798].[dbo].[15]
    ) inv
ON 
    ld.[ld_part] = inv.[in_part] 
    AND ld.[ld_site] = inv.[in_site]
LEFT JOIN
    InventoryByArea ia
ON
    ld.[ld_site] = ia.[Plant]
    AND ld.[ld_part] = ia.[Item Number]

UNION ALL

SELECT DISTINCT
    ld.[ld_site] AS [Plant],
    ld.[ld_part] AS [Item Number],
    sc.[sct_cst_tot] AS [Standard Cost],
    (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [Material Cost],
    (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
    pt.[pt_desc1] AS [Item Description], 
    pt.[pt_desc2] AS [pt_desc2], 
    pt.[pt_prod_line] AS [Prod Line], 
    pt.[pt_group] AS [Group], 
    pt.[pt_part_type] AS [pt_part_type], 
    pt.[pt_status] AS [Item Number Status],
    pt.[pt_added] AS [Date added],
    pt.[pt_abc] AS [ABC], 
    pt.[pt_cyc_int] AS [pt_cyc_int], 
    pt.[pt_sfty_stk] AS [Safety Stock], 
    pt.[pt_sfty_time] AS [Safety Time], 
    pt.[pt_buyer] AS [Item Planner], 
    pt.[pt_vend] AS [Item Supplier], 
    pt.[pt_routing] AS [Routing], 
    pt.[pt_net_wt] AS [Net weight], 
    pt.[pt_net_wt_um] AS [pt_net_wt_um], 
    pt.[pt__chr02] AS [Item Type], 
    pt.[pt_dsgn_grp] AS [Project],
    CASE
        WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
        WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
        WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
        ELSE 'No BOM'
    END AS [FG/SFG/RM],
    psop.[Operation] AS [Operation],
    inv.[in_qty_oh] AS [Total Qty Nettable],
    inv.[in_qty_nonet] AS [Total Qty Nonet],
    inv.[total_inv] AS [Total Inv],
    inv.[Last_ISSUE] AS [Last_ISSUE],
    inv.[Last_REC] AS [Last_REC],
    inv.[Last_CC] AS [Last_CC],
    CASE
        WHEN inv.[in_iss_date] IS NULL THEN 'No transactions'
        WHEN inv.[12 months] = 'yes' THEN '12 months'
        WHEN inv.[6 months] = 'yes' THEN '6 months'
        WHEN inv.[3 months] = 'yes' THEN '3 months'
        ELSE 'Active'
    END AS [Obsolete],
    (inv.[total_inv] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [Total CMAT],
    (inv.[total_inv] * sc.[sct_cst_tot]) AS [Total COGS],
    '' AS [New],
    ia.[WH_Qty],
    ia.[WIP_Qty],
    ia.[EXLPICK_Qty],
    ia.[WH_Qty] * sc.[sct_cst_tot] AS [WH_Value],
    ia.[WIP_Qty] * sc.[sct_cst_tot] AS [WIP_Value],
    ia.[EXLPICK_Qty] * sc.[sct_cst_tot] AS [EXLPICK_Value]
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
        UNION
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
            [pt_added],
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
        UNION
        SELECT 
            [pt_site], 
            [pt_part], 
            [pt_desc1], 
            [pt_desc2], 
            [pt_prod_line], 
            [pt_group], 
            [pt_part_type], 
            [pt_status], 
            [pt_added],
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
    ) pt
ON 
    ld.[ld_part] = pt.[pt_part] 
    AND ld.[ld_site] = pt.[pt_site]
LEFT JOIN 
    BOMStatusCTE b
ON 
    ld.[ld_site] = b.[Plant] 
    AND ld.[ld_part] = b.[Item Number]
LEFT JOIN 
    PsOpCTE psop
ON
    ld.[ld_site] = psop.[Plant]
    AND ld.[ld_part] = psop.[Item Number]
LEFT JOIN 
    (
        SELECT 
            [in_site],
            [in_part],
            [in_qty_oh],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            [in_iss_date],
            DATEDIFF(DAY, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(DAY, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(DAY, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 12 THEN 'yes' 
                ELSE 'no' 
            END AS [12 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 6 THEN 'yes' 
                ELSE 'no' 
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 3 THEN 'yes' 
                ELSE 'no' 
            END AS [3 months]
        FROM 
            [QADEE].[dbo].[in_mstr]
        UNION
        SELECT 
            [in_site],
            [in_part],
            [in_qty_oh],
            [in_qty_nonet],
            [in_qty_oh] + [in_qty_nonet] AS [total_inv],
            [in_iss_date],
            DATEDIFF(DAY, [in_iss_date], GETDATE()) AS [Last_ISSUE],
            DATEDIFF(DAY, [in_rec_date], GETDATE()) AS [Last_REC],
            DATEDIFF(DAY, [in_cnt_date], GETDATE()) AS [Last_CC],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 12 THEN 'yes' 
                ELSE 'no' 
            END AS [12 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 6 THEN 'yes' 
                ELSE 'no' 
            END AS [6 months],
            CASE 
                WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 3 THEN 'yes' 
                ELSE 'no' 
            END AS [3 months]
        FROM 
            [QADEE2798].[dbo].[15]
    ) inv
ON 
    ld.[ld_part] = inv.[in_part] 
    AND ld.[ld_site] = inv.[in_site]
LEFT JOIN
    InventoryByArea ia
ON
    ld.[ld_site] = ia.[Plant]
    AND ld.[ld_part] = ia.[Item Number]
