WITH
  -- First Table: Data from ps_mstr (5 columns)
  PS_Data AS (
    SELECT 
      '2674' AS [Plant],
      [ps_par],  
      [ps_comp],  
      [ps_qty_per],  
      [ps_rmks]
    FROM 
      [QADEE].[dbo].[ps_mstr]
    WHERE 
      [ps_end] IS NULL
    UNION ALL
    SELECT 
      '2798' AS [Plant],
      [ps_par],  
      [ps_comp],  
      [ps_qty_per],  
      [ps_rmks]
    FROM 
      [QADEE2798].[dbo].[ps_mstr]
    WHERE 
      [ps_end] IS NULL
  ),

  -- Second Table: Data from PodData + pt_mstr (25 columns)
  PodData AS (
    SELECT  
      pd.[pod_po_site],  
      pd.[pod__chr08],  
      pm.[po_vend],  
      pd.[pod_nbr],  
      pd.[pod_line],  
      pd.[pod_part],  
      pd.[pod_cum_qty[1]]],  
      pd.[pod_ord_mult],  
      pd.[pod_translt_days],
      pd.[pod_start_eff[1]]],
      pd.[pod_curr_rlse_id[1]]],
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
      pt.[pt__chr02],
      pt.[pt_dsgn_grp]
    FROM  
      [QADEE].[dbo].[pod_det] pd  
      JOIN [QADEE].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
      LEFT JOIN [QADEE].[dbo].[pt_mstr] pt 
        ON pd.[pod_po_site] = pt.[pt_site] 
        AND pd.[pod_part] = pt.[pt_part]
    WHERE  
      pd.[pod_end_eff[1]]] = '2049-12-31 00:00:00'  
    UNION ALL  
    SELECT  
      pd.[pod_po_site],  
      pd.[pod__chr08],  
      pm.[po_vend],  
      pd.[pod_nbr],  
      pd.[pod_line],  
      pd.[pod_part],  
      pd.[pod_cum_qty[1]]],  
      pd.[pod_ord_mult],  
      pd.[pod_translt_days],
      pd.[pod_start_eff[1]]],
      pd.[pod_curr_rlse_id[1]]],
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
      pt.[pt__chr02],
      pt.[pt_dsgn_grp]
    FROM  
      [QADEE2798].[dbo].[pod_det] pd  
      JOIN [QADEE2798].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
      LEFT JOIN [QADEE2798].[dbo].[pt_mstr] pt 
        ON pd.[pod_po_site] = pt.[pt_site] 
        AND pd.[pod_part] = pt.[pt_part]
    WHERE  
      pd.[pod_end_eff[1]]] = '2049-12-31 00:00:00'
  ),

  -- BOM Status CTE
  ParentCTE AS (
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

  -- Additional Data for sct_cst_tot, mat_cost, LBO, COGS, CMAT, etc.
  AdditionalData AS (
    SELECT 
        ld.[ld_site],
        ld.[ld_part],
        sc.[sct_cst_tot],
        (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [mat_cost],
        (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
        SUM(ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],
        SUM(ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT],
        pt.[pt_net_wt],
        pt.[pt_net_wt_um],
        ISNULL(b.[Parent], 'No') AS [Parent],  
        ISNULL(b.[Child], 'No') AS [Child],   
        ISNULL(b.[SFG], 'No') AS [SFG],        
        SUM(CASE WHEN xz.[xxwezoned_zone_id] = 'WH' THEN ld.[ld_qty_oh] ELSE 0 END) AS [QTY_WH],  
        SUM(CASE WHEN xz.[xxwezoned_zone_id] = 'EXLPICK' THEN ld.[ld_qty_oh] ELSE 0 END) AS [QTY_EXLPICK],  
        SUM(CASE WHEN xz.[xxwezoned_zone_id] = 'WIP' THEN ld.[ld_qty_oh] ELSE 0 END) AS [QTY_WIP]  
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
                [pt_net_wt], 
                [pt_net_wt_um]
            FROM [QADEE2798].[dbo].[pt_mstr]
            WHERE [pt_part_type] NOT IN ('xc', 'rc')  
            UNION ALL
            SELECT 
                [pt_site], 
                [pt_part], 
                [pt_net_wt], 
                [pt_net_wt_um]
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
    GROUP BY 
        ld.[ld_site],
        ld.[ld_part],
        sc.[sct_cst_tot],
        sc.[sct_mtl_tl],
        sc.[sct_mtl_ll],
        pt.[pt_net_wt],
        pt.[pt_net_wt_um],
        b.[Parent],
        b.[Child],
        b.[SFG]

    UNION ALL

    -- Repeat the same logic for the second database (QADEE)
    SELECT 
        ld.[ld_site],
        ld.[ld_part],
        sc.[sct_cst_tot],
        (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [mat_cost],
        (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
        SUM(ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],
        SUM(ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT],
        pt.[pt_net_wt],
        pt.[pt_net_wt_um],
        ISNULL(b.[Parent], 'No') AS [Parent],  
        ISNULL(b.[Child], 'No') AS [Child],   
        ISNULL(b.[SFG], 'No') AS [SFG],        
        SUM(CASE WHEN xz.[xxwezoned_zone_id] = 'WH' THEN ld.[ld_qty_oh] ELSE 0 END) AS [QTY_WH],  
        SUM(CASE WHEN xz.[xxwezoned_zone_id] = 'EXLPICK' THEN ld.[ld_qty_oh] ELSE 0 END) AS [QTY_EXLPICK],  
        SUM(CASE WHEN xz.[xxwezoned_zone_id] = 'WIP' THEN ld.[ld_qty_oh] ELSE 0 END) AS [QTY_WIP]  
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
                [pt_net_wt], 
                [pt_net_wt_um]
            FROM [QADEE2798].[dbo].[pt_mstr]
            WHERE [pt_part_type] NOT IN ('xc', 'rc')  
            UNION ALL
            SELECT 
                [pt_site], 
                [pt_part], 
                [pt_net_wt], 
                [pt_net_wt_um]
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
    GROUP BY 
        ld.[ld_site],
        ld.[ld_part],
        sc.[sct_cst_tot],
        sc.[sct_mtl_tl],
        sc.[sct_mtl_ll],
        pt.[pt_net_wt],
        pt.[pt_net_wt_um],
        b.[Parent],
        b.[Child],
        b.[SFG]
  )

-- Final Merged Result (30 Columns)
SELECT
  -- First 5 columns from PS_Data
  ps.[Plant],
  ps.[ps_par],
  ps.[ps_comp],
  ps.[ps_qty_per],
  ps.[ps_rmks],

  -- Next 25 columns from PodData
  pd.[pod__chr08],
  pd.[po_vend],
  pd.[pod_nbr],
  pd.[pod_line],
  pd.[pod_cum_qty[1]]],
  pd.[pod_ord_mult],
  pd.[pod_translt_days],
  pd.[pod_start_eff[1]]],
  pd.[pod_curr_rlse_id[1]]],
  pd.[pt_desc1],
  pd.[pt_desc2],
  pd.[pt_prod_line],
  pd.[pt_group],
  pd.[pt_status],
  pd.[pt_sfty_stk],
  pd.[pt_sfty_time],
  pd.[pt_buyer],
  pd.[pt_vend],
  pd.[pt__chr02],
  pd.[pt_dsgn_grp],

  -- Additional columns from AdditionalData
  ad.[sct_cst_tot],
  ad.[mat_cost],
  ad.[LBO],
  ad.[COGS],
  ad.[CMAT],
  ad.[pt_net_wt],
  ad.[pt_net_wt_um],
  ad.[Parent],
  ad.[Child],
  ad.[SFG],
  ad.[QTY_WH],
  ad.[QTY_EXLPICK],
  ad.[QTY_WIP]
FROM
  PS_Data ps
  LEFT JOIN PodData pd 
    ON ps.[Plant] = pd.[pod_po_site] 
    AND ps.[ps_comp] = pd.[pod_part]
  LEFT JOIN AdditionalData ad 
    ON ps.[Plant] = ad.[ld_site] 
    AND ps.[ps_comp] = ad.[ld_part];