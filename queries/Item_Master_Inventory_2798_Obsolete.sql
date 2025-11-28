WITH CombinedParts AS (
  SELECT 
    [pt_part],
    [pt_site],
    [pt_added]
  FROM [QADEE2798].[dbo].[pt_mstr]
  WHERE [pt_part_type] NOT IN ('xc','rc')
),
ParentCTE AS (
    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]
    FROM [QADEE2798].[dbo].[ps_mstr]  
    WHERE [ps_end] IS NULL
),
ChildCTE AS (
    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]
    FROM [QADEE2798].[dbo].[ps_mstr]  
    WHERE [ps_end] IS NULL  
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
),
ItemCosts AS (
    SELECT
        [sct_site],
        [sct_part],
        [sct_cst_tot] AS [Standard Cost],
        ([sct_mtl_tl] + [sct_mtl_ll]) AS [Material Cost],
        ([sct_cst_tot] - ([sct_mtl_tl] + [sct_mtl_ll])) AS [LBO]
    FROM 
        [QADEE2798].[dbo].[sct_det]
    WHERE 
        [sct_sim] = 'standard'
),
ItemDetails AS (
    SELECT 
        [pt_site], 
        [pt_part], 
        [pt_desc1] AS [Item Description], 
        [pt_desc2] AS [pt_desc2], 
        [pt_prod_line] AS [Prod Line], 
        [pt_group] AS [Group], 
        [pt_part_type] AS [pt_part_type], 
        [pt_status] AS [Item Number Status],
        [pt_added] AS [Date added],
        [pt_abc] AS [ABC], 
        [pt_cyc_int] AS [pt_cyc_int], 
        [pt_sfty_stk] AS [Safety Stock], 
        [pt_sfty_time] AS [Safety Time], 
        [pt_buyer] AS [Item Planner], 
        [pt_vend] AS [Item Supplier], 
        [pt_routing] AS [Routing], 
        [pt_net_wt] AS [Net weight], 
        [pt_net_wt_um] AS [pt_net_wt_um], 
        [pt__chr02] AS [Item Type], 
        [pt_dsgn_grp] AS [Project]
    FROM [QADEE2798].[dbo].[pt_mstr]
),
InventoryInfo AS (
    SELECT 
        [in_site],
        [in_part],
        [in_qty_oh] AS [Total Qty Nettable],
        [in_qty_nonet] AS [Total Qty Nonet],
        [in_qty_oh] + [in_qty_nonet] AS [Total Inv],
        [in_iss_date],
        CASE WHEN DATEDIFF(DAY, [in_iss_date], GETDATE()) < 0 THEN 0 ELSE DATEDIFF(DAY, [in_iss_date], GETDATE()) END AS [Last_ISSUE],
        DATEDIFF(DAY, [in_rec_date], GETDATE()) AS [Last_REC],
        DATEDIFF(DAY, [in_cnt_date], GETDATE()) AS [Last_CC],
        CASE 
            WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 12 THEN 'yes' 
            ELSE 'no' 
        END AS [12_months],
        CASE 
            WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 6 THEN 'yes' 
            ELSE 'no' 
        END AS [6_months],
        CASE 
            WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 3 THEN 'yes' 
            ELSE 'no' 
        END AS [3_months],
        CASE
            WHEN [in_iss_date] IS NULL THEN 'No transactions'
            WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 12 THEN '12 months'
            WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 6 THEN '6 months'
            WHEN DATEDIFF(MONTH, [in_iss_date], GETDATE()) > 3 THEN '3 months'
            ELSE 'Active'
        END AS [Obsolete]
    FROM 
        [QADEE2798].[dbo].[15]
),
DistinctPartSites AS (
    SELECT DISTINCT tr_part, tr_site
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE tr_type = 'iss-wo'
      AND tr_effdate >= DATEADD(WEEK, -4, GETDATE())
),
DailyISS_WO AS (
    SELECT
        tr_part,
        tr_site,
        tr_effdate,
        SUM(ABS(tr_qty_loc)) AS Daily_Qty
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE tr_type = 'iss-wo'
      AND tr_effdate >= DATEADD(WEEK, -4, GETDATE())
    GROUP BY
        tr_part,
        tr_site,
        tr_effdate
),
WeeklyTotals AS (
    SELECT
        tr_part,
        tr_site,
        DATEPART(YEAR, tr_effdate) AS Year,
        DATEPART(WEEK, tr_effdate) AS WeekNum,
        SUM(Daily_Qty) AS WeeklyTotal
    FROM DailyISS_WO
    GROUP BY
        tr_part,
        tr_site,
        DATEPART(YEAR, tr_effdate),
        DATEPART(WEEK, tr_effdate)
),
WeeklyPivot AS (
    SELECT 
        tr_part,
        tr_site,
        ISNULL([Week1], 0) AS [total_ISS-WO_CW_-1],
        ISNULL([Week2], 0) AS [total_ISS-WO_CW_-2],
        ISNULL([Week3], 0) AS [total_ISS-WO_CW_-3],
        ISNULL([Week4], 0) AS [total_ISS-WO_CW_-4]
    FROM (
        SELECT 
            wt.tr_part,
            wt.tr_site,
            'Week' + CAST(DENSE_RANK() OVER (
                PARTITION BY wt.tr_part, wt.tr_site 
                ORDER BY wt.Year DESC, wt.WeekNum DESC
            ) AS VARCHAR) AS WeekLabel,
            wt.WeeklyTotal
        FROM WeeklyTotals wt
    ) AS Source
    PIVOT (
        SUM(WeeklyTotal)
        FOR WeekLabel IN ([Week1], [Week2], [Week3], [Week4])
    ) AS PivotTable
),
WIPMinimumData AS (
    SELECT
        wp.tr_part,
        wp.tr_site,
        wp.[total_ISS-WO_CW_-1],
        wp.[total_ISS-WO_CW_-2],
        wp.[total_ISS-WO_CW_-3],
        wp.[total_ISS-WO_CW_-4],
        CASE 
            WHEN (wp.[total_ISS-WO_CW_-1] + wp.[total_ISS-WO_CW_-2] + 
                 wp.[total_ISS-WO_CW_-3] + wp.[total_ISS-WO_CW_-4]) / 4.0 * 3 < 0 
            THEN 0 
            ELSE CAST(ROUND((wp.[total_ISS-WO_CW_-1] + wp.[total_ISS-WO_CW_-2] + 
                 wp.[total_ISS-WO_CW_-3] + wp.[total_ISS-WO_CW_-4]) / 4.0 * 3, 0) AS INT) 
        END AS [recent_wip_min],
        CASE 
            WHEN (wp.[total_ISS-WO_CW_-1] + wp.[total_ISS-WO_CW_-2] + 
                 wp.[total_ISS-WO_CW_-3] + wp.[total_ISS-WO_CW_-4]) / 4.0 * 7 < 0 
            THEN 0 
            ELSE CAST(ROUND((wp.[total_ISS-WO_CW_-1] + wp.[total_ISS-WO_CW_-2] + 
                 wp.[total_ISS-WO_CW_-3] + wp.[total_ISS-WO_CW_-4]) / 4.0 * 7, 0) AS INT) 
        END AS [recent_wip_max]
    FROM WeeklyPivot wp
),
SOD_Data AS (
    SELECT 
        [sod_order_category],
        [sod_loc],
        [sod_nbr],
        [sod_line],
        [sod_part],
        [sod_qty_ship],
        [sod_site],
        [sod_cum_qty[1]]],
        [sod_cum_qty[2]]],
        [sod_cum_date[2]]],
        [sod_curr_rlse_id[1]]],
        [sod_end_eff[1]]],
        [sod__qadd04],
        [sod_unadjust_cum_qty],
        [sod_status]
    FROM [QADEE2798].[dbo].[sod_det]
),
SO_Data AS (
    SELECT 
        [so_nbr],
        [so_ship],
        [so_rmks],
        [so_fob],
        [so_ship_date],
        [so_bol]
    FROM [QADEE2798].[dbo].[so_mstr]
),
SCH_Data AS (
    SELECT
        m.[sch_nbr],
        m.[sch_line],
        m.[sch_rlse_id],
        m.[sch_cr_date],
        m.[sch_ship],
        m.[sch_sd_pat],
        m.[sch_pcr_qty],
        m.[sch_pcs_date],
        m.[sch__chr04],
        m.[sch_lr_asn[1]]],
        m.[sch_lr_qty[1]]],
        m.[sch_lr_cum_qty[1]]],
        m.[sch_eff_start],
        m.[sch_from_pid],
        DATEPART(WEEK, d.[schd_date]) AS [WeekNumber],
        DATEPART(YEAR, d.[schd_date]) AS [YearNumber],
        d.[schd_date],
        SUM(d.[schd_discr_qty]) AS [TotalDiscrQty]
    FROM [QADEE2798].[dbo].[sch_mstr] m
    INNER JOIN [QADEE2798].[dbo].[active_schd_det] d
        ON m.[sch_nbr] = d.[schd_nbr]
        AND m.[sch_line] = d.[schd_line]
        AND m.[sch_rlse_id] = d.[schd_rlse_id]
    WHERE 
        m.[sch_eff_end] IS NULL 
        AND m.[sch_pcr_qty] > 0
    GROUP BY
        m.[sch_nbr],
        m.[sch_line],
        m.[sch_rlse_id],
        m.[sch_cr_date],
        m.[sch_ship],
        m.[sch_sd_pat],
        m.[sch_pcr_qty],
        m.[sch_pcs_date],
        m.[sch__chr04],
        m.[sch_lr_asn[1]]],
        m.[sch_lr_qty[1]]],
        m.[sch_lr_cum_qty[1]]],
        m.[sch_eff_start],
        m.[sch_from_pid],
        DATEPART(WEEK, d.[schd_date]),
        DATEPART(YEAR, d.[schd_date]),
        d.[schd_date]
),
CombinedPS AS (
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
        ps_qty_per AS [Quantity_Per],
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
        m.ps_qty_per AS [Quantity_Per],
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
WeekReference AS (
    SELECT 
        DATEADD(WEEK, 0, DATEADD(DAY, (8 - DATEPART(WEEKDAY, GETDATE())) % 7, CAST(GETDATE() AS DATE))) AS NextWeekStart,
        DATEPART(WEEK, GETDATE()) AS CurrentWeekNumber,
        DATEPART(YEAR, GETDATE()) AS CurrentYear
),
ComponentDemand AS (
    SELECT
        SOD.[sod_site] AS [SO_Site],
        BOM.[component] AS [BOM_Component],
        SCH.[YearNumber],
        SCH.[WeekNumber],
        WR.CurrentWeekNumber,
        WR.CurrentYear,
        WR.NextWeekStart,
        SCH.[schd_date],
        SUM(SCH.[TotalDiscrQty] * BOM.[Quantity_Per]) AS [ComponentDemand]
    FROM SOD_Data SOD
    LEFT JOIN SO_Data SO 
        ON SOD.[sod_nbr] = SO.[so_nbr]
    LEFT JOIN SCH_Data SCH 
        ON SOD.[sod_nbr] = SCH.[sch_nbr] 
        AND SOD.[sod_line] = SCH.[sch_line]
        AND SOD.[sod_curr_rlse_id[1]]] = SCH.[sch_from_pid]
    LEFT JOIN BOMHierarchy BOM 
        ON SOD.[sod_part] = BOM.[root_parent] 
        AND SOD.[sod_site] = BOM.[Plant]
    CROSS JOIN WeekReference WR
    WHERE 
        (SOD.[sod_status] IS NULL OR SOD.[sod_status] <> 'C')
        AND (SOD.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR SOD.[sod_end_eff[1]]] IS NULL)
        AND (SO.[so_rmks] IS NULL OR SO.[so_rmks] <> 'inactive')
        AND SCH.[TotalDiscrQty] IS NOT NULL
        AND BOM.[component] IS NOT NULL
    GROUP BY 
        SOD.[sod_site],
        BOM.[component],
        SCH.[YearNumber],
        SCH.[WeekNumber],
        WR.CurrentWeekNumber,
        WR.CurrentYear,
        WR.NextWeekStart,
        SCH.[schd_date]
),
ComponentDemandWithWeekly AS (
    SELECT
        CD.[SO_Site],
        CD.[BOM_Component],
        SUM(CASE 
            WHEN (CD.[YearNumber] < CD.CurrentYear) OR 
                 (CD.[YearNumber] = CD.CurrentYear AND CD.[WeekNumber] < CD.CurrentWeekNumber) 
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Past_Due],
        SUM(CASE 
            WHEN CD.[YearNumber] = CD.CurrentYear AND CD.[WeekNumber] = CD.CurrentWeekNumber 
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Current_Week],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 1, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 1, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_1],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 2, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 2, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_2],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 3, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 3, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_3],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 4, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 4, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_4],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 5, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 5, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_5],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 6, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 6, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_6],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 7, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 7, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_7],
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 8, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 8, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_8],
        SUM(CASE 
            WHEN (CD.[YearNumber] > DATEPART(YEAR, DATEADD(WEEK, 8, CD.NextWeekStart))) OR 
                 (CD.[YearNumber] = DATEPART(YEAR, DATEADD(WEEK, 8, CD.NextWeekStart)) AND 
                  CD.[WeekNumber] > DATEPART(WEEK, DATEADD(WEEK, 8, CD.NextWeekStart)))
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Future_Demand]
    FROM ComponentDemand CD
    GROUP BY 
        CD.[SO_Site],
        CD.[BOM_Component]
),
LocationDetails AS (
    SELECT 
        xz.[xxwezoned_area_id],
        xz.[xxwezoned_zone_id],
        ld.[ld_loc],
        ld.[ld_part],
        SUM(ld.[ld_qty_oh]) as [ld_qty_oh],
        MAX(ld.[ld_status]) as [ld_status],
        sc.[sct_cst_tot],
        (sc.[sct_mtl_tl] + sc.[sct_mtl_ll]) AS [mat_cost],
        (sc.[sct_cst_tot] - (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [LBO],
        SUM(ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS],
        SUM(ld.[ld_qty_oh] * (sc.[sct_mtl_tl] + sc.[sct_mtl_ll])) AS [CMAT]
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
        ) sc
    ON 
        ld.[ld_part] = sc.[sct_part] 
        AND ld.[ld_site] = sc.[sct_site]
    GROUP BY
        xz.[xxwezoned_area_id],
        xz.[xxwezoned_zone_id],
        ld.[ld_loc],
        ld.[ld_part],
        sc.[sct_cst_tot],
        sc.[sct_mtl_tl],
        sc.[sct_mtl_ll]
),
SerialDetails AS (
    SELECT 
        [ser_pack_code] as [Pack Code],
        [ser_loc] as [Location],
        [ser_part] as [Item Number],
        SUM([ser_qty_avail]) AS [Serial Qty],
        COUNT([ser_serial_id]) AS count_ser_serial_id
    FROM 
        [QADEE2798].[dbo].[ser_active_picked]
    GROUP BY 
        [ser_part],
        [ser_loc],
        [ser_pack_code]
),
FirstQueryResult AS (
    SELECT 
        cp.[pt_part] AS [Item Number],
        ic.[Standard Cost],
        ic.[Material Cost],
        ic.[LBO],
        id.[Item Description], 
        id.[pt_desc2], 
        id.[Prod Line], 
        id.[Group], 
        id.[pt_part_type], 
        id.[Item Number Status],
        id.[Date added],
        id.[ABC], 
        id.[pt_cyc_int], 
        id.[Safety Stock], 
        id.[Safety Time], 
        id.[Item Planner], 
        id.[Item Supplier], 
        id.[Routing], 
        id.[Net weight], 
        id.[pt_net_wt_um], 
        id.[Item Type], 
        id.[Project],
        CASE
            WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
            WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
            WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
            ELSE 'No BOM'
        END AS [FG/SFG/RM],
        psop.[Operation],
        ii.[Total Qty Nettable],
        ii.[Total Qty Nonet],
        ii.[Total Inv],
        ii.[Last_ISSUE],
        ii.[Last_REC],
        ii.[Last_CC],
        ii.[Obsolete],
        (ii.[Total Inv] * ic.[Material Cost]) AS [Total CMAT],
        (ii.[Total Inv] * ic.[Standard Cost]) AS [Total COGS],
        CASE 
            WHEN DATEDIFF(DAY, cp.[pt_added], GETDATE()) < 30 THEN 'New' 
            ELSE NULL 
        END AS [New],
        ia.[WH_Qty],
        ia.[WIP_Qty],
        ia.[EXLPICK_Qty],
        ia.[WH_Qty] * ic.[Standard Cost] AS [WH_Value],
        ia.[WIP_Qty] * ic.[Standard Cost] AS [WIP_Value],
        ia.[EXLPICK_Qty] * ic.[Standard Cost] AS [EXLPICK_Value],
        CASE 
            WHEN ii.[Total Inv] <> 0 THEN 'Yes' 
            ELSE 'No' 
        END AS [Inventory Check],
        CASE 
            WHEN ic.[Standard Cost] = 0 AND CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END <> 'No BOM' THEN 'Yes'
            ELSE NULL
        END AS [No Cost - in BOM],
        CASE 
            WHEN id.[Prod Line] = '0000' AND CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END <> 'No BOM' THEN 'Yes'
            ELSE NULL
        END AS [No Prod Line - in BOM],
        CASE 
            WHEN id.[Group] = 'F000' AND CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END <> 'No BOM' THEN 'Yes'
            ELSE NULL
        END AS [No Group - in BOM],
        CASE 
            WHEN id.[Item Number Status] = 'EPIC' AND CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END <> 'No BOM' THEN 'Yes'
            ELSE NULL
        END AS [EPIC- in BOM],
        CASE 
            WHEN id.[Group] = 'LTH KIT' OR CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END IN ('SFG', 'FG') THEN 'A'
            WHEN id.[Group] = 'COMP' THEN 'B'
            WHEN id.[Group] IN ('Thread', 'Rolls') THEN 'C'
            ELSE 'D'
        END AS [ABC_per_group],
        CASE 
            WHEN id.[Routing] IS NULL AND CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END <> 'No BOM' THEN 'Yes'
            ELSE NULL
        END AS [Routing Missing],
        CASE 
            WHEN id.[Project] IS NULL AND CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END <> 'No BOM' THEN 'Yes'
            ELSE NULL
        END AS [Project missing],
        CASE 
            WHEN COUNT(cp.[pt_part]) OVER (PARTITION BY cp.[pt_part], psop.[Operation]) > 1 
                 AND (CASE WHEN COUNT(*) OVER (PARTITION BY cp.pt_part) > 1 
                      THEN 'yes' 
                      ELSE NULL 
                      END) IS NULL THEN 'Yes'
            ELSE NULL
        END AS [Operation check],
        CASE 
            WHEN id.[pt_cyc_int] - ii.[Last_CC] < 5 THEN 'Yes'
            ELSE NULL
        END AS [Cycle Count Due],
        CASE 
            WHEN ii.[Last_ISSUE] > 90 AND ii.[Last_ISSUE] < 180 THEN 'Yes'
            ELSE NULL
        END AS [Slow-moving Warning],
        CASE 
            WHEN id.[Item Type] <> CASE
                WHEN ISNULL(b.[SFG], 'No') = 'Yes' THEN 'SFG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'Yes' THEN 'FG'
                WHEN ISNULL(b.[SFG], 'No') = 'No' AND ISNULL(b.[Parent], 'No') = 'No' AND ISNULL(b.[Child], 'No') = 'Yes' THEN 'RM'
                ELSE 'No BOM'
            END THEN 'Yes'
            ELSE NULL
        END AS [Item Type Error],
        wip.[total_ISS-WO_CW_-1],
        wip.[total_ISS-WO_CW_-2],
        wip.[total_ISS-WO_CW_-3],
        wip.[total_ISS-WO_CW_-4],
        wip.[recent_wip_min],
        wip.[recent_wip_max]
    FROM 
        CombinedParts cp
    LEFT JOIN 
        BOMStatusCTE b
    ON 
        cp.[pt_site] = b.[Plant] 
        AND cp.[pt_part] = b.[Item Number]
    LEFT JOIN 
        PsOpCTE psop
    ON
        cp.[pt_site] = psop.[Plant]
        AND cp.[pt_part] = psop.[Item Number]
    LEFT JOIN
        ItemCosts ic
    ON
        cp.[pt_site] = ic.[sct_site]
        AND cp.[pt_part] = ic.[sct_part]
    LEFT JOIN
        ItemDetails id
    ON
        cp.[pt_site] = id.[pt_site]
        AND cp.[pt_part] = id.[pt_part]
    LEFT JOIN
        InventoryInfo ii
    ON
        cp.[pt_site] = ii.[in_site]
        AND cp.[pt_part] = ii.[in_part]
    LEFT JOIN
        InventoryByArea ia
    ON
        cp.[pt_site] = ia.[Plant]
        AND cp.[pt_part] = ia.[Item Number]
    LEFT JOIN
        WIPMinimumData wip
    ON
        cp.[pt_part] = wip.[tr_part]
        AND cp.[pt_site] = wip.[tr_site]
)
-- Final SELECT with all columns from first query plus weekly demand, calculated columns, and location details
SELECT DISTINCT 
    FQR.*,
    CDW.[Past_Due],
    CDW.[Current_Week],
    CDW.[Week_1],
    CDW.[Week_2],
    CDW.[Week_3],
    CDW.[Week_4],
    CDW.[Week_5],
    CDW.[Week_6],
    CDW.[Week_7],
    CDW.[Week_8],
    CDW.[Future_Demand],
    -- Calculate Total Demand
    ISNULL(CDW.[Past_Due], 0) + ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + 
    ISNULL(CDW.[Week_2], 0) + ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + 
    ISNULL(CDW.[Week_5], 0) + ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + 
    ISNULL(CDW.[Week_8], 0) + ISNULL(CDW.[Future_Demand], 0) AS [Total Demand],
    -- Calculate Obsolete based on the specified logic
    CASE 
        WHEN (ISNULL(CDW.[Past_Due], 0) + ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + 
              ISNULL(CDW.[Week_2], 0) + ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + 
              ISNULL(CDW.[Week_5], 0) + ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + 
              ISNULL(CDW.[Week_8], 0) + ISNULL(CDW.[Future_Demand], 0)) = 0 
             AND FQR.[Last_REC] > 90 AND FQR.[Last_REC] <= 180 THEN '3 months'
        WHEN (ISNULL(CDW.[Past_Due], 0) + ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + 
              ISNULL(CDW.[Week_2], 0) + ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + 
              ISNULL(CDW.[Week_5], 0) + ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + 
              ISNULL(CDW.[Week_8], 0) + ISNULL(CDW.[Future_Demand], 0)) = 0 
             AND FQR.[Last_REC] > 180 AND FQR.[Last_REC] <= 360 THEN '6 months'
        WHEN (ISNULL(CDW.[Past_Due], 0) + ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + 
              ISNULL(CDW.[Week_2], 0) + ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + 
              ISNULL(CDW.[Week_5], 0) + ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + 
              ISNULL(CDW.[Week_8], 0) + ISNULL(CDW.[Future_Demand], 0)) = 0 
             AND FQR.[Last_REC] > 360 THEN '12 months'
        ELSE NULL
    END AS [Obsolete_per_receipt],
    -- Calculate new WIP_MIN and WIP_MAX based on weekly demand
    CASE 
        WHEN (ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
              ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
              ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) = 0 
        THEN 0 
        ELSE CAST(ROUND((ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                        ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                        ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) / 7.0 * 3, 0) AS INT) 
    END AS [WIP_MIN],
    CASE 
        WHEN (ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
              ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
              ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) = 0 
        THEN 0 
        ELSE CAST(ROUND((ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                        ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                        ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) / 7.0 * 7, 0) AS INT) 
    END AS [WIP_MAX],
    -- Calculate WIP_overstock and WIP_overstock_Value using new WIP_MIN and WIP_MAX
    CASE 
        WHEN FQR.[WIP_Qty] < 0 THEN 0
        WHEN FQR.[WIP_Qty] > 
            CASE 
                WHEN (ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                      ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                      ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) = 0 
                THEN 0 
                ELSE CAST(ROUND((ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                                ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                                ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) / 7.0 * 7, 0) AS INT) 
            END 
        THEN FQR.[WIP_Qty] - 
            CASE 
                WHEN (ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                      ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                      ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) = 0 
                THEN 0 
                ELSE CAST(ROUND((ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                                ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                                ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) / 7.0 * 7, 0) AS INT) 
            END
        ELSE 0
    END AS [WIP_overstock],
    CASE 
        WHEN FQR.[WIP_Qty] < 0 THEN 0
        WHEN FQR.[WIP_Qty] > 
            CASE 
                WHEN (ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                      ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                      ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) = 0 
                THEN 0 
                ELSE CAST(ROUND((ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                                ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                                ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) / 7.0 * 7, 0) AS INT) 
            END 
        THEN (FQR.[WIP_Qty] - 
            CASE 
                WHEN (ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                      ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                      ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) = 0 
                THEN 0 
                ELSE CAST(ROUND((ISNULL(CDW.[Current_Week], 0) + ISNULL(CDW.[Week_1], 0) + ISNULL(CDW.[Week_2], 0) + 
                                ISNULL(CDW.[Week_3], 0) + ISNULL(CDW.[Week_4], 0) + ISNULL(CDW.[Week_5], 0) + 
                                ISNULL(CDW.[Week_6], 0) + ISNULL(CDW.[Week_7], 0) + ISNULL(CDW.[Week_8], 0)) / 7.0 * 7, 0) AS INT) 
            END) * FQR.[Standard Cost]
        ELSE 0
    END AS [WIP_overstock_Value],
    -- Add location details columns from the second query
    ld.[xxwezoned_area_id] as [Area],
    ld.[xxwezoned_zone_id] as [Zone],
    ld.[ld_loc] as [Location],
    ld.[ld_qty_oh] as [Inventory per location],
    ld.[ld_status] as [Inventory Status],
    ld.[sct_cst_tot] as [Standard Cost Location],
    ld.[mat_cost] as [Material cost Location],
    ld.[LBO] as [LBO Location],
    ld.[COGS] as [COGS Location],
    ld.[CMAT] as [CMAT Location],
    -- Add serial details columns from the third query
    ser.[Pack Code],
    ser.[Serial Qty],
    ser.[count_ser_serial_id],
    ld.[ld_qty_oh] - ISNULL(ser.[Serial Qty], 0) AS [Loose Inv]
FROM 
    FirstQueryResult FQR
LEFT JOIN 
    ComponentDemandWithWeekly CDW
ON 
    FQR.[Item Number] = CDW.[BOM_Component]
LEFT JOIN 
    LocationDetails ld
ON 
    FQR.[Item Number] = ld.[ld_part]
LEFT JOIN
    SerialDetails ser
ON 
    FQR.[Item Number] = ser.[Item Number]
    AND ld.[ld_loc] = ser.[Location]
ORDER BY 
    FQR.[Item Number],
    [Location];