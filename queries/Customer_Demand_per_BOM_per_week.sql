WITH SOD_Data AS (
    -- SOD_Data CTE
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
    UNION ALL
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
    FROM [QADEE].[dbo].[sod_det]
),
SO_Data AS (
    -- SO_Data CTE
    SELECT 
        [so_nbr],
        [so_ship],
        [so_rmks],
        [so_fob],
        [so_ship_date],
        [so_bol]
    FROM [QADEE2798].[dbo].[so_mstr]
    UNION ALL
    SELECT 
        [so_nbr],
        [so_ship],
        [so_rmks],
        [so_fob],
        [so_ship_date],
        [so_bol]
    FROM [QADEE].[dbo].[so_mstr]
),
SCH_Data AS (
    -- SCH_Data CTE
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
    UNION ALL
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
    FROM [QADEE].[dbo].[sch_mstr] m
    INNER JOIN [QADEE].[dbo].[active_schd_det] d
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
    -- BOM CTEs
    SELECT 
        '2674' AS [Plant],
        ps_par,
        ps_comp,
        ps_qty_per,
        ps_rmks,
        ps_op,
        ps_ref
    FROM [QADEE].[dbo].[ps_mstr]
    WHERE [ps_end] IS NULL
    UNION ALL
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
-- Identify items that are both parents and components (SFGs)
ItemClassification AS (
    SELECT DISTINCT
        [Plant],
        ps_par AS [Item],
        CASE
            WHEN ps_par IN (SELECT ps_comp FROM CombinedPS c WHERE c.[Plant] = CombinedPS.[Plant]) THEN 1
            ELSE 0
        END AS [IsComponent],
        CASE
            WHEN ps_par IN (SELECT ps_par FROM CombinedPS c WHERE c.[Plant] = CombinedPS.[Plant]) THEN 1
            ELSE 0
        END AS [IsParent]
    FROM CombinedPS
    UNION
    SELECT DISTINCT
        [Plant],
        ps_comp AS [Item],
        CASE
            WHEN ps_comp IN (SELECT ps_comp FROM CombinedPS c WHERE c.[Plant] = CombinedPS.[Plant]) THEN 1
            ELSE 0
        END AS [IsComponent],
        CASE
            WHEN ps_comp IN (SELECT ps_par FROM CombinedPS c WHERE c.[Plant] = CombinedPS.[Plant]) THEN 1
            ELSE 0
        END AS [IsParent]
    FROM CombinedPS
),
ItemTypes AS (
    SELECT
        [Plant],
        [Item],
        CASE
            WHEN [IsComponent] = 1 AND [IsParent] = 1 THEN 'SFG'
            WHEN [IsParent] = 1 THEN 'Parent_Item'
            WHEN [IsComponent] = 1 THEN 'Component'
            ELSE 'Unknown'
        END AS [Item_Type]
    FROM ItemClassification
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
        CASE
            WHEN EXISTS (
                SELECT 1 
                FROM CombinedPS c2 
                WHERE c2.ps_par = c1.ps_comp
                AND c2.[Plant] = c1.[Plant]
            ) THEN 'SFG'
            ELSE 'Component'
        END AS [Structure_Type]
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
-- Add Date References for Week Calculations
WeekReference AS (
    SELECT 
        DATEADD(WEEK, 0, DATEADD(DAY, (8 - DATEPART(WEEKDAY, GETDATE())) % 7, CAST(GETDATE() AS DATE))) AS NextWeekStart,
        DATEPART(WEEK, GETDATE()) AS CurrentWeekNumber,
        DATEPART(YEAR, GETDATE()) AS CurrentYear
),
-- Calculate Component Demand
ComponentDemand AS (
    SELECT
        SOD.[sod_site] AS [SO_Site],
        SOD.[sod_part] AS [Parent_Item],
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
        SOD.[sod_part],
        BOM.[component],
        SCH.[YearNumber],
        SCH.[WeekNumber],
        WR.CurrentWeekNumber,
        WR.CurrentYear,
        WR.NextWeekStart,
        SCH.[schd_date]
),
-- Calculate Parent Item Demand
ParentDemand AS (
    SELECT
        SOD.[sod_site] AS [SO_Site],
        SOD.[sod_part] AS [Item_ID],
        NULL AS [Parent_Item],
        SCH.[YearNumber],
        SCH.[WeekNumber],
        WR.CurrentWeekNumber,
        WR.CurrentYear,
        WR.NextWeekStart,
        SCH.[schd_date],
        SUM(SCH.[TotalDiscrQty]) AS [ItemDemand]
    FROM SOD_Data SOD
    LEFT JOIN SO_Data SO 
        ON SOD.[sod_nbr] = SO.[so_nbr]
    LEFT JOIN SCH_Data SCH 
        ON SOD.[sod_nbr] = SCH.[sch_nbr] 
        AND SOD.[sod_line] = SCH.[sch_line]
        AND SOD.[sod_curr_rlse_id[1]]] = SCH.[sch_from_pid]
    CROSS JOIN WeekReference WR
    WHERE 
        (SOD.[sod_status] IS NULL OR SOD.[sod_status] <> 'C')
        AND (SOD.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR SOD.[sod_end_eff[1]]] IS NULL)
        AND (SO.[so_rmks] IS NULL OR SO.[so_rmks] <> 'inactive')
        AND SCH.[TotalDiscrQty] IS NOT NULL
    GROUP BY 
        SOD.[sod_site],
        SOD.[sod_part],
        SCH.[YearNumber],
        SCH.[WeekNumber],
        WR.CurrentWeekNumber,
        WR.CurrentYear,
        WR.NextWeekStart,
        SCH.[schd_date]
),
-- Combine Component and Parent Demand with Weekly Buckets
CombinedWeeklyDemand AS (
    -- Component Demand
    SELECT
        CD.[SO_Site],
        IT.[Item_Type],
        CD.[Parent_Item],
        CD.[BOM_Component] AS [Item_ID],
        -- Past Due (all weeks before current week)
        SUM(CASE 
            WHEN (CD.[YearNumber] < CD.CurrentYear) OR 
                 (CD.[YearNumber] = CD.CurrentYear AND CD.[WeekNumber] < CD.CurrentWeekNumber) 
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Past_Due],
        -- Current Week
        SUM(CASE 
            WHEN CD.[YearNumber] = CD.CurrentYear AND CD.[WeekNumber] = CD.CurrentWeekNumber 
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Current_Week],
        -- Week 1 (Next Week)
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 1, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 1, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_1],
        -- Week 2
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 2, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 2, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_2],
        -- Week 3
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 3, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 3, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_3],
        -- Week 4
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 4, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 4, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_4],
        -- Week 5
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 5, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 5, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_5],
        -- Week 6
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 6, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 6, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_6],
        -- Week 7
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 7, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 7, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_7],
        -- Week 8
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 8, CD.NextWeekStart)) = CD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 8, CD.NextWeekStart)) = CD.[YearNumber]
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Week_8],
        -- Future Demand (Week 9 and beyond)
        SUM(CASE 
            WHEN (CD.[YearNumber] > DATEPART(YEAR, DATEADD(WEEK, 8, CD.NextWeekStart))) OR 
                 (CD.[YearNumber] = DATEPART(YEAR, DATEADD(WEEK, 8, CD.NextWeekStart)) AND 
                  CD.[WeekNumber] > DATEPART(WEEK, DATEADD(WEEK, 8, CD.NextWeekStart)))
            THEN CD.[ComponentDemand] 
            ELSE 0 
        END) AS [Future_Demand]
    FROM ComponentDemand CD
    LEFT JOIN ItemTypes IT ON CD.[BOM_Component] = IT.[Item] AND CD.[SO_Site] = IT.[Plant]
    GROUP BY 
        CD.[SO_Site],
        IT.[Item_Type],
        CD.[Parent_Item],
        CD.[BOM_Component]
    
    UNION ALL
    
    -- Parent Item Demand
    SELECT
        PD.[SO_Site],
        COALESCE(IT.[Item_Type], 'Parent_Item') AS [Item_Type],
        PD.[Parent_Item],
        PD.[Item_ID],
        -- Past Due (all weeks before current week)
        SUM(CASE 
            WHEN (PD.[YearNumber] < PD.CurrentYear) OR 
                 (PD.[YearNumber] = PD.CurrentYear AND PD.[WeekNumber] < PD.CurrentWeekNumber) 
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Past_Due],
        -- Current Week
        SUM(CASE 
            WHEN PD.[YearNumber] = PD.CurrentYear AND PD.[WeekNumber] = PD.CurrentWeekNumber 
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Current_Week],
        -- Week 1 (Next Week)
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 1, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 1, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_1],
        -- Week 2
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 2, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 2, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_2],
        -- Week 3
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 3, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 3, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_3],
        -- Week 4
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 4, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 4, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_4],
        -- Week 5
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 5, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 5, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_5],
        -- Week 6
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 6, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 6, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_6],
        -- Week 7
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 7, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 7, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_7],
        -- Week 8
        SUM(CASE 
            WHEN DATEPART(WEEK, DATEADD(WEEK, 8, PD.NextWeekStart)) = PD.[WeekNumber] AND 
                 DATEPART(YEAR, DATEADD(WEEK, 8, PD.NextWeekStart)) = PD.[YearNumber]
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Week_8],
        -- Future Demand (Week 9 and beyond)
        SUM(CASE 
            WHEN (PD.[YearNumber] > DATEPART(YEAR, DATEADD(WEEK, 8, PD.NextWeekStart))) OR 
                 (PD.[YearNumber] = DATEPART(YEAR, DATEADD(WEEK, 8, PD.NextWeekStart)) AND 
                  PD.[WeekNumber] > DATEPART(WEEK, DATEADD(WEEK, 8, PD.NextWeekStart)))
            THEN PD.[ItemDemand] 
            ELSE 0 
        END) AS [Future_Demand]
    FROM ParentDemand PD
    LEFT JOIN ItemTypes IT ON PD.[Item_ID] = IT.[Item] AND PD.[SO_Site] = IT.[Plant]
    GROUP BY 
        PD.[SO_Site],
        IT.[Item_Type],
        PD.[Parent_Item],
        PD.[Item_ID]
),
-- Aggregate by Item_ID to ensure distinct values and calculate Min/Max Stock
AggregatedDemand AS (
    SELECT
        [SO_Site],
        [Item_Type],
        [Item_ID],
        SUM([Past_Due]) AS [Past_Due],
        SUM([Current_Week]) AS [Current_Week],
        SUM([Week_1]) AS [Week_1],
        SUM([Week_2]) AS [Week_2],
        SUM([Week_3]) AS [Week_3],
        SUM([Week_4]) AS [Week_4],
        SUM([Week_5]) AS [Week_5],
        SUM([Week_6]) AS [Week_6],
        SUM([Week_7]) AS [Week_7],
        SUM([Week_8]) AS [Week_8],
        SUM([Future_Demand]) AS [Future_Demand],
        -- Calculate daily average of weeks 1-4 (divide by 20 work days)
        (SUM([Week_1]) + SUM([Week_2]) + SUM([Week_3]) + SUM([Week_4])) / 20.0 AS [Daily_Average]
    FROM CombinedWeeklyDemand
    GROUP BY
        [SO_Site],
        [Item_Type],
        [Item_ID]
)
-- Final result
SELECT
    [SO_Site],
    [Item_Type],
    [Item_ID],
    [Past_Due],
    [Current_Week],
    [Week_1],
    [Week_2],
    [Week_3],
    [Week_4],
    [Week_5],
    [Week_6],
    [Week_7],
    [Week_8],
    [Future_Demand],
    -- Min_Stock = Daily_Average * 5
    [Daily_Average] * 5 AS [Min_Stock],
    -- Max_Stock = Daily_Average * 15
    [Daily_Average] * 15 AS [Max_Stock]
FROM AggregatedDemand
ORDER BY 
    [SO_Site],
    [Item_Type],
    [Item_ID];