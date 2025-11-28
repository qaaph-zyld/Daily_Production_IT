WITH SCH_Data AS (
    -- SCH/SCHD data from QADEE
    SELECT
        m.[sch_nbr],
        m.[sch_line],
        m.[sch_rlse_id],
        DATEPART(WEEK, d.[schd_date]) AS [WeekNumber],
        SUM(d.[schd_discr_qty]) AS [TotalDiscrQty],
        m.[sch_pcr_qty],
        m.[sch_pcs_date],
        m.[sch_lr_asn[1]]],
        m.[sch_lr_qty[1]]],
        m.[sch_lr_cum_qty[1]]]
    FROM 
        [QADEE].[dbo].[sch_mstr] m
    INNER JOIN 
        [QADEE].[dbo].[active_schd_det] d
        ON m.[sch_nbr] = d.[schd_nbr]
        AND m.[sch_line] = d.[schd_line]
        AND m.[sch_rlse_id] = d.[schd_rlse_id]
    WHERE 
        m.[sch_eff_end] IS NULL 
        AND m.[sch_pcr_qty] > 0
    GROUP BY
        m.[sch_nbr], m.[sch_line], m.[sch_rlse_id],
        m.[sch_pcr_qty], m.[sch_pcs_date], m.[sch_lr_asn[1]]],
        m.[sch_lr_qty[1]]], m.[sch_lr_cum_qty[1]]],
        DATEPART(WEEK, d.[schd_date])
        
    UNION ALL
    
    -- SCH/SCHD data from QADEE2798
    SELECT
        m.[sch_nbr],
        m.[sch_line],
        m.[sch_rlse_id],
        DATEPART(WEEK, d.[schd_date]) AS [WeekNumber],
        SUM(d.[schd_discr_qty]) AS [TotalDiscrQty],
        m.[sch_pcr_qty],
        m.[sch_pcs_date],
        m.[sch_lr_asn[1]]],
        m.[sch_lr_qty[1]]],
        m.[sch_lr_cum_qty[1]]]
    FROM 
        [QADEE2798].[dbo].[sch_mstr] m
    INNER JOIN 
        [QADEE2798].[dbo].[active_schd_det] d
        ON m.[sch_nbr] = d.[schd_nbr]
        AND m.[sch_line] = d.[schd_line]
        AND m.[sch_rlse_id] = d.[schd_rlse_id]
    WHERE 
        m.[sch_eff_end] IS NULL 
        AND m.[sch_pcr_qty] > 0
    GROUP BY
        m.[sch_nbr], m.[sch_line], m.[sch_rlse_id],
        m.[sch_pcr_qty], m.[sch_pcs_date], m.[sch_lr_asn[1]]],
        m.[sch_lr_qty[1]]], m.[sch_lr_cum_qty[1]]],
        DATEPART(WEEK, d.[schd_date])
),

SOD_Data AS (
    SELECT 
        [sod_nbr],
        [sod_line],
        [sod_part],
        [sod_loc],
        [sod_std_cost],
        [sod_custpart],
        [sod_site],
        [sod_prodline],
        [sod_contr_id],
        [sod_cum_qty[1]]] AS sod_cum_qty_1,
        [sod_cum_qty[2]]] AS sod_cum_qty_2,
        [sod_cum_date[1]]] AS sod_cum_date_1,
        [sod_curr_rlse_id[1]]] AS sod_curr_rlse_id_1,
        [sod_curr_rlse_id[3]]] AS sod_curr_rlse_id_3,
        [sod_custref],
        [sod_unadjust_cum_qty]
    FROM [QADEE].[dbo].[sod_det]
    WHERE [sod_status] IS NULL 
      AND ([sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR [sod_end_eff[1]]] IS NULL)
      AND [sod_curr_rlse_id[1]]] IS NOT NULL
      AND [sod_prodline] <> 'N_FG'

    UNION ALL

    SELECT 
        [sod_nbr],
        [sod_line],
        [sod_part],
        [sod_loc],
        [sod_std_cost],
        [sod_custpart],
        [sod_site],
        [sod_prodline],
        [sod_contr_id],
        [sod_cum_qty[1]]] AS sod_cum_qty_1,
        [sod_cum_qty[2]]] AS sod_cum_qty_2,
        [sod_cum_date[1]]] AS sod_cum_date_1,
        [sod_curr_rlse_id[1]]] AS sod_curr_rlse_id_1,
        [sod_curr_rlse_id[3]]] AS sod_curr_rlse_id_3,
        [sod_custref],
        [sod_unadjust_cum_qty]
    FROM [QADEE2798].[dbo].[sod_det]
    WHERE [sod_status] IS NULL 
      AND ([sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR [sod_end_eff[1]]] IS NULL)
      AND [sod_curr_rlse_id[1]]] IS NOT NULL
),

SO_Data AS (
    SELECT 
        [so_nbr],
        CAST([so_ship] AS VARCHAR(255)) AS so_ship,
        [so_fob],
        [so_ship_date],
        [so_bol],
        [so_site]
    FROM [QADEE].[dbo].[so_mstr]
    WHERE [so_nbr] IN ('10006', '10026', '10027', '10028', 'LOZNICA')

    UNION ALL

    SELECT 
        [so_nbr],
        CAST([so_ship] AS VARCHAR(255)) AS so_ship,
        [so_fob],
        [so_ship_date],
        [so_bol],
        [so_site]
    FROM [QADEE2798].[dbo].[so_mstr]
    WHERE [so_nbr] NOT IN ('SO10007', 'SO10009', 'SO10011', 'SO10012', 'SO10017')
),

SOD_SO_Joined AS (
    SELECT 
        SOD.*,
        SO.[so_ship],
        SO.[so_ship_date]
    FROM SOD_Data SOD
    LEFT JOIN SO_Data SO
        ON SOD.[sod_nbr] = SO.[so_nbr]
)

-- Final Combined Result
SELECT
    SCH.[sch_nbr],
    SCH.[sch_line],
    SCH.[WeekNumber],
    SCH.[TotalDiscrQty],
    SCH.[sch_pcr_qty],
    SCH.[sch_pcs_date],
    SOD.[sod_part],
    SOD.[sod_cum_qty_1],
    SOD.[sod_cum_date_1],
    SOD.[so_ship],
    SOD.[so_ship_date],
    SCH.[sch_lr_asn[1]]] AS sch_lr_asn_1,
    SCH.[sch_lr_qty[1]]] AS sch_lr_qty_1,
    SCH.[sch_lr_cum_qty[1]]] AS sch_lr_cum_qty_1
FROM SCH_Data SCH
INNER JOIN SOD_SO_Joined SOD
    ON SCH.[sch_nbr] = SOD.[sod_nbr]
    AND SCH.[sch_line] = SOD.[sod_line]
ORDER BY 
    SCH.[sch_line], 
    SCH.[WeekNumber];