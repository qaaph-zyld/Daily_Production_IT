WITH SOD_Data AS (
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
    FROM 
        [QADEE2798].[dbo].[sod_det]
),
SO_Data AS (
    SELECT 
        [so_nbr],
        [so_ship],
        [so_rmks],
        [so_fob],
        [so_ship_date],
        [so_bol]
    FROM 
        [QADEE2798].[dbo].[so_mstr]
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
        SUM(d.[schd_discr_qty]) AS [TotalDiscrQty]
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
        DATEPART(WEEK, d.[schd_date])
)
-- Final query joining all tables together
SELECT
    SOD.[sod_order_category] AS [SO Line Project],
    SOD.[sod_nbr] AS [SO Number],
    SOD.[sod_line] AS [SO Line],
    SOD.[sod_part] AS [SO Item Number],
    SOD.[sod_qty_ship] AS [SO Qty Shipped],
    SOD.[sod_site] AS [SO Site],
    SOD.[sod_cum_qty[1]]] AS [Cum Shipped],
    SOD.[sod_cum_qty[2]]] AS [Prior Day Cum Shipped],
    SOD.[sod_cum_date[2]]] AS [Prior Day Cum Date],
    SOD.[sod_curr_rlse_id[1]]] AS [SO Plan Release ID],
    SOD.[sod__qadd04] AS [sod__qadd04],
    SOD.[sod_unadjust_cum_qty] AS [sod_unadjust_cum_qty],
    SOD.[sod_loc] AS [SO Location],
    SO.[so_ship] AS [Ship-to],
    SO.[so_rmks] AS [SO Project(s)],
    SO.[so_fob] AS [SO FOB],
    SO.[so_ship_date] AS [Last Ship Date],
    SO.[so_bol] AS [Last Shipper],
    SCH.[sch_rlse_id] AS [Schedule Release ID],
    SCH.[sch_cr_date] AS [Schedule Create Date],
    SCH.[sch_ship] AS [Schedule Ship-to],
    SCH.[sch_sd_pat] AS [Schedule Ship Day Pattern],
    SCH.[sch_pcr_qty] AS [Schedule PCR Qty],
    SCH.[sch_pcs_date] AS [Schedule PCS Date],
    SCH.[sch__chr04] AS [Schedule Chr04],
    SCH.[sch_lr_asn[1]]] AS [Schedule LR ASN],
    SCH.[sch_lr_qty[1]]] AS [Schedule LR Qty],
    SCH.[sch_lr_cum_qty[1]]] AS [Schedule LR Cum Qty],
    SCH.[sch_eff_start] AS [Schedule Effective Start],
    SCH.[WeekNumber] AS [Schedule Week Number],
    SCH.[TotalDiscrQty] AS [Schedule Total Discrete Qty]
FROM 
    SOD_Data SOD
LEFT JOIN 
    SO_Data SO ON SOD.[sod_nbr] = SO.[so_nbr]
LEFT JOIN 
    SCH_Data SCH ON SOD.[sod_nbr] = SCH.[sch_nbr] 
                 AND SOD.[sod_line] = SCH.[sch_line]
                 AND SOD.[sod_curr_rlse_id[1]]] = SCH.[sch_from_pid]
WHERE 
    (SOD.[sod_status] IS NULL OR SOD.[sod_status] <> 'C')
   
    AND (SO.[so_rmks] IS NULL OR SO.[so_rmks] <> 'inactive')
ORDER BY 
    SOD.[sod_site],
    SOD.[sod_nbr],
    SOD.[sod_line],
    SCH.[WeekNumber];