SELECT 
    sod.[sod_nbr] as [SO],
    sod.[sod_line] as [SO Line],
    sod.[sod_part] as [Item Number],
    sod.[sod_qty_ship] as [Cum shipped],
    sod.[sod_cum_qty[1]]] as [Cum Required],
    sod.[sod_cum_qty[2]]] as [Prior Day Cum Req],
    sod.[sod_cum_date[2]]] as [Prior Day Cum Date],
    sod.[sod_order_category] as [SO Project],
    schd.[schd_date] as [Date],
    schd.[schd_discr_qty] as [Discrete Qty],
    schd.[schd_cum_qty] as [Cum Shipped (Schedule)]
FROM [QADEE2798].[dbo].[sod_det] sod
LEFT JOIN [QADEE2798].[dbo].[active_schd_det] schd
    ON sod.[sod_nbr] = schd.[schd_nbr] 
    AND sod.[sod_line] = schd.[schd_line]
    AND sod.[sod_curr_rlse_id[3]]] = schd.[schd_rlse_id]
WHERE sod.[sod_status] IS NULL
    AND (sod.[sod_end_eff[1]]] IS NULL OR sod.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE))
ORDER BY sod.[sod_nbr], sod.[sod_line], schd.[schd_date];