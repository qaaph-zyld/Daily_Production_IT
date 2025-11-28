SELECT 
    [sod_nbr] as [SO],
    [sod_line] as [SO Line],
    [sod_part] as [Item Number],
    [sod_qty_ship] as [Cum shipped],
    [sod_loc] as [SO Location],
    [sod_contr_id] as [PO_SO],
    [sod_cum_qty[1]]] as [Cum Required],
    [sod_cum_qty[2]]] as [Prior Day Cum Req],
    [sod_cum_date[2]]] as [Prior Day Cum Date],
    [sod_curr_rlse_id[1]]] as [Plan Release],
    [sod_curr_rlse_id[3]]] as [Required Release],
    [sod_order_category] as [SO Project]
FROM [QADEE2798].[dbo].[sod_det]
WHERE [sod_status] IS NULL
  AND ([sod_end_eff[1]]] IS NULL OR [sod_end_eff[1]]] > CAST(GETDATE() AS DATE))