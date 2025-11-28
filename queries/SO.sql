WITH SOD_Data AS (
    -- First part of the UNION ALL for sod_det
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
        [sod_status]  -- Ensure sod_status is selected
    FROM 
        [QADEE2798].[dbo].[sod_det]

    UNION ALL

    -- Second part of the UNION ALL for sod_det
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
        [sod_status]  -- Ensure sod_status is selected
    FROM 
        [QADEE].[dbo].[sod_det]
),
SO_Data AS (
    -- First part of the UNION ALL for so_mstr
    SELECT 
        [so_nbr],
        [so_ship],
        [so_rmks],
        [so_fob],
        [so_ship_date],
        [so_bol]
    FROM 
        [QADEE2798].[dbo].[so_mstr]

    UNION ALL

    -- Second part of the UNION ALL for so_mstr
    SELECT 
        [so_nbr],
        [so_ship],
        [so_rmks],
        [so_fob],
        [so_ship_date],
        [so_bol]
    FROM 
        [QADEE].[dbo].[so_mstr]
)
-- Perform the LEFT JOIN and filter
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
    SO.[so_bol] AS [Last Shipper]
FROM 
    SOD_Data SOD
LEFT JOIN 
    SO_Data SO
    ON SOD.[sod_nbr] = SO.[so_nbr]
WHERE 
    -- Handle sod_status: Allow NULL or values not equal to 'C'
    (SOD.[sod_status] IS NULL OR SOD.[sod_status] <> 'C')
    -- Handle sod_end_eff[1]]]: Allow rows where sod_end_eff[1]]] is greater than today or NULL
    AND (SOD.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR SOD.[sod_end_eff[1]]] IS NULL)
    -- Handle so_rmks: Allow NULL or values not equal to 'inactive'
    AND (SO.[so_rmks] IS NULL OR SO.[so_rmks] <> 'inactive')
	
	order by [SO Site],[SO Number],[SO Line];