WITH SOD_Data AS (
    -- First part of the UNION ALL for sod_det
    SELECT 
        [sod_order_category],
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
        [sod_unadjust_cum_qty]
    FROM 
        [QADEE2798].[dbo].[sod_det]
    UNION ALL
    -- Second part of the UNION ALL for sod_det
    SELECT 
        [sod_order_category],
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
        [sod_unadjust_cum_qty]
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
),
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
        pd.[pod_curr_rlse_id[1]]]
    FROM  
        [QADEE].[dbo].[pod_det] pd  
        JOIN [QADEE].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
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
        pd.[pod_curr_rlse_id[1]]]
    FROM  
        [QADEE2798].[dbo].[pod_det] pd  
        JOIN [QADEE2798].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
    WHERE  
        pd.[pod_end_eff[1]]] = '2049-12-31 00:00:00'
)
-- Combined query joining SOD_Data, SO_Data, and PodData
SELECT
    SOD.[sod_order_category],
    SOD.[sod_nbr],
    SOD.[sod_line],
    SOD.[sod_part],
    SOD.[sod_qty_ship],
    SOD.[sod_site],
    SOD.[sod_cum_qty[1]]],
    SOD.[sod_cum_qty[2]]],
    SOD.[sod_cum_date[2]]],
    SOD.[sod_curr_rlse_id[1]]],
    SOD.[sod_end_eff[1]]],
    SOD.[sod__qadd04],
    SOD.[sod_unadjust_cum_qty],
    SO.[so_ship],
    SO.[so_rmks],
    SO.[so_fob],
    SO.[so_ship_date],
    SO.[so_bol],
    POD.[pod_po_site],
    POD.[pod__chr08],
    POD.[po_vend],
    POD.[pod_nbr],
    POD.[pod_line],
    POD.[pod_part],
    POD.[pod_cum_qty[1]]] AS pod_cum_qty_1,
    POD.[pod_ord_mult],
    POD.[pod_translt_days],
    POD.[pod_start_eff[1]]],
    POD.[pod_curr_rlse_id[1]]]
FROM 
    SOD_Data SOD
LEFT JOIN 
    SO_Data SO ON SOD.[sod_nbr] = SO.[so_nbr]
LEFT JOIN
    PodData POD ON SOD.[sod_part] = POD.[pod_part]
WHERE
    SO.[so_rmks] <> 'inactive';