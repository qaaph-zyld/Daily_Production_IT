WITH SOD_Data AS (
    -- First part of the UNION ALL for sod_det
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
        [sod_cum_qty[1]]] AS sod_cum_qty_1,  -- Properly escape column names
        [sod_cum_qty[2]]] AS sod_cum_qty_2,
        [sod_cum_date[1]]] AS sod_cum_date_1,
        [sod_curr_rlse_id[1]]] AS sod_curr_rlse_id_1,
        [sod_curr_rlse_id[3]]] AS sod_curr_rlse_id_3,
        [sod_custref],
        [sod_unadjust_cum_qty]
    FROM 
        [QADEE].[dbo].[sod_det]
    WHERE 
        [sod_status] IS NULL 
        AND ([sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR [sod_end_eff[1]]] IS NULL)
        AND [sod_curr_rlse_id[1]]] IS NOT NULL
        AND [sod_prodline] <> 'N_FG'

    UNION ALL

    -- Second part of the UNION ALL for sod_det
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
        [sod_cum_qty[1]]] AS sod_cum_qty_1,  -- Properly escape column names
        [sod_cum_qty[2]]] AS sod_cum_qty_2,
        [sod_cum_date[1]]] AS sod_cum_date_1,
        [sod_curr_rlse_id[1]]] AS sod_curr_rlse_id_1,
        [sod_curr_rlse_id[3]]] AS sod_curr_rlse_id_3,
        [sod_custref],
        [sod_unadjust_cum_qty]
    FROM 
        [QADEE2798].[dbo].[sod_det]
    WHERE 
        [sod_status] IS NULL 
        AND ([sod_end_eff[1]]] > CAST(GETDATE() AS DATE) OR [sod_end_eff[1]]] IS NULL)
        AND [sod_curr_rlse_id[1]]] IS NOT NULL
),
SO_Data AS (
    -- First part of the UNION ALL for so_mstr
    SELECT 
        [so_nbr],
        CAST([so_ship] AS varchar(255)) AS so_ship,  -- Ensure text format
        [so_fob],
        [so_ship_date],
        [so_bol],
        [so_site]
    FROM 
        [QADEE].[dbo].[so_mstr]
    WHERE 
        [so_nbr] IN ('10006', '10026', '10027', '10028', 'LOZNICA')

    UNION ALL

    -- Second part of the UNION ALL for so_mstr
    SELECT 
        [so_nbr],
        CAST([so_ship] AS varchar(255)) AS so_ship,  -- Ensure text format
        [so_fob],
        [so_ship_date],
        [so_bol],
        [so_site]
    FROM 
        [QADEE2798].[dbo].[so_mstr]
    WHERE 
        [so_nbr] NOT IN ('SO10007', 'SO10009', 'SO10011', 'SO10012', 'SO10017')
)
-- Perform the LEFT JOIN between SOD_Data and SO_Data
SELECT
    SOD.[sod_nbr],
    SOD.[sod_line],
    SOD.[sod_part],
    SOD.[sod_site],
    SOD.[sod_cum_qty_1],
    SOD.[sod_cum_qty_2],
    SOD.[sod_cum_date_1],
    SOD.[sod_contr_id],
    SO.[so_ship],
    SO.[so_ship_date],
    SOD.[sod_curr_rlse_id_3]
FROM 
    SOD_Data SOD
LEFT JOIN 
    SO_Data SO
    ON SOD.[sod_nbr] = SO.[so_nbr];