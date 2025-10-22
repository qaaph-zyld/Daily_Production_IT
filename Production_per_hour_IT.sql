SELECT
    Project,
    ISNULL([6-7], 0) AS [6-7],
    ISNULL([7-8], 0) AS [7-8],
    ISNULL([8-9], 0) AS [8-9],
    ISNULL([9-10], 0) AS [9-10],
    ISNULL([10-11], 0) AS [10-11],
    ISNULL([11-12], 0) AS [11-12],
    ISNULL([12-13], 0) AS [12-13],
    ISNULL([13-14], 0) AS [13-14],
    ISNULL([14-22], 0) AS [14-22],
    ISNULL([II Shift], 0) AS [II Shift]
FROM
(
    SELECT
        th.[tr_qty_loc],
        CASE 
            WHEN pt.[pt_prod_line] = 'H_FG' THEN 'BJA'
            WHEN pt.[pt_prod_line] = 'B_FG' THEN 'BR223 - SEW'
            WHEN pt.[pt_prod_line] = 'C_FG' THEN 'CDPO - ASSY'
            WHEN pt.[pt_prod_line] = 'Z_FG' THEN 'CDPO - SEW'
            WHEN pt.[pt_prod_line] = '0000' THEN 'Pre-production'
            WHEN pt.[pt_prod_line] = 'F_FG' THEN 'FIAT - SEW'
            WHEN pt.[pt_prod_line] = 'K_FG' THEN 'KIA - ASSY'
            WHEN pt.[pt_prod_line] = 'Q_FG' THEN 'KIA - SEW'
            WHEN pt.[pt_prod_line] = 'U_FG' THEN 'MAN'
            WHEN pt.[pt_prod_line] = 'M_FG' THEN 'MMA - ASSY'
            WHEN pt.[pt_prod_line] = 'N_FG' THEN 'MMA - SEW'
            WHEN pt.[pt_prod_line] = 'O_FG' THEN 'OV5X - ASSY'
            WHEN pt.[pt_prod_line] = 'S_FG' THEN 'OV5X - SEW'
            WHEN pt.[pt_prod_line] = 'P_FG' THEN 'PO426 - SEW'
            WHEN pt.[pt_prod_line] = 'G_FG' THEN 'PZ1D'
            WHEN pt.[pt_prod_line] = 'R_FG' THEN 'Renault'
            WHEN pt.[pt_prod_line] = 'E_FG' THEN 'SCANIA'
            WHEN pt.[pt_prod_line] = 'A_FG' THEN 'VOLVO- SEW'
            WHEN pt.[pt_prod_line] = 'V_FG' THEN 'VOLVO- ASSY'
            WHEN pt.[pt_prod_line] = 'T_FG' THEN 'P13A'
            ELSE 'Other' 
        END AS Project,
        CASE
            WHEN (th.[tr_time] / 3600) >= 0 AND (th.[tr_time] / 3600) < 6 THEN '6-7'
            WHEN (th.[tr_time] / 3600) >= 6 AND (th.[tr_time] / 3600) < 7 THEN '6-7'
            WHEN (th.[tr_time] / 3600) >= 7 AND (th.[tr_time] / 3600) < 8 THEN '7-8'
            WHEN (th.[tr_time] / 3600) >= 8 AND (th.[tr_time] / 3600) < 9 THEN '8-9'
            WHEN (th.[tr_time] / 3600) >= 9 AND (th.[tr_time] / 3600) < 10 THEN '9-10'
            WHEN (th.[tr_time] / 3600) >= 10 AND (th.[tr_time] / 3600) < 11 THEN '10-11'
            WHEN (th.[tr_time] / 3600) >= 11 AND (th.[tr_time] / 3600) < 12 THEN '11-12'
            WHEN (th.[tr_time] / 3600) >= 12 AND (th.[tr_time] / 3600) < 13 THEN '12-13'
            WHEN (th.[tr_time] / 3600) >= 13 AND (th.[tr_time] / 3600) < 14 THEN '13-14'
            WHEN (th.[tr_time] / 3600) >= 14 AND (th.[tr_time] / 3600) < 22 THEN '14-22'
            WHEN (th.[tr_time] / 3600) >= 22 AND (th.[tr_time] / 3600) < 24 THEN 'II Shift'
        END AS TimeBucket
    FROM
        [QADEE2798].[dbo].[tr_hist] AS th
    INNER JOIN
        [QADEE2798].[dbo].[pt_mstr] AS pt ON th.[tr_part] = pt.[pt_part]
    WHERE
        th.[tr_type] = 'RCT-WO'
        AND CAST(th.[tr_effdate] AS DATE) = CAST(GETDATE() AS DATE)
) AS SourceData
PIVOT
(
    SUM(tr_qty_loc)
    FOR TimeBucket IN ([6-7], [7-8], [8-9], [9-10], [10-11], [11-12], [12-13], [13-14], [14-22], [II Shift])
) AS PivotTable
ORDER BY
    Project;