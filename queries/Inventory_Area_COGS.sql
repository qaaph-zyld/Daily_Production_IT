SELECT 
    pt.[pt_part] as [Item Number],
    SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WH' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WH],
    SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WH-FG' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WH_FG],
    SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WIP' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WIP],
    SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'EXLPICK' THEN ld_data.[COGS] ELSE 0 END) as [COGS_EXLPICK],
    SUM(CASE WHEN ld_data.[xxwezoned_area_id] = 'WH-FG-E' THEN ld_data.[COGS] ELSE 0 END) as [COGS_WH_FG_E]

FROM 
    [QADEE2798].[dbo].[pt_mstr] pt
LEFT JOIN 
    (
        SELECT 
            xz.[xxwezoned_area_id],
            ld.[ld_part],
            SUM(ld.[ld_qty_oh] * sc.[sct_cst_tot]) AS [COGS]
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
                    [sct_cst_tot]
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
            ld.[ld_part]
    ) ld_data
ON 
    pt.[pt_part] = ld_data.[ld_part]
LEFT JOIN
    (
        SELECT 
            [ser_part] as [Item Number]
       
        FROM 
            [QADEE2798].[dbo].[ser_active_picked]
        GROUP BY 
            [ser_part]
    ) ser_data
ON 
    pt.[pt_part] = ser_data.[Item Number]
WHERE 
    ld_data.[COGS] IS NOT NULL 
    AND ld_data.[COGS] <> 0
    AND ld_data.[xxwezoned_area_id] IN ('WH', 'WH-FG', 'WIP', 'EXLPICK', 'WH-FG-E')
GROUP BY
    pt.[pt_part]
ORDER BY
    pt.[pt_part]