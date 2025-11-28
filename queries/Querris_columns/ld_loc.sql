SELECT 
    [ld_loc] as [Location],
    [ld_part] as [Item Number],
    CASE 
        WHEN [ld_status] IN ('AV', 'N - Y - N', 'OBS', 'VH') THEN [ld_qty_oh]  -- Nettable statuses (Netta = Yes)
        ELSE 0  -- Non-nettable statuses (Netta = No)
    END as [MRP Qty],
    CASE 
        WHEN [ld_status] = 'NN' THEN [ld_qty_oh]  -- Special handling for NN status
        ELSE 0
    END as [NN]
FROM [QADEE2798].[dbo].[ld_det]
ORDER BY [Item Number];