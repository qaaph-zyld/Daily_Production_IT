SELECT 
    [sct_part] as [Item Number],
    [sct_cst_tot] as [Standard Cost], 
    ([sct_mtl_tl] + [sct_mtl_ll]) AS [CMAT],
    ([sct_cst_tot] - ([sct_mtl_tl] + [sct_mtl_ll])) AS [LBO],
    CASE 
        WHEN [sct_cst_tot] = 0 THEN ''
        WHEN [sct_cst_tot] = ([sct_mtl_tl] + [sct_mtl_ll]) THEN 'RM'
        ELSE 'SFG/FG'
    END AS [Prod/Mfg]
FROM [QADEE2798].[dbo].[sct_det]
WHERE [sct_sim] = 'standard';