SELECT 
    [in_part],
    [in_site],
    [in_qty_oh],
    [in_iss_date],
    [in_rec_date],
    [in_cnt_date],
    [in_qty_nonet],
    [in_qty_oh] + [in_qty_nonet] AS [total_inv],
    DATEDIFF(day, [in_iss_date], GETDATE()) AS [Last_ISSUE],  -- Calculate days since last issue
    DATEDIFF(day, [in_rec_date], GETDATE()) AS [Last_REC],    -- Calculate days since last receipt
    DATEDIFF(day, [in_cnt_date], GETDATE()) AS [Last_CC],     -- Calculate days since last cycle count
    CASE 
        WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 91 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 180 THEN 'yes'
        ELSE 'no'
    END AS [3 months],  -- Check if Last_ISSUE is between 91 and 179 days
    CASE 
        WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 180 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 365 THEN 'yes'
        ELSE 'no'
    END AS [6 months],  -- Check if Last_ISSUE is between 180 and 364 days
    CASE 
        WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 365 THEN 'yes'
        ELSE 'no'
    END AS [12 months]  -- Check if Last_ISSUE is 365 days or more
FROM 
    [QADEE].[dbo].[in_mstr]
WHERE 
    [in_qty_oh] + [in_qty_nonet] <> 0  -- Filter where total inventory is not zero

UNION ALL

SELECT 
    [in_part],
    [in_site],
    [in_qty_oh],
    [in_iss_date],
    [in_rec_date],
    [in_cnt_date],
    [in_qty_nonet],
    [in_qty_oh] + [in_qty_nonet] AS [total_inv],
    DATEDIFF(day, [in_iss_date], GETDATE()) AS [Last_ISSUE],  -- Calculate days since last issue
    DATEDIFF(day, [in_rec_date], GETDATE()) AS [Last_REC],    -- Calculate days since last receipt
    DATEDIFF(day, [in_cnt_date], GETDATE()) AS [Last_CC],     -- Calculate days since last cycle count
    CASE 
        WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 91 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 180 THEN 'yes'
        ELSE 'no'
    END AS [3 months],  -- Check if Last_ISSUE is between 91 and 179 days
    CASE 
        WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 180 AND DATEDIFF(day, [in_iss_date], GETDATE()) < 365 THEN 'yes'
        ELSE 'no'
    END AS [6 months],  -- Check if Last_ISSUE is between 180 and 364 days
    CASE 
        WHEN DATEDIFF(day, [in_iss_date], GETDATE()) >= 365 THEN 'yes'
        ELSE 'no'
    END AS [12 months]  -- Check if Last_ISSUE is 365 days or more
FROM 
    [QADEE2798].[dbo].[15]
WHERE 
    [in_qty_oh] + [in_qty_nonet] <> 0  -- Filter where total inventory is not zero

ORDER BY 
    [in_part];  -- Apply ORDER BY to the combined result set