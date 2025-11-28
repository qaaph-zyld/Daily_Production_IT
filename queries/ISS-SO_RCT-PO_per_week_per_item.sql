SELECT 
    [tr_addr] AS [Ship-to],
    [tr_part] AS [Item Number],
    [tr_nbr] AS [PO/SO],
    [tr_site] AS [Plant],
    yearnum,
    weeknum,
    daynum,
    SUM([tr_qty_loc]) AS total_qty_loc
FROM (
    -- SO Transactions
    SELECT 
        [tr_addr],
        [tr_part],
        [tr_nbr],
        [tr_site],
        DATEPART(YEAR, [tr_effdate]) AS yearnum,
        DATEPART(WEEK, [tr_effdate]) AS weeknum,
        DATEPART(DAY, [tr_effdate]) AS daynum,
        [tr_qty_loc]
    FROM [QADEE].[dbo].[tr_hist]
    WHERE [tr_type] = 'iss-so'
        AND [tr_effdate] > '2024-01-10'
    
    UNION ALL
    
    SELECT 
        [tr_addr],
        [tr_part],
        [tr_nbr],
        [tr_site],
        DATEPART(YEAR, [tr_effdate]) AS yearnum,
        DATEPART(WEEK, [tr_effdate]) AS weeknum,
        DATEPART(DAY, [tr_effdate]) AS daynum,
        [tr_qty_loc]
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE [tr_type] = 'iss-so'
        AND [tr_effdate] > '2024-01-10'
    
    UNION ALL
    
    -- PO Transactions
    SELECT 
        [tr_addr],
        [tr_part],
        [tr_nbr],
        [tr_site],
        DATEPART(YEAR, [tr_effdate]) AS yearnum,
        DATEPART(WEEK, [tr_effdate]) AS weeknum,
        DATEPART(DAY, [tr_effdate]) AS daynum,
        [tr_qty_loc]
    FROM [QADEE].[dbo].[tr_hist]
    WHERE [tr_type] = 'rct-po'
        AND [tr_effdate] > '2024-01-10'
    
    UNION ALL
    
    SELECT 
        [tr_addr],
        [tr_part],
        [tr_nbr],
        [tr_site],
        DATEPART(YEAR, [tr_effdate]) AS yearnum,
        DATEPART(WEEK, [tr_effdate]) AS weeknum,
        DATEPART(DAY, [tr_effdate]) AS daynum,
        [tr_qty_loc]
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE [tr_type] = 'rct-po'
        AND [tr_effdate] > '2024-01-10'
) AS combined_data
GROUP BY 
    [tr_addr],
    [tr_part],
    [tr_nbr],
    [tr_site],
    yearnum,
    weeknum,
    daynum
ORDER BY 
    [tr_part],
    yearnum,
    weeknum,
    daynum;