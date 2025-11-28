SELECT 
    [tr_part],
    [tr_type],
    [tr_site],
    DATEPART(YEAR, [tr_effdate]) AS yearnum,
    DATEPART(WEEK, [tr_effdate]) AS weeknum,
    SUM([tr_qty_loc]) AS total_qty_loc
FROM (
    SELECT 
        [tr_part],
        [tr_type],
        [tr_qty_loc],
        [tr_effdate],
        [tr_site]
    FROM [QADEE].[dbo].[tr_hist]
    WHERE [tr_type] in ( 'rct-wo','iss-so')
	and [tr_effdate] >  '2024-01-10'
    
    UNION ALL
    
    SELECT 
        [tr_part],
        [tr_type],
        [tr_qty_loc],
        [tr_effdate],
        [tr_site]
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE [tr_type] in ( 'rct-wo','iss-so')
	and [tr_effdate] >  '2024-01-10'
    
) AS combined_data
GROUP BY 
    [tr_part],
    [tr_type],
    [tr_site],
    DATEPART(YEAR, [tr_effdate]),
    DATEPART(WEEK, [tr_effdate])
ORDER BY 
    [tr_part],
    yearnum,
    weeknum;