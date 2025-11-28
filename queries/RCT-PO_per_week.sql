SELECT[tr_addr] as [Ship-to],
    [tr_part] as [Item Number],
	[tr_nbr] as [PO],
    [tr_site] as [Plant],
    DATEPART(YEAR, [tr_effdate]) AS yearnum,
    DATEPART(WEEK, [tr_effdate]) AS weeknum,
	DATEPART(DAY, [tr_effdate]) AS daynum,
    SUM([tr_qty_loc]) AS total_qty_loc
FROM (
    SELECT 
	[tr_addr],
        [tr_part],
        [tr_type],
			[tr_nbr],
        [tr_qty_loc],
        [tr_effdate],
        [tr_site]
    FROM [QADEE].[dbo].[tr_hist]
    WHERE [tr_type] in ( 'rct-po')
	and [tr_effdate] >  '2024-01-10'
    
    UNION ALL
    
    SELECT [tr_addr],
        [tr_part],
        [tr_type],
			[tr_nbr],
        [tr_qty_loc],
        [tr_effdate],
        [tr_site]
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE [tr_type] in ( 'rct-po')
	and [tr_effdate] >  '2024-01-10'
    
) AS combined_data
GROUP BY [tr_addr],
    [tr_part],
    [tr_type],
		[tr_nbr],
    [tr_site],
    DATEPART(YEAR, [tr_effdate]),
    DATEPART(WEEK, [tr_effdate]),
	DATEPART(DAY, [tr_effdate])
ORDER BY 
    [tr_part],
    yearnum,
    weeknum;