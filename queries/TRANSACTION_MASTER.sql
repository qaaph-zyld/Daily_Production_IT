SELECT 

[tr_part]
      ,[tr_type]
      ,[tr_nbr]
      ,[tr_addr]
      ,[tr_lot]
      ,[tr_qty_loc]
      ,[tr_effdate]
      ,[tr_site]
      ,[tr_ship_date]
  FROM [QADEE].[dbo].[tr_hist]
    where [tr_type] ='rct-wo'

	UNION ALL
	SELECT 

[tr_part]
      ,[tr_type]
      ,[tr_nbr]
      ,[tr_addr]
      ,[tr_lot]
      ,[tr_qty_loc]
      ,[tr_effdate]
      ,[tr_site]
      ,[tr_ship_date]
  FROM [QADEE2798].[dbo].[tr_hist]
    where [tr_type] ='rct-wo
'
	order by [tr_part]
;