SELECT
      [ld_part] as [Item Number]
	   ,[ld_loc] as [Location]
      ,[ld_qty_oh] as [Qty per location]
      ,[ld_status] as [Inv Status]
  FROM [QADEE2798].[dbo].[ld_det]
  order by [ld_part];
