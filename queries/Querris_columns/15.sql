SELECT [in_part] as [Item Number]
      ,[in_qty_oh] as [MRP Qty]
      ,[in_iss_date] as [Last Issue]
      ,[in_rec_date] as [Last Receipt]
      ,[in_cnt_date] as [Last CC]
      ,[in_qty_nonet] as [Non Nettable],[Total Inv] = [in_qty_oh] + [in_qty_nonet]
  FROM [QADEE2798].[dbo].[15]
  order by [in_qty_nonet];