SELECT
      [dtset_nbr] as [Set]
      ,[dtset_product_qty] as [Set Qty]
   FROM [TCIS_MES2_265].[dbo].[dtset_mstr]
  where [dtset_last] = 1;