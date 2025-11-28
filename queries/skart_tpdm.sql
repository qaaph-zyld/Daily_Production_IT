SELECT [model]
      ,[item]
      ,[descr1]
      ,[drawing_nb]
      ,[length]
      ,[date_from]
      ,[material_usage]
  FROM [TPDM-TSF].[dbo].[bom]
  where model = '7021574' and date_to = '2099-01-01' and material_usage > 0
  ORDER BY item;
