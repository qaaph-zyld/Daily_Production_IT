SELECT Distinct
      SUBSTRING([trhu_hu1], 2, LEN([trhu_hu1])) as [HU Number]
      ,[trhu_qty]
      ,[trhu_scanned]
      ,[trhu_date]
FROM [TCIS_MES2_265].[dbo].[trhu_mstr]
ORDER BY SUBSTRING([trhu_hu1], 2, LEN([trhu_hu1]));