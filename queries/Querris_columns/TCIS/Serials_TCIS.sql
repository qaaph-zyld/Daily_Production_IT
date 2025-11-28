SELECT 
    t.[Serial ID],
    t.[trhu_qty],
    t.[trhu_scanned],
    t.[trhu_date],
    i.[ihu_hu],
    i.[ihu_vhilm],
    i.[BKFL time],
    i.[ihu_btch_nbr],
    i.[ihu_piid_name]
FROM 
    (
        SELECT Distinct
            SUBSTRING([trhu_hu1], 2, LEN([trhu_hu1])) as [Serial ID]
            ,[trhu_qty]
            ,[trhu_scanned]
            ,[trhu_date]
        FROM [TCIS_MES2_265].[dbo].[trhu_mstr]
    ) t
LEFT JOIN 
    (
        SELECT 
            [ihu_hu]
            ,[ihu_vhilm]
            ,[ihu_date] as [BKFL time]
            ,[ihu_btch_nbr]
            ,[ihu_piid_name]
        FROM [TCIS_MES2_265].[dbo].[ihu_mstr]
    ) i
    ON t.[Serial ID] = i.[ihu_hu]
	where [BKFL time] >= '2025-09-01' or [BKFL time] is null and [Serial ID] > '430216379'
ORDER BY 
    t.[Serial ID];
	---- trhu_mstr+ihu_mstr