SELECT 
    t.[Serial ID],
    t.[trhu_qty],
    c.[cov_btch_nbr],
    c.[cov_ifs],
    c.[cov_set_nbr],
    c.[cov_counter],
    c.[cov_pck_name],
    c.[cov_planned_qty]
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
RIGHT JOIN 
    (
        SELECT distinct
            c.[cov_btch_nbr],
            SUBSTRING(c.[cov_hu], 2, LEN(c.[cov_hu])) as [Serial ID],
            c.[cov_ifs],
            c.[cov_set_nbr],
            c.[cov_counter],
            c.[cov_pck_name],
            c.[cov_planned_qty]
        FROM 
            [TCIS_MES2_265].[dbo].[cov_hist] c
        INNER JOIN (
            SELECT 
                [cov_btch_nbr],
                MAX(CAST([cov_counter] AS INT)) as MaxCounter
            FROM 
                [TCIS_MES2_265].[dbo].[cov_hist]
            WHERE 
                [cov_valid] = '1' 
                AND [cov_hu] is not null 
                AND [cov_hu] <> ''
                AND [cov_timestamp] >= '2025-09-01'
            GROUP BY 
                [cov_btch_nbr]
        ) m 
            ON c.[cov_btch_nbr] = m.[cov_btch_nbr]
            AND CAST(c.[cov_counter] AS INT) = m.MaxCounter
        WHERE 
            c.[cov_valid] = '1' 
            AND c.[cov_hu] is not null 
            AND c.[cov_hu] <> ''
            AND c.[cov_timestamp] >= '2025-09-01'
    ) c
    ON t.[Serial ID] = c.[Serial ID]
WHERE 
    (i.[BKFL time] >= '2025-09-01' OR i.[BKFL time] IS NULL)
    AND (t.[Serial ID] > '430216379' OR t.[Serial ID] IS NULL)
ORDER BY 
    COALESCE(t.[Serial ID], c.[Serial ID]);