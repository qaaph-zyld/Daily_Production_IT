SELECT
    -- Columns from c_prod_batches (aliased as pb)
    pb.[plan_code],
    pb.[batch_num],
    pb.[sew_finish],

    pb.[state] AS prod_state,
    pb.[qty] AS prod_qty,
    pb.[batch_date],
    pb.[batch_id],
    pb.[tcis_send],
    pb.[creation_user],
    pb.[creation_time] AS prod_creation_time,
    pb.[status_changed],
    pb.[sew_start],
    pb.[total_time],

    -- Columns from the subquery (aliased as batch_details)
    batch_details.[creation_time] AS pack_creation_time,
    batch_details.[cover_ims],
    batch_details.[sewing_start],
    batch_details.[source_app]
FROM 
    [TPDM-TSF].[dbo].[c_prod_batches] AS pb
LEFT JOIN (
    -- This is your second query, acting as a derived table
    SELECT DISTINCT
        cp.[batch_id],
        cp.[creation_time],
        cbc.[cover_ims],
        cbc.[sewing_start],
        cbc.[status],
        cbc.[fin_qty],
        cbc.[source_app]
    FROM 
        [TPDM-TSF].[dbo].[c_batch_pack] AS cp
    INNER JOIN 
        [TPDM-TSF].[dbo].[c_batch_plan] AS cpl ON cp.[batch_id] = cpl.[batch_id]
    LEFT JOIN 
        [TPDM-TSF].[dbo].[c_batch_cover] AS cbc ON cp.[batch_id] = cbc.[batch_id] AND cbc.[sewing_start] > '2025-10-01'
    WHERE 
        cp.[creation_time] > '2025-11-01'
) AS batch_details ON pb.[batch_id] = batch_details.[batch_id]
WHERE 
    pb.[creation_time] > '2025-11-01'
    AND pb.[plan_code] not in ('Signal','Test')
ORDER BY 
    batch_date;