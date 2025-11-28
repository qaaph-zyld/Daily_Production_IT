WITH CombinedSerActivePicked AS (
    SELECT   
        [ser_serial_id] as [Serial ID],
        [ser_stage] as [Serial Stage],   
        [ser_pack_code] as [Pack Code],
        [ser_site] as [Site],  
        [ser_loc] as [Location],  
        [ser_part] as [Item Number],   
        [ser_qty_avail] as [Qty]
    FROM [QADEE2798].[dbo].[ser_active_picked]  
),
CombinedXxwezonedDet AS (
    SELECT 
        [xxwezoned_warehouse_id] as [Site],
        [xxwezoned_area_id] as [Area],
        [xxwezoned_loc] as [Location]
    FROM [QADEE2798].[dbo].[xxwezoned_det]
),
ModifiedPackd AS (
    SELECT 
        SUBSTRING([pckpackd_hu], 2, LEN([pckpackd_hu])) AS pckpackd_hu, -- Remove the first character 'S'
        [pckpackd_qty_pack],
        [pckpackd_qty_act],
        ([pckpackd_qty_pack] - [pckpackd_qty_act]) AS scrap
    FROM 
        [TCIS_MES2_265].[dbo].[pckpackd_det]
    WHERE 
        [pckpackd_qty_act] IS NOT NULL 
        AND [pckpackd_qty_pack] IS NOT NULL 
        AND [pckpackd_qty_act] > 0
        AND ([pckpackd_qty_pack] - [pckpackd_qty_act]) <> 0
),
ScrapData AS (
    SELECT 
        i.ihu_hu,
        i.ihu_plant,
        i.ihu_piid_name,
        i.ihu_vhilm,
        CAST(i.ihu_date AS DATE) AS ihu_date_only,
        CONVERT(VARCHAR(8), CAST(i.ihu_date AS TIME), 108) AS ihu_time_only,
        CASE 
            WHEN i.ihu_plant = '2798' THEN
                CASE 
                    WHEN CONVERT(VARCHAR(8), CAST(i.ihu_date AS TIME), 108) BETWEEN '06:00:00' AND '14:00:00' THEN '1st'
                    ELSE '2nd'
                END
            ELSE
                CASE 
                    WHEN CONVERT(VARCHAR(8), CAST(i.ihu_date AS TIME), 108) BETWEEN '06:00:00' AND '14:00:00' THEN '1st'
                    WHEN CONVERT(VARCHAR(8), CAST(i.ihu_date AS TIME), 108) BETWEEN '14:00:00' AND '22:00:00' THEN '2nd'
                    ELSE '3rd'
                END
        END AS shift,
        CONVERT(VARCHAR(8), CAST(b.ibck_prd_date AS TIME), 108) AS ibck_time_only,
        b.ibck_mat,
        b.ibck_qty,
        b.ibck_btch_nbr,
        b.ibck_hu,
        t.trbtch_ad_plant,
        t.trbtch_prd_qty,
        t.trbtch_mat,
        t.trbtch_comp_ifs,
        t.trbtch_comp_qty,
        t.trbtch_customer,
        t.trbtch_set,
        t.trbtch_workshop,
        t.trbtch_mat_desc,
        ws.ws_name,
        ws.ws_desc,
        c.cust_name,
        c.cust_data1,
        p.scrap
    FROM 
        [TCIS_MES2_265].[dbo].[ihu_mstr] i
    INNER JOIN 
        [TCIS_MES2_265].[dbo].[ibck_mstr] b ON i.ihu_hu COLLATE SQL_Latin1_General_CP1_CI_AS = b.ibck_hu
    INNER JOIN 
        [TCIS_MES2_265].[dbo].[trbtch_mstr] t ON t.trbtch_nbr = b.ibck_btch_nbr
           AND t.trbtch_piid_name = i.ihu_piid_name
           AND t.trbtch_comp_ifs = b.ibck_mat
    LEFT JOIN 
        [TCIS_MES2_265].[dbo].[ws_mstr] ws ON t.trbtch_workshop = ws.ws_code
    LEFT JOIN (
        SELECT 
            cust_kanban,
            cust_name,
            cust_data1,
            ROW_NUMBER() OVER (PARTITION BY cust_kanban ORDER BY cust_name) AS rn
        FROM [TCIS_MES2_265].[dbo].[cust_mstr]
        WHERE cust_kanban IS NOT NULL
    ) c ON t.trbtch_customer = c.cust_kanban AND c.rn = 1 AND t.trbtch_customer IS NOT NULL
    LEFT JOIN 
        ModifiedPackd p ON b.ibck_hu COLLATE SQL_Latin1_General_CP1_CI_AS = p.pckpackd_hu
),
ItemDetails AS (
    SELECT 
        [pt_part] as [Item Number],
        [pt_desc1] as [Item Description],
        [pt_prod_line] as [Prod Line],
        CONVERT(VARCHAR(10), [pt_mod_date], 103) as [Modified Date], -- Format as DD/MM/YYYY
        [pt_dsgn_grp] as [Project]
    FROM [QADEE2798].[dbo].[pt_mstr]
)
SELECT 
    cser.*,
    cxxw.[Area],
    id.[Item Description],
    id.[Prod Line],
    id.[Modified Date],
    id.[Project],
    sd.ihu_piid_name as [Kit Number],
    sd.ihu_date_only as [Production Date],
    sd.ihu_time_only as [Serial Time],
    sd.shift as [Shift],
    sd.ibck_time_only as [Batch_Time],
    sd.ibck_btch_nbr as [Batch ID],
    sd.trbtch_prd_qty as [Batch Qty],
    sd.trbtch_set as [Set],
    sd.trbtch_mat_desc as [Pack Code Description],
    sd.ws_name as [Workshop name],
    sd.cust_name as [Customer],
    sd.cust_data1 as [Customer Address],
    CASE 
        WHEN RIGHT(sd.ihu_piid_name, 1) IN ('L', 'R') AND id.[Item Description] NOT LIKE '%RCB%' THEN 0.5
        ELSE 1
    END AS [Boxes]
FROM CombinedSerActivePicked cser
LEFT JOIN CombinedXxwezonedDet cxxw
    ON cser.[Site] = cxxw.[Site]
    AND cser.[Location] = cxxw.[Location]
LEFT JOIN ScrapData sd
    ON cser.[Serial ID] COLLATE SQL_Latin1_General_CP1_CI_AS = sd.ibck_hu
LEFT JOIN ItemDetails id
    ON cser.[Item Number] = id.[Item Number]
ORDER BY 
    cser.[Site],
    cser.[Location],
    cser.[Item Number];