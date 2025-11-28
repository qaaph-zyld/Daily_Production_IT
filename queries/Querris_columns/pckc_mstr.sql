SELECT 
    [pckc_pack_code] AS [Pack Code],
    [pckc_desc] AS [Pack Code Description],
    CASE 
        WHEN [pckc_ship_wt_um] = 'kg' THEN [pckc_ship_wt]
        WHEN [pckc_ship_wt_um] = 'g' THEN [pckc_ship_wt] / 1000.0
        ELSE NULL  -- For any other unit of measure
    END AS [Pack Ship KG]
FROM [QADEE2798].[dbo].[pckc_mstr]