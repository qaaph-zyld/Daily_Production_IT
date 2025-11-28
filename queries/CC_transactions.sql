SELECT   
    th.tr_site,  
    th.tr_type,  
    th.tr_part,  
    SUM(th.tr_qty_loc) AS [Total Qty Loc],  
    latest_costs.latest_mtl_std AS [Latest Mtl Std],      
    (latest_costs.latest_mtl_std * SUM(th.tr_qty_loc)) AS [CMAT]
FROM [QADEE2798].[dbo].[tr_hist] th
INNER JOIN (
    SELECT 
        tr_site,
        tr_part,
        FIRST_VALUE(tr_mtl_std) OVER (
            PARTITION BY tr_site, tr_part 
            ORDER BY tr_effdate DESC
        ) AS latest_mtl_std
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE tr_effdate > '2025-02-11' AND tr_mtl_std IS NOT NULL
) latest_costs ON th.tr_site = latest_costs.tr_site AND th.tr_part = latest_costs.tr_part
WHERE th.tr_type IN ('rct-tr','iss-tr','iss-wo','iss-scrp','iss-unp')   
    AND th.tr_loc IN ('wip')  
    AND th.tr_qty_loc <> 0   
    AND th.tr_effdate > '2025-02-11'  
GROUP BY   
    th.tr_site,   
    th.tr_part,
    th.tr_type,
    latest_costs.latest_mtl_std
ORDER BY   
    th.tr_site,   
    th.tr_part,
    th.tr_type;