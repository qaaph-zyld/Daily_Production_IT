WITH SOD_Data AS (
    SELECT 
        sod.[sod_order_category],
        sod.[sod_nbr],
        sod.[sod_part],
        so.[so_ord_date]  -- Order creation date
    FROM 
        [QADEE2798].[dbo].[sod_det] sod
    INNER JOIN 
        [QADEE2798].[dbo].[so_mstr] so 
        ON sod.[sod_nbr] = so.[so_nbr]
    WHERE 
        (sod.[sod_status] IS NULL OR sod.[sod_status] <> 'C')
        AND (sod.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE) 
             OR sod.[sod_end_eff[1]]] IS NULL)
        AND (so.[so_rmks] IS NULL OR so.[so_rmks] <> 'inactive')
),
SO_Data AS (
    SELECT 
        [so_nbr],
        MAX([so_ship_date]) AS [Last_Ship_Date],
        MAX([so_rmks]) AS [SO_Projects]
    FROM 
        [QADEE2798].[dbo].[so_mstr]
    GROUP BY
        [so_nbr]
),
SCH_Data AS (
    SELECT
        m.[sch_nbr],
        SUM(d.[schd_discr_qty]) AS [TotalDiscrQty]
    FROM 
        [QADEE2798].[dbo].[sch_mstr] m
    INNER JOIN 
        [QADEE2798].[dbo].[active_schd_det] d
        ON m.[sch_nbr] = d.[schd_nbr]
        AND m.[sch_line] = d.[schd_line]
        AND m.[sch_rlse_id] = d.[schd_rlse_id]
    WHERE 
        m.[sch_eff_end] IS NULL 
        AND m.[sch_pcr_qty] > 0
    GROUP BY
        m.[sch_nbr]
),
PartSummary AS (
    SELECT
        SOD.[sod_order_category] AS [SO Line Project],
        SOD.[sod_nbr] AS [SO Number],
        SOD.[sod_part] AS [SO Item Number],
        SO.[SO_Projects] AS [SO Project(s)],
        SO.[Last_Ship_Date] AS [Last Ship Date],
        ISNULL(SCH.[TotalDiscrQty], 0) AS [Total Schedule Discrete Qty],
        CASE 
            WHEN SOD.[so_ord_date] >= DATEADD(DAY, -30, CAST(GETDATE() AS DATE))
                 AND ISNULL(SCH.[TotalDiscrQty], 0) = 0 
            THEN 'New'
            WHEN SO.[Last_Ship_Date] IS NULL THEN
                CASE 
                    WHEN ISNULL(SCH.[TotalDiscrQty], 0) <> 0 THEN 'Active' 
                    ELSE 'Inactive' 
                END
            ELSE
                CASE 
                    WHEN DATEDIFF(DAY, SO.[Last_Ship_Date], GETDATE()) < 90 THEN
                        CASE 
                            WHEN ISNULL(SCH.[TotalDiscrQty], 0) <> 0 THEN 'Active' 
                            ELSE 'Slow Moving' 
                        END
                    WHEN DATEDIFF(DAY, SO.[Last_Ship_Date], GETDATE()) BETWEEN 90 AND 179 THEN
                        CASE 
                            WHEN ISNULL(SCH.[TotalDiscrQty], 0) = 0 THEN '3 Months' 
                            ELSE 'Active' 
                        END
                    WHEN DATEDIFF(DAY, SO.[Last_Ship_Date], GETDATE()) BETWEEN 180 AND 364 THEN
                        CASE 
                            WHEN ISNULL(SCH.[TotalDiscrQty], 0) = 0 THEN '6 Months' 
                            ELSE 'Active' 
                        END
                    WHEN DATEDIFF(DAY, SO.[Last_Ship_Date], GETDATE()) >= 365 THEN
                        CASE 
                            WHEN ISNULL(SCH.[TotalDiscrQty], 0) = 0 THEN '12 Months' 
                            ELSE 'Active' 
                        END
                END
        END AS [Active/Obs Flag]
    FROM 
        SOD_Data SOD
    LEFT JOIN 
        SO_Data SO ON SOD.[sod_nbr] = SO.[so_nbr]
    LEFT JOIN 
        SCH_Data SCH ON SOD.[sod_nbr] = SCH.[sch_nbr]
),
PartDetails AS (
    SELECT 
        PS.*,  -- All columns from the PartSummary CTE
        PT.[pt_desc1],
        PT.[pt_desc2],
        PT.[pt_prod_line],
        PT.[pt_group],
        PT.[pt_part_type],
        PT.[pt_status],
        PT.[pt_buyer],
        PT.[pt_vend],
        PT.[pt_routing]
    FROM 
        PartSummary PS
    INNER JOIN  -- Only include matches
        [QADEE2798].[dbo].[pt_mstr] PT 
        ON PS.[SO Item Number] = PT.[pt_part]
),
InventoryStatus AS (
    SELECT 
        [in_part],
        CASE 
            WHEN DATEDIFF(day, [in_rec_date], GETDATE()) < 90 THEN 'Active'
            WHEN DATEDIFF(day, [in_rec_date], GETDATE()) BETWEEN 90 AND 179 THEN '3 Months'
            WHEN DATEDIFF(day, [in_rec_date], GETDATE()) BETWEEN 180 AND 364 THEN '6 months'
            WHEN DATEDIFF(day, [in_rec_date], GETDATE()) >= 365 THEN '12 months'
        END AS [Last Produced]
    FROM 
        [QADEE2798].[dbo].[15]
)
-- Final join with FULL OUTER JOIN to include all rows from both queries
SELECT 
    COALESCE(PD.[SO Line Project], INV.[Last Produced]) AS [SO Line Project],
    COALESCE(PD.[SO Number], '') AS [SO Number],
    COALESCE(PD.[SO Item Number], INV.[in_part]) AS [SO Item Number],
    COALESCE(PD.[SO Project(s)], '') AS [SO Project(s)],
    PD.[Last Ship Date],
    PD.[Total Schedule Discrete Qty],
    PD.[Active/Obs Flag],
    PD.[pt_desc1],
    PD.[pt_desc2],
    PD.[pt_prod_line],
    PD.[pt_group],
    PD.[pt_part_type],
    PD.[pt_status],
    PD.[pt_buyer],
    PD.[pt_vend],
    PD.[pt_routing],
    INV.[Last Produced]
FROM 
    PartDetails PD
FULL OUTER JOIN 
    InventoryStatus INV ON PD.[SO Item Number] = INV.[in_part]
ORDER BY 
    [SO Line Project],
    [SO Number],
    [SO Item Number];