SELECT
    COALESCE(ld.ld_part, pt.pt_part) AS [Item Number], -- Handles NULLs from RIGHT JOIN
    ld.ld_loc AS [Location],
    ld.ld_qty_oh AS [Qty per location],
    ld.ld_status AS [Inv Status],
    pt.pt__chr02 AS [Item Type],
    pt.pt_desc1 AS [Description],
    pt.pt_prod_line AS [Prod Line],
    pt.pt_group AS [Group],
    pt.pt_status AS [Item Status],
    pt.pt_sfty_stk AS [Safety Stock],
    pt.pt_dsgn_grp AS [Project],
    pt.pt_buyer AS [Planner],
    pt.pt_vend AS [Supplier/Customer],
    pt.pt_routing AS [Routing],
    CASE 
        WHEN pt.pt_net_wt_um = 'kg' THEN pt.pt_net_wt
        WHEN pt.pt_net_wt_um = 'g' THEN pt.pt_net_wt / 1000.0
        ELSE NULL
    END AS [Net weight in KG],
    CASE 
        WHEN DATEDIFF(day, pt.pt_added, GETDATE()) < 90 THEN 'New Item'
        ELSE ''
    END AS [New Item]
FROM [QADEE2798].[dbo].[ld_det] AS ld
RIGHT JOIN [QADEE2798].[dbo].[pt_mstr] AS pt 
    ON ld.ld_part = pt.pt_part -- Join condition
ORDER BY [Item Number]; -- Order by the unified Item Number