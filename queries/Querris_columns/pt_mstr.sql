SELECT [pt__chr02] as [Item Type],
    [pt_part] as [Item Number],
    [pt_desc1] as [Description],
    [pt_prod_line] as [Prod Line],
    [pt_group] as [Group],
    [pt_status] as [Item Status],
    [pt_sfty_stk] as [Safety Stock],
    [pt_dsgn_grp] as [Project],
    [pt_buyer] as [Planner],
    [pt_vend] as [Supplier/Customer],
    [pt_routing] as [Routing],
    CASE 
        WHEN [pt_net_wt_um] = 'kg' THEN [pt_net_wt]
        WHEN [pt_net_wt_um] = 'g' THEN [pt_net_wt] / 1000.0
        ELSE NULL  -- For any other unit of measure
    END as [Net weight in KG],
    CASE 
        WHEN DATEDIFF(day, [pt_added], GETDATE()) < 90 THEN 'New Item'
        ELSE ''
    END as [New Item]
FROM [QADEE2798].[dbo].[pt_mstr]