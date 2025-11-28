SELECT 
    '2798' AS [Plant],
    [ps_par],  
    [ps_comp],  
    [ps_ref],  
    [ps_qty_per],  
    [ps_rmks],  
    [ps_op],  
    [ps_userid]

FROM 
    [QADEE2798].[dbo].[ps_mstr]
WHERE 
    [ps_end] is null;  -- Apply the filter condition