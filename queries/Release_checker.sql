SELECT 
    s.[sch_nbr] AS PO,
    s.[sch_line] AS PO_Line,
    s.[sch_rlse_id] AS Release_ID,
    CONVERT(VARCHAR(10), s.[sch_cr_date], 105) AS Release_Created,
    RIGHT('0' + CAST(s.[sch_cr_time] / 3600 AS VARCHAR(2)), 2) + ':' + 
    RIGHT('0' + CAST((s.[sch_cr_time] % 3600) / 60 AS VARCHAR(2)), 2) AS Release_Created_Time,
    s.[sch_pcr_qty] AS Prior_Cum_Req,
    CONVERT(VARCHAR(10), s.[sch_pcs_date], 105) AS Prior_Cum_Date,
    p.[pod_part] AS Part,
    p.[pod_qty_rcvd] AS Quantity_Received,
    p.[pod_end_eff[1]]] AS End_Effective
FROM 
    [QADEE2798].[dbo].[sch_mstr] s
INNER JOIN 
    [QADEE2798].[dbo].[pod_det] p ON p.[pod_nbr] = s.[sch_nbr] AND p.[pod_line] = s.[sch_line]
WHERE 
    s.[sch__chr04] IS NULL 
    AND s.[sch_type] = '4'
    AND p.[pod_end_eff[1]]] > CAST(GETDATE() AS DATE);