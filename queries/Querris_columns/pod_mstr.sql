SELECT 
    [pod_nbr],
    [pod_line],
    [pod_part],
    [pod_qty_rcvd],
    [pod_qty_rtnd],
    [pod__qad05],
    [pod__chr04],
    [pod__chr08],
    [pod_cum_qty[1]]],
    [pod_plan_weeks],
    [pod_curr_rlse_id[1]]],
    [pod_ord_mult],
    [pod_translt_days],
    [pod_sd_pat],
    [pod_plan_mths],
    [pod_firm_days],
    [pod_sftylt_days],
	[pod_status]
FROM [QADEE2798].[dbo].[pod_det]
WHERE [pod_end_eff[1]]] > CAST(GETDATE() AS DATE)