SELECT 
    sch.[sch_nbr],
    sch.[sch_line],
    sch.[sch_rlse_id],
    sch.[sch_cr_date],
    sch.[sch_sd_pat],
    sch.[sch_pcr_qty],
    sch.[sch_pcs_date],
    po.[po_vend],
    po.[po_sd_pat]
FROM [QADEE2798].[dbo].[sch_mstr] sch
LEFT JOIN [QADEE2798].[dbo].[po_mstr] po
    ON sch.[sch_nbr] = po.[po_nbr]
WHERE sch.[sch_eff_end] IS NULL
    AND sch.[sch_type] = '4';