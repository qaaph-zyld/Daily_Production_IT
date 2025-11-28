SELECT [sch_nbr]
      ,[sch_line]
      ,[sch_rlse_id]
      ,[sch_cr_date]
      ,[sch_sd_pat]
      ,[sch_pcr_qty]
      ,[sch_pcs_date]
  FROM [QADEE2798].[dbo].[sch_mstr]
  where [sch_eff_end] is null and [sch_type] = '4';