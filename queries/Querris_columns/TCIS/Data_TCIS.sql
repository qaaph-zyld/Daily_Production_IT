SELECT
    m.[dtset_nbr] as [Set],
    m.[dtset_product_qty] as [Set Qty],
    d.[dtbin_cov_ver] as [Cover Level],
    d.[dtbin_ifs] as [Item Number],
    d.[dtbin_pi] as [Kit Number],
    d.[dtbin_set_ver] as [Set Level],
    dm.[data_product_qty] as [data_product_qty],
    dm.[data_left],
    dm.[data_right],
    dm.[data_item_nbr_desc1] as [Item Description],
    dm.[data_model] as [Model],
    dm.[data_cust_name] as [Customer Name],
    dm.[data_workshop] as [Workshop]
FROM 
    [TCIS_MES2_265].[dbo].[dtset_mstr] m
INNER JOIN 
    [TCIS_MES2_265].[dbo].[dtbin_det] d
    ON m.[dtset_nbr] = d.[dtbin_set_nbr]
INNER JOIN 
    [TCIS_MES2_265].[dbo].[data_mstr] dm
    ON m.[dtset_nbr] = dm.[data_set] 
    AND d.[dtbin_ifs] = dm.[data_ifs]
WHERE 
    m.[dtset_last] = 1;
	------dtset_mstr+dtbin_det+data_mstr

