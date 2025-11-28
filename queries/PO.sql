WITH PodData AS (
    SELECT  
        pd.[pod_po_site] AS [Plant],  
        pd.[pod__chr08] AS [PO Item buyer],  
        pm.[po_vend] AS [PO Supplier],  
        pd.[pod_nbr] AS [PO Number],  
        pd.[pod_line] AS [PO Line],  
        pd.[pod_part] AS [Item Number],  
        pd.[pod_cum_qty[1]]] AS [Received Qty],  
        pd.[pod_ord_mult] AS [Standard Pack],  
        pd.[pod_translt_days] AS [Transport Days],
        pd.[pod_start_eff[1]]] AS [Start Effective Date],
        pd.[pod_curr_rlse_id[1]]] AS [Release ID],
        pd.[pod_end_eff[1]]] AS [End Effective Date]
    FROM  
        [QADEE].[dbo].[pod_det] pd  
        JOIN [QADEE].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]  
    UNION ALL  
    SELECT  
        pd.[pod_po_site] AS [Plant],  
        pd.[pod__chr08] AS [PO Item buyer],  
        pm.[po_vend] AS [PO Supplier],  
        pd.[pod_nbr] AS [PO Number],  
        pd.[pod_line] AS [PO Line],  
        pd.[pod_part] AS [Item Number],  
        pd.[pod_cum_qty[1]]] AS [Received Qty],  
        pd.[pod_ord_mult] AS [Standard Pack],  
        pd.[pod_translt_days] AS [Transport Days],
        pd.[pod_start_eff[1]]] AS [Start Effective Date],
        pd.[pod_curr_rlse_id[1]]] AS [Release ID],
        pd.[pod_end_eff[1]]] AS [End Effective Date]
    FROM  
        [QADEE2798].[dbo].[pod_det] pd  
        JOIN [QADEE2798].[dbo].[po_mstr] pm ON pd.[pod_nbr] = pm.[po_nbr]
),
-- Calculate active item counts by plant for duplicate check
ActiveItemCounts AS (
    SELECT 
        [Plant],
        [Item Number],
        COUNT(*) AS ActiveItemCount
    FROM 
        PodData
    WHERE
        [End Effective Date] > GETDATE() -- Only count active items
    GROUP BY 
        [Plant], 
        [Item Number]
),
-- Keep the original ItemCounts for other references if needed
ItemCounts AS (
    SELECT 
        [Plant],
        [Item Number],
        COUNT(*) AS ItemCount
    FROM 
        PodData
    GROUP BY 
        [Plant], 
        [Item Number]
),
-- Identify suppliers present in both plants
DualSuppliers AS (
    SELECT 
        [PO Supplier]
    FROM 
        PodData
    GROUP BY 
        [PO Supplier]
    HAVING 
        COUNT(DISTINCT [Plant]) > 1
),
-- Identify suppliers with inconsistent transport days
TransportDaysCheck AS (
    SELECT
        [PO Supplier],
        CASE 
            WHEN COUNT(DISTINCT [Transport Days]) > 1 THEN 'NOK'
            ELSE NULL
        END AS HasInconsistentTransportDays
    FROM
        PodData
    GROUP BY
        [PO Supplier]
)
SELECT 
    pd.[Plant],
    pd.[PO Item buyer],
    pd.[PO Supplier],
    ad.[ad_name] AS [Supplier Name],
    pd.[PO Number],
    pd.[PO Line],
    pd.[Item Number],
    pd.[Received Qty],
    pd.[Standard Pack],
    pd.[Transport Days],
    pd.[Start Effective Date],
    pd.[Release ID],
    -- Active/Closed Status
    CASE 
        WHEN pd.[End Effective Date] > GETDATE() THEN 'Active'
        ELSE 'Closed'
    END AS [Active/Closed],
    -- Modified Duplicate PO Line check - only flag active duplicates
    CASE 
        WHEN aic.ActiveItemCount > 1 AND pd.[End Effective Date] > GETDATE() THEN 'Yes'
        ELSE NULL
    END AS [Duplicate PO Line],
    -- Dual Supplier check
    CASE 
        WHEN ds.[PO Supplier] IS NOT NULL THEN 'Yes'
        ELSE NULL
    END AS [Dual Supplier],
    -- Transport Days Check
    tdc.HasInconsistentTransportDays AS [Transport Days Check],
    -- Inactive Release Check
    CASE 
        WHEN pd.[End Effective Date] < GETDATE() AND pd.[Release ID] IS NOT NULL THEN 'Yes'
        ELSE NULL
    END AS [Inactive_Release_Check]
FROM 
    PodData pd
LEFT JOIN 
    [QADEE2798].[dbo].[ad_mstr] ad ON pd.[PO Supplier] = ad.[ad_addr]
LEFT JOIN 
    ItemCounts ic ON pd.[Plant] = ic.[Plant] AND pd.[Item Number] = ic.[Item Number]
LEFT JOIN 
    ActiveItemCounts aic ON pd.[Plant] = aic.[Plant] AND pd.[Item Number] = aic.[Item Number]
LEFT JOIN 
    DualSuppliers ds ON pd.[PO Supplier] = ds.[PO Supplier]
LEFT JOIN
    TransportDaysCheck tdc ON pd.[PO Supplier] = tdc.[PO Supplier]
ORDER BY 
    pd.[Plant], 
    pd.[Item Number];