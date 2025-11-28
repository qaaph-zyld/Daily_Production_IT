WITH ParentCTE AS (
    SELECT DISTINCT  
        '2674' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]  -- Adding the Parent column  
    FROM   
        [QADEE].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  

    UNION ALL  

    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_par] AS [Item Number],  
        'Yes' AS [Parent]  -- Adding the Parent column  
    FROM   
        [QADEE2798].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  
),
ChildCTE AS (
    SELECT DISTINCT  
        '2674' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]  -- Adding the Child column  
    FROM   
        [QADEE].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  

    UNION ALL  

    SELECT DISTINCT  
        '2798' AS [Plant],  
        [ps_comp] AS [Item Number],  
        'Yes' AS [Child]  -- Adding the Child column  
    FROM   
        [QADEE2798].[dbo].[ps_mstr]  
    WHERE   
        [ps_end] IS NULL  -- Apply the filter condition  
)
SELECT 
    COALESCE(p.[Plant], c.[Plant]) AS [Plant],  
    COALESCE(p.[Item Number], c.[Item Number]) AS [Item Number],  
    ISNULL(p.[Parent], 'No') AS [Parent],  
    ISNULL(c.[Child], 'No') AS [Child],  
    CASE 
        WHEN p.[Parent] = 'Yes' AND c.[Child] = 'Yes' THEN 'Yes' 
        ELSE 'No' 
    END AS [SFG]  -- Determine if the item is SFG
FROM 
    ParentCTE p
FULL OUTER JOIN 
    ChildCTE c
ON 
    p.[Plant] = c.[Plant] 
    AND p.[Item Number] = c.[Item Number];