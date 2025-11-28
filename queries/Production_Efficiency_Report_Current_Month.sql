-- Step 1: Create a Common Table Expression (CTE) to hold the pivoted routing data.
-- This makes the main query much cleaner and easier to read.
WITH PivotedRouting AS (
    SELECT
        [ro_routing] AS [Routing],
        SUM(CASE WHEN [ro_op] = 10 THEN [ro_run] ELSE 0 END) AS [Op_10],
        SUM(CASE WHEN [ro_op] = 20 THEN [ro_run] ELSE 0 END) AS [Op_20],
        SUM(CASE WHEN [ro_op] = 30 THEN [ro_run] ELSE 0 END) AS [Op_30],
        SUM(CASE WHEN [ro_op] = 999 THEN [ro_run] ELSE 0 END) AS [Op_999]
    FROM
        [QADEE2798].[dbo].[ro_det]
                where [ro_end] is null
    GROUP BY
        [ro_routing]
)
-- Step 2: Build the main query, joining to the CTE.
SELECT
    th.[tr_part] AS [Item Number],
    th.[tr_type] AS [Transaction Type],
    CAST(th.[tr_effdate] AS DATE) AS [Date],
    th.[tr_prod_line] AS [Prod line],
    pm.[pt__chr02] AS [Item Type],
    pm.[pt_dsgn_grp] AS [Project],
    pm.[pt_drwg_loc] AS [Platform],
    SUM(th.[tr_qty_loc]) AS [Total Qty],
    
    -- Step 3: Select the routing columns from the CTE, applying the conditional logic.
    -- If the Prod line is one of the specified values, the column will be NULL.
    CASE 
        WHEN th.[tr_prod_line] IN ('C_FG','K_FG','L_FG','M_FG','V_FG','O_FG') THEN 0 
        ELSE pr.[Op_10] 
    END AS [Op_10],
    CASE 
        WHEN th.[tr_prod_line] IN ('C_FG','K_FG','L_FG','M_FG','V_FG','O_FG') THEN 0 
        ELSE pr.[Op_20] 
    END AS [Op_20],
    CASE 
        WHEN th.[tr_prod_line] IN ('C_FG','K_FG','L_FG','M_FG','V_FG','O_FG') THEN 0 
        ELSE pr.[Op_30] 
    END AS [Op_30],
    CASE 
        WHEN th.[tr_prod_line] IN ('C_FG','K_FG','L_FG','M_FG','V_FG','O_FG') THEN 0 
        ELSE pr.[Op_999] 
    END AS [Op_999],

    -- Step 4: Add the new calculated column [EQU Sew]
    -- It is also NULL if the Prod line is in the specified list.
    -- ISNULL is used to treat any missing routing times as 0 for the calculation.
    CASE 
        WHEN th.[tr_prod_line] IN ('C_FG','K_FG','L_FG','M_FG','V_FG','O_FG') THEN 0 
        ELSE SUM(th.[tr_qty_loc]) * (ISNULL(pr.[Op_10], 0) + ISNULL(pr.[Op_20], 0) + ISNULL(pr.[Op_30], 0))
    END AS [EQU Sew]

FROM
    [QADEE2798].[dbo].[tr_hist] AS th
INNER JOIN
    [QADEE2798].[dbo].[pt_mstr] AS pm ON th.[tr_part] = pm.[pt_part]
-- Use a LEFT JOIN to include all items from tr_hist, even if they have no routing.
LEFT JOIN
    PivotedRouting AS pr ON th.[tr_part] = pr.[Routing]
WHERE
    -- Filter for 'rct-wo' transactions
    th.[tr_type] = 'rct-wo'
    -- *** MODIFIED DATE FILTER ***
    -- Include transactions from the first day of the current month up to yesterday.
    AND CAST(th.[tr_effdate] AS DATE) >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)
    AND CAST(th.[tr_effdate] AS DATE) < CAST(GETDATE() AS DATE)
GROUP BY
    th.[tr_part],
    th.[tr_type],
    CAST(th.[tr_effdate] AS DATE),
    th.[tr_prod_line],
    pm.[pt__chr02],
    pm.[pt_dsgn_grp],
    pm.[pt_drwg_loc],
    -- We must also group by the routing columns to include them in the SELECT list
    -- since they are not inside an aggregate function in the final query.
    pr.[Op_10],
    pr.[Op_20],
    pr.[Op_30],
    pr.[Op_999]
ORDER BY
    [Date],
    [Item Number];