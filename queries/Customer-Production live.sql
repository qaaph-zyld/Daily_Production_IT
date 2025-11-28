-- Declare a variable to hold the start date of the current week (Monday)
SET DATEFIRST 1; -- This must be at the top of the batch
DECLARE @StartDateOfWeek DATE;
-- CORRECTED: This calculation now reliably finds the Monday of the current week.
SET @StartDateOfWeek = DATEADD(day, -(DATEPART(weekday, GETDATE()) - 1), CAST(GETDATE() AS DATE));

-- Step 1: CTE for the first query (Weekly Receipts)
WITH WeeklyReceiptsPivoted AS (
    SELECT
        pd.[tr_part] AS [Item Number],
        pd.[Monday],
        pd.[Tuesday],
        pd.[Wednesday],
        pd.[Thursday],
        pd.[Friday],
        pd.[Saturday],
        pd.[Sunday],
        (pd.[Monday] + pd.[Tuesday] + pd.[Wednesday] + pd.[Thursday] + pd.[Friday] + pd.[Saturday]) AS [Total Received]
    FROM (
        SELECT
            [tr_part],
            ISNULL([Monday], 0) AS Monday,
            ISNULL([Tuesday], 0) AS Tuesday,
            ISNULL([Wednesday], 0) AS Wednesday,
            ISNULL([Thursday], 0) AS Thursday,
            ISNULL([Friday], 0) AS Friday,
            ISNULL([Saturday], 0) AS Saturday,
            ISNULL([Sunday], 0) AS Sunday
        FROM (
            SELECT
                [tr_part],
                [tr_qty_loc],
                DATENAME(weekday, [tr_effdate]) AS DayOfWeek
            FROM [QADEE2798].[dbo].[tr_hist]
            WHERE [tr_type] = 'rct-wo'
                AND [tr_effdate] >= @StartDateOfWeek 
                AND [tr_effdate] < DATEADD(day, 7, @StartDateOfWeek)
        ) AS WeeklyData
        PIVOT (
            SUM([tr_qty_loc]) 
            FOR [DayOfWeek] IN ([Monday], [Tuesday], [Wednesday], [Thursday], [Friday], [Saturday], [Sunday])
        ) AS PivotTable
    ) AS pd
),
-- Step 2: CTE for the base data of the second query (Scheduled Demand)
ScheduledDemandBaseData AS (
    SELECT 
        sod.[sod_part] AS [Item Number],
        schd.[schd_date] AS [Date],
        CAST(schd.[schd_discr_qty] AS INT) AS [Discrete Qty],
        CASE 
            WHEN schd.[schd_date] < CAST(GETDATE() AS DATE) THEN 'Past Due'
            WHEN DATEDIFF(WEEK, GETDATE(), schd.[schd_date]) BETWEEN 1 AND 8 THEN 
                'Week ' + CAST(DATEDIFF(WEEK, GETDATE(), schd.[schd_date]) AS VARCHAR(10))
            WHEN DATEDIFF(WEEK, GETDATE(), schd.[schd_date]) >= 9 THEN 'Long Term'
            ELSE 'Other'
        END AS [Time Period]
    FROM [QADEE2798].[dbo].[sod_det] sod
    LEFT JOIN [QADEE2798].[dbo].[active_schd_det] schd
        ON sod.[sod_nbr] = schd.[schd_nbr] 
        AND sod.[sod_line] = schd.[schd_line]
        AND sod.[sod_curr_rlse_id[3]]] = schd.[schd_rlse_id]
    WHERE sod.[sod_status] IS NULL
        AND (sod.[sod_end_eff[1]]] IS NULL OR sod.[sod_end_eff[1]]] > CAST(GETDATE() AS DATE))
        AND schd.[schd_date] IS NOT NULL
),
-- Step 3: CTE for the aggregated results of the second query
ScheduledDemandAggregated AS (
    SELECT 
        [Item Number],
        SUM(CAST(CASE WHEN [Time Period] = 'Week 1' THEN [Discrete Qty] ELSE 0 END AS INT)) AS [Week 1]
    FROM ScheduledDemandBaseData
    GROUP BY [Item Number]
    HAVING 
        SUM(CAST(CASE WHEN [Time Period] = 'Week 1' THEN [Discrete Qty] ELSE 0 END AS INT)) > 0
)
-- Step 4: Final SELECT to join the two CTEs and the item master
SELECT
    -- Use COALESCE to get the Item Number from whichever CTE has it
    COALESCE(wr.[Item Number], sd.[Item Number]) AS [Item Number],

    -- Columns from WeeklyReceipts (use ISNULL to show 0 if no match)
    ISNULL(wr.[Monday], 0) AS Monday,
    ISNULL(wr.[Tuesday], 0) AS Tuesday,
    ISNULL(wr.[Wednesday], 0) AS Wednesday,
    ISNULL(wr.[Thursday], 0) AS Thursday,
    ISNULL(wr.[Friday], 0) AS Friday,
    ISNULL(wr.[Saturday], 0) AS Saturday,
    ISNULL(wr.[Sunday], 0) AS Sunday,
    ISNULL(wr.[Total Received], 0) AS [Total Received],

    -- Columns from ScheduledDemand (use ISNULL to show 0 if no match)
    ISNULL(sd.[Week 1], 0) AS [Week 1],

    -- *** NEW COLUMN: Percentage Calculation ***
    CASE 
        WHEN ISNULL(sd.[Week 1], 0) = 0 THEN 0 -- Avoid division by zero
        ELSE (ISNULL(wr.[Total Received], 0) * 100.0) / ISNULL(sd.[Week 1], 0)
    END AS [Received vs Week 1 %],

    -- Columns from the item master table
    pt.[pt__chr02] AS [Item Type],
    pt.[pt_desc1] AS [Description], 
    pt.[pt_prod_line] AS [Prod Line],
    pt.[pt_group] AS [Group],
    pt.[pt_status] AS [Item Status],

    -- *** MODIFIED: Project column based on Prod Line mapping ***
    CASE 
        WHEN pt.[pt_prod_line] = 'H_FG' THEN 'BJA'
        WHEN pt.[pt_prod_line] = 'B_FG' THEN 'BR223 - SEW'
        WHEN pt.[pt_prod_line] = 'C_FG' THEN 'CDPO - ASSY'
        WHEN pt.[pt_prod_line] = 'Z_FG' THEN 'CDPO - SEW'
        WHEN pt.[pt_prod_line] = '0000' THEN 'Pre-production'
        WHEN pt.[pt_prod_line] = 'F_FG' THEN 'FIAT - SEW'
        WHEN pt.[pt_prod_line] = 'K_FG' THEN 'KIA - ASSY'
        WHEN pt.[pt_prod_line] = 'Q_FG' THEN 'KIA - SEW'
        WHEN pt.[pt_prod_line] = 'U_FG' THEN 'MAN'
        WHEN pt.[pt_prod_line] = 'M_FG' THEN 'MMA - ASSY'
        WHEN pt.[pt_prod_line] = 'N_FG' THEN 'MMA - SEW'
        WHEN pt.[pt_prod_line] = 'O_FG' THEN 'OV5X - ASSY'
        WHEN pt.[pt_prod_line] = 'S_FG' THEN 'OV5X - SEW'
        WHEN pt.[pt_prod_line] = 'P_FG' THEN 'PO426 - SEW'
        WHEN pt.[pt_prod_line] = 'G_FG' THEN 'PZ1D'
        WHEN pt.[pt_prod_line] = 'R_FG' THEN 'Renault'
        WHEN pt.[pt_prod_line] = 'E_FG' THEN 'SCANIA'
        WHEN pt.[pt_prod_line] = 'A_FG' THEN 'VOLVO- SEW'
        WHEN pt.[pt_prod_line] = 'V_FG' THEN 'VOLVO- ASSY'
        WHEN pt.[pt_prod_line] = 'T_FG' THEN 'P13A'
        ELSE 'Other' -- Default value for any unmapped prod lines
    END AS [Project]

FROM 
    WeeklyReceiptsPivoted AS wr
FULL OUTER JOIN -- Use FULL OUTER JOIN to include all items from both lists
    ScheduledDemandAggregated AS sd 
    ON wr.[Item Number] = sd.[Item Number]
LEFT JOIN -- Then join to the item master to get descriptive data
    [QADEE2798].[dbo].[pt_mstr] AS pt
    ON COALESCE(wr.[Item Number], sd.[Item Number]) = pt.[pt_part]
ORDER BY
    [Item Number];