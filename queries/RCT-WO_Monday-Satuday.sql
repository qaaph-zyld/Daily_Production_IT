-- Declare a variable to hold the start date of the current week (Monday)
DECLARE @StartDateOfWeek DATE;
SET @StartDateOfWeek = DATEADD(day, -(DATEPART(weekday, GETDATE()) + 5) % 7, CAST(GETDATE() AS DATE));

-- Step 1: CTE to prepare the data for the pivot
WITH WeeklyData AS (
    SELECT
        [tr_part],
        [tr_qty_loc],
        DATENAME(weekday, [tr_effdate]) AS DayOfWeek -- Get the day name (e.g., 'Monday')
    FROM
        [QADEE2798].[dbo].[tr_hist]
    WHERE
        [tr_type] = 'rct-wo'
        -- Filter for transactions within the current week
        AND [tr_effdate] >= @StartDateOfWeek 
        AND [tr_effdate] < DATEADD(day, 7, @StartDateOfWeek)
),
-- Step 2: CTE to perform the pivot operation
PivotedData AS (
    SELECT
        [tr_part],
        ISNULL([Monday], 0) AS Monday,
        ISNULL([Tuesday], 0) AS Tuesday,
        ISNULL([Wednesday], 0) AS Wednesday,
        ISNULL([Thursday], 0) AS Thursday,
        ISNULL([Friday], 0) AS Friday,
        ISNULL([Saturday], 0) AS Saturday,
        ISNULL([Sunday], 0) AS Sunday
    FROM
        WeeklyData
    PIVOT
    (
        -- Aggregate the quantity for each part
        SUM([tr_qty_loc]) 
        -- Pivot the unique day names into new columns
        FOR [DayOfWeek] IN ([Monday], [Tuesday], [Wednesday], [Thursday], [Friday], [Saturday], [Sunday])
    ) AS PivotTable
)
-- Step 3: Final SELECT to join with the item master and add the Total column
SELECT
    -- Columns from the PivotedData CTE
    pd.[Monday],
    pd.[Tuesday],
    pd.[Wednesday],
    pd.[Thursday],
    pd.[Friday],
    pd.[Saturday],
    pd.[Sunday],

    -- New Total Column (summing Monday through Saturday as requested)
    (pd.[Monday] + pd.[Tuesday] + pd.[Wednesday] + pd.[Thursday] + pd.[Friday] + pd.[Saturday]) AS Total,

    -- Columns from the item master table (pt_mstr)
    pt.[pt__chr02] AS [Item Type],
    pt.[pt_part] AS [Item Number],
    pt.[pt_desc1] AS [Description],
    pt.[pt_prod_line] AS [Prod Line],
    pt.[pt_group] AS [Group],
    pt.[pt_status] AS [Item Status],
    pt.[pt_dsgn_grp] AS [Project]
FROM
    PivotedData AS pd
LEFT JOIN 
    [QADEE2798].[dbo].[pt_mstr] AS pt
    ON pd.[tr_part] = pt.[pt_part] -- Joining on the part number
ORDER BY
    pt.[pt_part];