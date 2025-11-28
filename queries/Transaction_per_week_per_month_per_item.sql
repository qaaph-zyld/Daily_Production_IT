DECLARE @DynamicPivotSQL NVARCHAR(MAX);
DECLARE @ColumnList NVARCHAR(MAX);

-- Create temporary table for time period management
CREATE TABLE #TimePeriods (
    time_period NVARCHAR(50),
    sort_order INT,
    week_start_date DATE
);

-- Calculate the last 30 weeks with proper chronological ordering
WITH WeekCalculation AS (
    SELECT 
        DATEADD(WEEK, DATEDIFF(WEEK, 0, GETDATE()), 0) AS CurrentWeekStart
),
WeekRange AS (
    SELECT 
        week_offset,
        DATEADD(WEEK, -week_offset, CurrentWeekStart) AS week_start_date,
        DATEADD(WEEK, -week_offset + 1, CurrentWeekStart) AS week_end_date
    FROM WeekCalculation
    CROSS JOIN (
        SELECT 0 AS week_offset UNION ALL SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL SELECT 4 UNION ALL
        SELECT 5 UNION ALL SELECT 6 UNION ALL SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9 UNION ALL
        SELECT 10 UNION ALL SELECT 11 UNION ALL SELECT 12 UNION ALL SELECT 13 UNION ALL SELECT 14 UNION ALL
        SELECT 15 UNION ALL SELECT 16 UNION ALL SELECT 17 UNION ALL SELECT 18 UNION ALL SELECT 19 UNION ALL
        SELECT 20 UNION ALL SELECT 21 UNION ALL SELECT 22 UNION ALL SELECT 23 UNION ALL SELECT 24 UNION ALL
        SELECT 25 UNION ALL SELECT 26 UNION ALL SELECT 27 UNION ALL SELECT 28 UNION ALL SELECT 29
    ) AS weeks
),
WeeklyPeriods_2798 AS (
    SELECT DISTINCT 
        'Week ' + FORMAT(wr.week_start_date, 'yyyy-MM-dd') AS time_period,
        1000 - wr.week_offset AS sort_order,  -- Reverse order for chronological display (most recent first)
        wr.week_start_date
    FROM [QADEE2798].[dbo].[tr_hist] th
    INNER JOIN WeekRange wr ON th.[tr_effdate] >= wr.week_start_date 
                            AND th.[tr_effdate] < wr.week_end_date
    WHERE th.[tr_type] NOT IN ('cst-adj','cyc-cnt','iss-chl','rct-chl','cyc-err','ord-so','cum-rres','rct-adj','cum-radj','cum-sadj','iss-tr','rct-tr')
    AND th.[tr_qty_loc] <> 0
)

-- Populate time periods table with weeks from QADEE2798
INSERT INTO #TimePeriods (time_period, sort_order, week_start_date)
SELECT DISTINCT time_period, sort_order, week_start_date 
FROM WeeklyPeriods_2798;

-- Generate dynamic column list with proper SQL escaping
SELECT @ColumnList = COALESCE(@ColumnList + ', ', '') + QUOTENAME(time_period)
FROM #TimePeriods
ORDER BY sort_order DESC; -- Most recent week first

-- Construct optimized dynamic pivot SQL
SET @DynamicPivotSQL = N'
WITH WeekCalculation AS (
    SELECT 
        DATEADD(WEEK, DATEDIFF(WEEK, 0, GETDATE()), 0) AS CurrentWeekStart
),
WeekRange AS (
    SELECT 
        week_offset,
        DATEADD(WEEK, -week_offset, CurrentWeekStart) AS week_start_date,
        DATEADD(WEEK, -week_offset + 1, CurrentWeekStart) AS week_end_date
    FROM WeekCalculation
    CROSS JOIN (
        SELECT 0 AS week_offset UNION ALL SELECT 1 UNION ALL SELECT 2 UNION ALL SELECT 3 UNION ALL SELECT 4 UNION ALL
        SELECT 5 UNION ALL SELECT 6 UNION ALL SELECT 7 UNION ALL SELECT 8 UNION ALL SELECT 9 UNION ALL
        SELECT 10 UNION ALL SELECT 11 UNION ALL SELECT 12 UNION ALL SELECT 13 UNION ALL SELECT 14 UNION ALL
        SELECT 15 UNION ALL SELECT 16 UNION ALL SELECT 17 UNION ALL SELECT 18 UNION ALL SELECT 19 UNION ALL
        SELECT 20 UNION ALL SELECT 21 UNION ALL SELECT 22 UNION ALL SELECT 23 UNION ALL SELECT 24 UNION ALL
        SELECT 25 UNION ALL SELECT 26 UNION ALL SELECT 27 UNION ALL SELECT 28 UNION ALL SELECT 29
    ) AS weeks
),
WeeklyData_2798 AS (
    SELECT 
        th.[tr_part],
        th.[tr_type],
        th.[tr_site],
        SUM(th.[tr_qty_loc]) AS total_qty_loc,
        ''Week '' + FORMAT(wr.week_start_date, ''yyyy-MM-dd'') AS time_period
    FROM [QADEE2798].[dbo].[tr_hist] th
    INNER JOIN WeekRange wr ON th.[tr_effdate] >= wr.week_start_date 
                            AND th.[tr_effdate] < wr.week_end_date
    WHERE th.[tr_type] NOT IN (''cst-adj'',''cyc-cnt'',''iss-chl'',''rct-chl'',''cyc-err'',''ord-so'',''cum-rres'',''rct-adj'',''cum-radj'',''cum-sadj'',''iss-tr'',''rct-tr'')
    AND th.[tr_qty_loc] <> 0
    GROUP BY 
        th.[tr_part],
        th.[tr_type],
        th.[tr_site],
        ''Week '' + FORMAT(wr.week_start_date, ''yyyy-MM-dd'')
)

-- Execute pivot transformation with proper NULL handling
SELECT 
    pvt.[tr_part],
    pvt.[tr_type],
    pvt.[tr_site],
    ' + @ColumnList + '
FROM WeeklyData_2798
PIVOT (
    SUM([total_qty_loc]) 
    FOR [time_period] IN (' + @ColumnList + ')
) AS pvt
ORDER BY 
    pvt.[tr_part],
    pvt.[tr_type],
    pvt.[tr_site];';

-- Execute the optimized dynamic SQL
EXEC sp_executesql @DynamicPivotSQL;

-- Optional: Print SQL for debugging and optimization review
-- PRINT @DynamicPivotSQL;

-- Resource cleanup
DROP TABLE #TimePeriods;

-- Performance monitoring query (optional)
/*
PRINT 'Query execution completed for last 30 weeks analysis';
PRINT 'Column structure: ' + @ColumnList;
PRINT 'Execution timestamp: ' + CONVERT(VARCHAR, GETDATE(), 120);
*/