/*
==============================================================================
MONTHLY DAILY RECEIPTS REPORT - GROUPED BY PROJECT (Read-Only Version)
==============================================================================
Purpose: Generate a dynamic report showing daily receipts for the current month,
         grouped by Project and Production Line.

Features:
- Groups daily quantities by Project and Prod Line (no individual items)
- Dynamic column generation for each day of the current month
- Automatic first/last day detection
- Project mapping based on production line
- Optimized for performance with indexed lookups (if available)
- Open-source SQL Server compatible
- Proper variable scoping for dynamic SQL execution
==============================================================================
*/

-- Ensure we are in the correct database context
USE [QADEE2798];
GO

SET DATEFIRST 1; -- Set Monday as first day of week
SET NOCOUNT ON;   -- Suppress row count messages for performance

-- ============================================================================
-- VARIABLE DECLARATIONS
-- ============================================================================
DECLARE @StartOfMonth DATE;
DECLARE @EndOfMonth DATE;
DECLARE @DaysInMonth INT;
DECLARE @SQL NVARCHAR(MAX);
DECLARE @ColumnList NVARCHAR(MAX) = '';
DECLARE @PivotColumns NVARCHAR(MAX) = '';
DECLARE @CurrentDay INT = 1;

-- Calculate month boundaries
SET @StartOfMonth = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
SET @EndOfMonth = EOMONTH(@StartOfMonth);
SET @DaysInMonth = DAY(@EndOfMonth);

-- ============================================================================
-- EXECUTION METADATA (Before dynamic SQL)
-- ============================================================================
PRINT '==============================================================================';
PRINT 'MONTHLY RECEIPTS REPORT EXECUTION (GROUPED BY PROJECT)';
PRINT '==============================================================================';
PRINT 'Report Period: ' + CONVERT(VARCHAR(10), @StartOfMonth, 120) + ' to ' + CONVERT(VARCHAR(10), @EndOfMonth, 120);
PRINT 'Days in Month: ' + CAST(@DaysInMonth AS VARCHAR(2));
PRINT 'Report Generated: ' + CONVERT(VARCHAR(23), GETDATE(), 121);
PRINT '==============================================================================';
PRINT '';

-- ============================================================================
-- DYNAMIC COLUMN GENERATION
-- ============================================================================
-- Build the column list for SELECT and PIVOT operations
WHILE @CurrentDay <= @DaysInMonth
BEGIN
    -- Add to pivot columns list (used in PIVOT IN clause)
    SET @PivotColumns = @PivotColumns + 
        CASE WHEN @CurrentDay > 1 THEN ', ' ELSE '' END +
        'Day_' + RIGHT('00' + CAST(@CurrentDay AS VARCHAR(2)), 2);
    
    -- MODIFICATION: Add to output column list with SUM() and NULL handling for aggregation
    SET @ColumnList = @ColumnList + 
        CASE WHEN @CurrentDay > 1 THEN ',' + CHAR(13) + CHAR(10) + '    ' ELSE '' END +
        'SUM(ISNULL([Day_' + RIGHT('00' + CAST(@CurrentDay AS VARCHAR(2)), 2) + '], 0)) AS [Day ' + 
        CAST(@CurrentDay AS VARCHAR(2)) + ']';
    
    SET @CurrentDay = @CurrentDay + 1;
END;

-- ============================================================================
-- DYNAMIC SQL QUERY CONSTRUCTION
-- ============================================================================
SET @SQL = N'
-- ============================================================================
-- CTE 1: BASE RECEIPT DATA
-- ============================================================================
-- Extract all receipts for the current month with normalized day format
WITH MonthlyReceipts AS (
    SELECT
        tr.[tr_part] AS ItemNumber,
        ''Day_'' + RIGHT(''00'' + CAST(DAY(tr.[tr_effdate]) AS VARCHAR(2)), 2) AS DayColumn,
        CAST(tr.[tr_qty_loc] AS DECIMAL(18,2)) AS Quantity
    FROM [dbo].[tr_hist] tr
    WHERE tr.[tr_type] = ''rct-wo''
        AND tr.[tr_effdate] >= @StartOfMonth
        AND tr.[tr_effdate] <= @EndOfMonth
        AND tr.[tr_qty_loc] > 0  -- Filter out zero/negative quantities
),

-- ============================================================================
-- CTE 2: AGGREGATED RECEIPT DATA
-- ============================================================================
-- Aggregate quantities by item and day to prepare for pivoting
AggregatedReceipts AS (
    SELECT
        ItemNumber,
        DayColumn,
        SUM(Quantity) AS TotalQuantity
    FROM MonthlyReceipts
    GROUP BY ItemNumber, DayColumn
),

-- ============================================================================
-- CTE 3: PIVOTED DAILY RECEIPTS
-- ============================================================================
-- Transform rows (days) into columns using PIVOT operation
PivotedReceipts AS (
    SELECT
        ItemNumber,
        ' + @PivotColumns + '
    FROM AggregatedReceipts
    PIVOT (
        SUM(TotalQuantity)
        FOR DayColumn IN (' + @PivotColumns + ')
    ) AS PivotTable
),

-- ============================================================================
-- CTE 4: PROJECT MAPPING
-- ============================================================================
-- Join to item master to get production line and map to project name.
-- This creates one row per item, with daily columns.
ProjectMappedData AS (
    SELECT
        -- Project mapping based on production line
        CASE 
            WHEN pt.[pt_prod_line] = ''H_FG'' THEN ''BJA''
            WHEN pt.[pt_prod_line] = ''B_FG'' THEN ''BR223 - SEW''
			WHEN pt.[pt_prod_line] = ''J_FG'' THEN ''JLR - SEW''
			WHEN pt.[pt_prod_line] = ''L_FG'' THEN ''JLR - ASSY''
            WHEN pt.[pt_prod_line] = ''C_FG'' THEN ''CDPO - ASSY''
            WHEN pt.[pt_prod_line] = ''Z_FG'' THEN ''CDPO - SEW''
            WHEN pt.[pt_prod_line] = ''0000'' THEN ''Pre-production''
            WHEN pt.[pt_prod_line] = ''F_FG'' THEN ''FIAT - SEW''
            WHEN pt.[pt_prod_line] = ''K_FG'' THEN ''KIA - ASSY''
            WHEN pt.[pt_prod_line] = ''Q_FG'' THEN ''KIA - SEW''
            WHEN pt.[pt_prod_line] = ''U_FG'' THEN ''MAN''
            WHEN pt.[pt_prod_line] = ''M_FG'' THEN ''MMA - ASSY''
            WHEN pt.[pt_prod_line] = ''N_FG'' THEN ''MMA - SEW''
            WHEN pt.[pt_prod_line] = ''O_FG'' THEN ''OV5X - ASSY''
            WHEN pt.[pt_prod_line] = ''S_FG'' THEN ''OV5X - SEW''
            WHEN pt.[pt_prod_line] = ''P_FG'' THEN ''PO426 - SEW''
            WHEN pt.[pt_prod_line] = ''G_FG'' THEN ''PZ1D''
            WHEN pt.[pt_prod_line] = ''R_FG'' THEN ''Renault''
            WHEN pt.[pt_prod_line] = ''E_FG'' THEN ''SCANIA''
            WHEN pt.[pt_prod_line] = ''A_FG'' THEN ''VOLVO- SEW''
            WHEN pt.[pt_prod_line] = ''V_FG'' THEN ''VOLVO- ASSY''
            WHEN pt.[pt_prod_line] = ''T_FG'' THEN ''P13A''
            ELSE ''Other''
        END AS [Project],
        
        -- Production line from master table
        ISNULL(pt.[pt_prod_line], ''Unknown'') AS [Prod Line],
        
        -- Include all pivoted day columns from the previous CTE
        ' + @PivotColumns + '
    FROM PivotedReceipts pvt
    LEFT JOIN [dbo].[pt_mstr] pt
        ON pvt.ItemNumber = pt.[pt_part]
    WHERE pt.[pt_prod_line] IS NOT NULL  -- Exclude items without production line
)

-- ============================================================================
-- FINAL SELECT WITH GROUPING
-- ============================================================================
-- Aggregate the item-level data up to the Project and Prod Line level
SELECT
    [Project],
    [Prod Line],
    ' + @ColumnList + '
FROM ProjectMappedData
GROUP BY
    [Project],
    [Prod Line]
ORDER BY
    [Project],
    [Prod Line];
';

-- ============================================================================
-- EXECUTE DYNAMIC SQL
-- ============================================================================
EXEC sp_executesql @SQL, 
    N'@StartOfMonth DATE, @EndOfMonth DATE', 
    @StartOfMonth = @StartOfMonth, 
    @EndOfMonth = @EndOfMonth;

PRINT '';
PRINT '==============================================================================';
PRINT 'EXECUTION COMPLETED SUCCESSFULLY';
PRINT '==============================================================================';