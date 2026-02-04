/*
==============================================================================
MONTHLY DAILY RECEIPTS REPORT - 3 RESULT SETS (PROJECT, SEW, ASSY)
==============================================================================
Purpose: Generate a dynamic report showing daily receipts for the current month.
         Returns 3 separate result sets.
==============================================================================
*/

USE [QADEE2798];
GO

SET DATEFIRST 1; 
SET NOCOUNT ON;   

DECLARE @StartOfMonth DATE;
DECLARE @EndOfMonth DATE;
DECLARE @DaysInMonth INT;
DECLARE @SQL NVARCHAR(MAX);
DECLARE @ColumnList NVARCHAR(MAX) = '';
DECLARE @PivotColumns NVARCHAR(MAX) = '';
DECLARE @TableCols NVARCHAR(MAX) = ''; -- New variable for CREATE TABLE definition
DECLARE @CurrentDay INT = 1;

SET @StartOfMonth = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
SET @EndOfMonth = EOMONTH(@StartOfMonth);
SET @DaysInMonth = DAY(@EndOfMonth);

PRINT '==============================================================================';
PRINT 'MONTHLY RECEIPTS REPORT EXECUTION (3 RESULT SETS)';
PRINT '==============================================================================';
PRINT 'Report Period: ' + CONVERT(VARCHAR(10), @StartOfMonth, 120) + ' to ' + CONVERT(VARCHAR(10), @EndOfMonth, 120);
PRINT '==============================================================================';
PRINT '';

WHILE @CurrentDay <= @DaysInMonth
BEGIN
    -- Columns for the PIVOT table (Raw daily values)
    SET @PivotColumns = @PivotColumns + 
        CASE WHEN @CurrentDay > 1 THEN ', ' ELSE '' END +
        'Day_' + RIGHT('00' + CAST(@CurrentDay AS VARCHAR(2)), 2);
    
    -- Columns for the Final Output (Aggregated Sums)
    SET @ColumnList = @ColumnList + 
        CASE WHEN @CurrentDay > 1 THEN ',' + CHAR(13) + CHAR(10) + '    ' ELSE '' END +
        'SUM(ISNULL([Day_' + RIGHT('00' + CAST(@CurrentDay AS VARCHAR(2)), 2) + '], 0)) AS [Day ' + 
        CAST(@CurrentDay AS VARCHAR(2)) + ']';

    -- Columns for the Temporary Table (Fixed Syntax)
    SET @TableCols = @TableCols + 
        CASE WHEN @CurrentDay > 1 THEN ', ' ELSE '' END +
        '[Day_' + RIGHT('00' + CAST(@CurrentDay AS VARCHAR(2)), 2) + '] DECIMAL(18,2) NULL';
    
    SET @CurrentDay = @CurrentDay + 1;
END;

-- Construct the dynamic SQL
SET @SQL = N'
-- Define a temporary table to hold the detailed mapping data
CREATE TABLE #TempReport (
    PROJECT NVARCHAR(50),
    SEW NVARCHAR(50),
    ASSY NVARCHAR(50),
    ProdLine NVARCHAR(50),
    ' + @TableCols + '
);

;WITH MonthlyReceipts AS (
    SELECT
        tr.[tr_part] AS ItemNumber,
        ''Day_'' + RIGHT(''00'' + CAST(DAY(tr.[tr_effdate]) AS VARCHAR(2)), 2) AS DayColumn,
        CAST(tr.[tr_qty_loc] AS DECIMAL(18,2)) AS Quantity
    FROM [dbo].[tr_hist] tr
    WHERE tr.[tr_type] = ''rct-wo''
        AND tr.[tr_effdate] >= @StartOfMonth
        AND tr.[tr_effdate] <= @EndOfMonth
        AND tr.[tr_qty_loc] > 0
),
AggregatedReceipts AS (
    SELECT
        ItemNumber,
        DayColumn,
        SUM(Quantity) AS TotalQuantity
    FROM MonthlyReceipts
    GROUP BY ItemNumber, DayColumn
),
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
ProjectMappedData AS (
    SELECT
        -- Mapping: PROJECT
        CASE 
            WHEN pt.[pt_prod_line] IN (''B_FG'', ''F_FG'', ''P_FG'') THEN ''LUCENEC''
            WHEN pt.[pt_prod_line] = ''H_FG'' THEN ''BJA''
            WHEN pt.[pt_prod_line] IN (''J_FG'', ''L_FG'') THEN ''JLR''
            WHEN pt.[pt_prod_line] IN (''C_FG'', ''Z_FG'', ''G_FG'') THEN ''CDPO''
            WHEN pt.[pt_prod_line] IN (''V_FG'', ''A_FG'') THEN ''VOLVO''
            WHEN pt.[pt_prod_line] IN (''K_FG'', ''Q_FG'') THEN ''KIA''
            WHEN pt.[pt_prod_line] IN (''M_FG'', ''N_FG'') THEN ''MMA''
            WHEN pt.[pt_prod_line] IN (''O_FG'', ''S_FG'') THEN ''OPEL''
            WHEN pt.[pt_prod_line] IN (''R_FG'', ''E_FG'', ''U_FG'') THEN ''CV''
            WHEN pt.[pt_prod_line] IN ( ''E_FG'',''T_FG'') THEN ''NISSAN''
            ELSE ''Other''
        END AS [PROJECT],

        -- Mapping: SEW
        CASE 
            WHEN pt.[pt_prod_line] = ''A_FG'' THEN ''Volvo - SEW''
            WHEN pt.[pt_prod_line] = ''B_FG'' THEN ''BR223''
            WHEN pt.[pt_prod_line] = ''F_FG'' THEN ''FIAT''
            WHEN pt.[pt_prod_line] = ''J_FG'' THEN ''JLR - SEW''
            WHEN pt.[pt_prod_line] = ''N_FG'' THEN ''MMA - SEW''
            WHEN pt.[pt_prod_line] = ''P_FG'' THEN ''PO426''
            WHEN pt.[pt_prod_line] = ''Q_FG'' THEN ''KIA - SEW''
            WHEN pt.[pt_prod_line] = ''S_FG'' THEN ''Opel - SEW''
            WHEN pt.[pt_prod_line] = ''Z_FG'' THEN ''CDPO - SEW''
            WHEN pt.[pt_prod_line] = ''H_FG'' THEN ''BJA''
            WHEN pt.[pt_prod_line] = ''R_FG'' THEN ''RENAULT''
            WHEN pt.[pt_prod_line] = ''E_FG'' THEN ''SCANIA''
            WHEN pt.[pt_prod_line] = ''U_FG'' THEN ''MAN''
            WHEN pt.[pt_prod_line] = ''T_FG'' THEN ''P13A''
            WHEN pt.[pt_prod_line] = ''G_FG'' THEN ''PZ1D''
            ELSE NULL
        END AS [SEW],

        -- Mapping: ASSY
        CASE 
            WHEN pt.[pt_prod_line] = ''V_FG'' THEN ''Volvo - ASSY''
            WHEN pt.[pt_prod_line] = ''L_FG'' THEN ''JLR - ASSY''
            WHEN pt.[pt_prod_line] = ''K_FG'' THEN ''KIA - ASSY''
            WHEN pt.[pt_prod_line] = ''M_FG'' THEN ''MMA - ASSY''
            WHEN pt.[pt_prod_line] = ''C_FG'' THEN ''CDPO - ASSY''
            WHEN pt.[pt_prod_line] = ''O_FG'' THEN ''Opel - ASSY''
            ELSE NULL
        END AS [ASSY],

        pt.[pt_prod_line] AS [ProdLine],
        ' + @PivotColumns + '
    FROM PivotedReceipts pvt
    LEFT JOIN [dbo].[pt_mstr] pt
        ON pvt.ItemNumber = pt.[pt_part]
    WHERE pt.[pt_prod_line] IS NOT NULL
)
-- Insert all mapped data into the temp table
INSERT INTO #TempReport
SELECT * FROM ProjectMappedData;

-- =============================================================================
-- RESULT SET 1: GROUPED BY PROJECT
-- =============================================================================
SELECT
    [PROJECT],
    ' + @ColumnList + '
FROM #TempReport
GROUP BY
    [PROJECT]
ORDER BY
    [PROJECT];

-- =============================================================================
-- RESULT SET 2: GROUPED BY SEW
-- =============================================================================
SELECT
    [SEW],
    ' + @ColumnList + '
FROM #TempReport
WHERE [SEW] IS NOT NULL
GROUP BY
    [SEW]
ORDER BY
    [SEW];

-- =============================================================================
-- RESULT SET 3: GROUPED BY ASSY
-- =============================================================================
SELECT
    [ASSY],
    ' + @ColumnList + '
FROM #TempReport
WHERE [ASSY] IS NOT NULL
GROUP BY
    [ASSY]
ORDER BY
    [ASSY];

-- Clean up
DROP TABLE #TempReport;
';

-- Execute the dynamic SQL
EXEC sp_executesql @SQL, 
    N'@StartOfMonth DATE, @EndOfMonth DATE', 
    @StartOfMonth = @StartOfMonth, 
    @EndOfMonth = @EndOfMonth;

PRINT 'ALL REPORTS EXECUTION COMPLETED';