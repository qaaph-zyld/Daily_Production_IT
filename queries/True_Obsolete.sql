WITH LastDates AS (
    SELECT 
        tr_part,
        MAX(CASE WHEN tr_type = 'iss-wo' THEN tr_effdate END) AS LastIssueDate,
        MAX(CASE WHEN tr_type = 'iss-so' THEN tr_effdate END) AS LastSaleDate,
        MAX(CASE WHEN tr_type = 'rct-wo' THEN tr_effdate END) AS LastProdDate,
        MAX(CASE WHEN tr_type = 'rct-po' THEN tr_effdate END) AS LastReceiptDate
    FROM 
        [QADEE2798].[dbo].[tr_hist]
    WHERE 
        tr_type IN ('iss-wo','rct-po','iss-so','rct-wo')
        AND tr_effdate < DATEADD(day, -365, GETDATE())
        AND tr_userid <> 'ajelacn'
    GROUP BY 
        tr_part
),
MergedDates AS (
    SELECT 
        tr_part,
        -- Calculate merged issue date (most recent between iss-wo and iss-so)
        CASE 
            WHEN LastIssueDate IS NULL AND LastSaleDate IS NULL THEN NULL
            WHEN LastIssueDate IS NULL THEN LastSaleDate
            WHEN LastSaleDate IS NULL THEN LastIssueDate
            WHEN LastIssueDate > LastSaleDate THEN LastIssueDate
            ELSE LastSaleDate
        END AS MergedIssueDate,
        
        -- Calculate merged receipt date (most recent between rct-wo and rct-po)
        CASE 
            WHEN LastProdDate IS NULL AND LastReceiptDate IS NULL THEN NULL
            WHEN LastProdDate IS NULL THEN LastReceiptDate
            WHEN LastReceiptDate IS NULL THEN LastProdDate
            WHEN LastProdDate > LastReceiptDate THEN LastProdDate
            ELSE LastReceiptDate
        END AS MergedReceiptDate
    FROM 
        LastDates
)
SELECT 
    tr_part as [Item Number],
    -- Categorize merged issue date
    CASE 
        WHEN MergedIssueDate IS NULL THEN 'No transactions'
        WHEN DATEDIFF(day, MergedIssueDate, GETDATE()) < 90 THEN 'active'
        WHEN DATEDIFF(day, MergedIssueDate, GETDATE()) BETWEEN 91 AND 180 THEN '3 months'
        WHEN DATEDIFF(day, MergedIssueDate, GETDATE()) BETWEEN 181 AND 365 THEN '6 months'
        ELSE 'obsolete'
    END AS [Last Issue],
    
    -- Categorize merged receipt date
    CASE 
        WHEN MergedReceiptDate IS NULL THEN 'No transactions'
        WHEN DATEDIFF(day, MergedReceiptDate, GETDATE()) < 90 THEN 'active'
        WHEN DATEDIFF(day, MergedReceiptDate, GETDATE()) BETWEEN 91 AND 180 THEN '3 months'
        WHEN DATEDIFF(day, MergedReceiptDate, GETDATE()) BETWEEN 181 AND 365 THEN '6 months'
        ELSE 'obsolete'
    END AS [Last Receipt]
FROM 
    MergedDates