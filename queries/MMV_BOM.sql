WITH ParentItems AS (
    SELECT DISTINCT
        tr_part AS parent_item,
        tr_lot,
        tr_site
    FROM [QADEE2798].[dbo].[tr_hist]
    WHERE tr_type = 'RCT-WO'
    AND tr_effdate >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)
    AND tr_effdate < DATEADD(month, 1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
),
BaseData AS (
    SELECT
        MONTH(CONVERT(date, tr_effdate)) AS [Month],
        h.tr_lot,
        h.tr_part,
        h.tr_type,
        h.tr_prod_line,
        h.tr_mtl_std,
        h.tr_qty_loc,
        h.tr_mtl_std * h.tr_qty_loc AS amount,
        h.tr_site
    FROM [QADEE2798].[dbo].[tr_hist] h
    WHERE tr_type IN ('RCT-WO', 'ISS-WO')
    AND tr_effdate >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)
    AND tr_effdate < DATEADD(month, 1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
),
PS_Mstr_Combined AS (
    SELECT '2798' AS [Plant], [ps_par], [ps_comp], [ps_qty_per], [ps_op]
    FROM [QADEE2798].[dbo].[ps_mstr] WHERE [ps_end] IS NULL
),
AggregatedData AS (
    SELECT
        b.[Month],
        b.tr_prod_line AS [Prod Line],
        p.parent_item AS [Parent Item],
        b.tr_lot AS [Batch ID],
        b.tr_type AS [Transaction Type],
        b.tr_part AS [Item Number],
        b.tr_mtl_std AS [Material Cost],
        ps.ps_op AS [Operation],
        ps.ps_qty_per,
        SUM(-b.tr_qty_loc) AS total_qty_change,
        SUM(-b.amount) AS total_amount
    FROM BaseData b
    LEFT JOIN ParentItems p ON b.tr_lot = p.tr_lot
    LEFT JOIN PS_Mstr_Combined ps 
        ON p.parent_item = ps.ps_par 
        AND b.tr_part = ps.ps_comp 
        AND b.tr_site = ps.Plant
    GROUP BY 
        b.[Month],
        b.tr_prod_line,
        p.parent_item,
        b.tr_lot,
        b.tr_type,
        b.tr_part,
        b.tr_mtl_std,
        ps.ps_op,
        ps.ps_qty_per
),
WithRCTWO AS (
    SELECT *,
        SUM(CASE WHEN [Transaction Type] = 'RCT-WO' THEN total_qty_change ELSE 0 END) 
            OVER (PARTITION BY [Batch ID]) AS RCT_WO_Qty
    FROM AggregatedData
),
WithQtyBOM AS (
    SELECT *,
        CASE 
            WHEN ps_qty_per IS NOT NULL THEN RCT_WO_Qty * ps_qty_per
            ELSE -1 * RCT_WO_Qty
        END AS [Qty per BOM]
    FROM WithRCTWO
),
WithDeltaColumns AS (
    SELECT *,
        total_qty_change + [Qty per BOM] AS [delta BKFL Qty],
        (total_qty_change + [Qty per BOM]) * [Material Cost] AS [delta BKFL Amount]
    FROM WithQtyBOM
),
BatchChecks AS (
    SELECT *,
        SUM([delta BKFL Qty]) OVER (PARTITION BY [Batch ID]) AS BatchTotalDelta
    FROM WithDeltaColumns
),
CostData AS (
    SELECT 
        [sct_part] as [Item Number],
        [sct_cst_tot] as [Standard Cost],
        ([sct_mtl_tl] + [sct_mtl_ll]) AS [CMAT],
        ([sct_cst_tot] - ([sct_mtl_tl] + [sct_mtl_ll])) AS [LBO],
        CASE 
            WHEN [sct_cst_tot] = 0 THEN ''
            WHEN [sct_cst_tot] = ([sct_mtl_tl] + [sct_mtl_ll]) THEN 'RM'
            ELSE 'SFG/FG'
        END AS [Prod/Mfg]
    FROM [QADEE2798].[dbo].[sct_det]
    WHERE [sct_sim] = 'standard'
)
SELECT
    bc.[Month],
    bc.[Prod Line],
    bc.[Parent Item],
    bc.[Batch ID],
    bc.[Transaction Type],
    bc.[Item Number],
    bc.[Material Cost],
    bc.[Operation],
    bc.ps_qty_per,
    bc.total_qty_change,
    bc.total_amount,
    bc.[Qty per BOM],
    bc.[delta BKFL Qty],
    bc.[delta BKFL Amount],
    bc.BatchTotalDelta,
    CASE 
        WHEN bc.BatchTotalDelta = 0 THEN 'OK'
        ELSE 'NOK'
    END AS [Batch ID check],
    cd.[Standard Cost],
    cd.[CMAT],
    cd.[LBO],
    cd.[Prod/Mfg],
    cd.CMAT * bc.total_qty_change AS Total_Amount_corr
FROM BatchChecks bc
LEFT JOIN CostData cd ON bc.[Item Number] = cd.[Item Number]
ORDER BY
    bc.[Month],
    bc.[Parent Item],
    bc.[Batch ID];