-- CTE 1: Get the base list of serial IDs that meet the initial criteria.
WITH BaseSerials AS (
    SELECT
        serh_serial_id
    FROM
        [QADEE2798].[dbo].[serh_hist]
    WHERE
        serh_trans_type = 'pck-bld'
        AND serh_loc = 'prod'
        AND serh_trans_date >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)
        AND serh_trans_date < DATEADD(MONTH, 1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
),

-- CTE 2: Get the detailed list of each serial ID, its item number, quantity, and status.
DetailedData AS (
    SELECT
        h.serh_serial_id as Serial_ID,
        MAX(h.serh_part) AS Item_Number,
        SUM(CASE WHEN h.serh_trans_type = 'pck-bld' THEN h.serh_qty_chg ELSE 0 END) AS qty_per_boxes,
        CASE
            WHEN SUM(CASE WHEN h.serh_trans_type = 'pck-iss' THEN h.serh_qty_chg ELSE 0 END) = 0 THEN 'active'
            ELSE 'sold'
        END AS [active/sold]
    FROM
        [QADEE2798].[dbo].[serh_hist] h
    INNER JOIN
        BaseSerials b ON h.serh_serial_id = b.serh_serial_id
    WHERE
        h.serh_trans_type IN ('pck-bld', 'pck-dec', 'pck-iss')
    GROUP BY
        h.serh_serial_id
    HAVING
        SUM(CASE WHEN h.serh_trans_type = 'pck-dec' THEN h.serh_qty_chg ELSE 0 END) = 0
),

-- CTE 3: Aggregate the detailed serial data by Item_Number
AggregatedSerials AS (
    SELECT
        Item_Number,
        SUM(qty_per_boxes) AS qty_per_boxes,
        COUNT(Serial_ID) AS SerialID_Count,
        SUM(CASE WHEN [active/sold] = 'active' THEN qty_per_boxes ELSE 0 END) AS active_qty,
        SUM(CASE WHEN [active/sold] = 'sold' THEN qty_per_boxes ELSE 0 END) AS sold_qty
    FROM
        DetailedData
    GROUP BY
        Item_Number
),

-- CTE 4: Get work order transaction data (iss-wo and rct-wo) - CORRECTED
WorkOrderData AS (
    SELECT
        [tr_part] AS [Item_Number],
        SUM(CASE WHEN [tr_type] = 'iss-wo' THEN [tr_qty_loc] ELSE 0 END) AS [ISS_WO_Qty],
        SUM(CASE WHEN [tr_type] = 'rct-wo' THEN [tr_qty_loc] ELSE 0 END) AS [RCT_WO_Qty],
        -- Use MAX or MIN to consolidate prod_line if multiple exist per item
        MAX([tr_prod_line]) AS [tr_prod_line]
    FROM
        [QADEE2798].[dbo].[tr_hist]
    WHERE
        [tr_type] IN ('iss-wo', 'rct-wo')
        -- Current month filter - VERIFIED CORRECT
        AND [tr_effdate] >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)
        AND [tr_effdate] < DATEADD(MONTH, 1, DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1))
    GROUP BY
        [tr_part]
    -- Remove the HAVING clause if you want ALL items, even with 0 quantities
),

-- CTE 5: Join the data and create the RM/SFG/FG column
PreFinalData AS (
    SELECT
        a.Item_Number,
        w.[tr_prod_line],
        CASE
            WHEN w.tr_prod_line IN ('A_FG', 'Q_FG', 'J_Fg', 'N_Fg', 's_Fg', 'z_fg') THEN 'SFG'
            WHEN w.tr_prod_line LIKE '%_FG%' THEN 'FG'
            WHEN w.tr_prod_line LIKE '%RM%' THEN 'RM'
            ELSE 'UNKNOWN'
        END AS [RM/SFG/FG],
        a.qty_per_boxes,
        a.SerialID_Count,
        a.active_qty,
        a.sold_qty,
        ISNULL(w.ISS_WO_Qty, 0) AS ISS_WO_Qty,
        ISNULL(w.RCT_WO_Qty, 0) AS RCT_WO_Qty
    FROM
        AggregatedSerials a
    LEFT JOIN
        WorkOrderData w ON a.Item_Number = w.Item_Number
)

-- Final Step: Select and apply conditional delta calculation
SELECT
    Item_Number,
    [tr_prod_line] AS [Prod_Line],
    [RM/SFG/FG],
    qty_per_boxes,
    SerialID_Count,
    active_qty,
    sold_qty,
    ISS_WO_Qty,
    RCT_WO_Qty,
    -- Conditional logic for delta_bkfl
    CASE
        WHEN [RM/SFG/FG] = 'SFG' THEN qty_per_boxes - (RCT_WO_Qty + ISS_WO_Qty)
        ELSE qty_per_boxes - RCT_WO_Qty
    END AS delta_bkfl
FROM
    PreFinalData
ORDER BY
    Item_Number;