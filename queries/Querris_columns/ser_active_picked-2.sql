SELECT 
    [ser_pack_code] as [Pack Code],
    [ser_loc] as [Location],
    [ser_part] as [Item Number],
    [ser_qty_pck] as [Standard Pack Qty],
    SUM([ser_qty_avail]) AS Qty,
    COUNT([ser_serial_id]) AS count_ser_serial_id,
    SUM(CASE WHEN ser_stage = 'picked' THEN ser_qty_avail ELSE 0 END) AS [Picked],
    CASE 
        WHEN ser_qty_pck = (SUM([ser_qty_avail]) * 1.0 / COUNT([ser_serial_id])) 
        THEN 'OK' 
        ELSE 'check standard pack' 
    END AS [standard pack check]
FROM 
    [QADEE2798].[dbo].[ser_active_picked]
GROUP BY 
    [ser_part],
    [ser_loc],
    [ser_stage],
    [ser_pack_code],
    [ser_qty_pck];