WITH MainData AS (
    SELECT 
        ser.[Pack Code],
        ser.[Location],
        ser.[Item Number],
        ser.[Standard Pack Qty],
        ser.Qty,
        ser.count_ser_serial_id,
        ser.[Picked],
        ser.[standard pack check],
        pt.[Item Type],
        pt.[Description],
        pt.[Prod Line],
        pt.[Group],
        pt.[FG Type],
        pt.[Item Status],
        pt.[Safety Stock],
        pt.[Project],
        pt.[Planner],
        pt.[Supplier/Customer],
        pt.[Routing],
        pt.[Net weight in KG],
        pt.[New Item],
        loc.[Area],
        loc.[Zone],
        kit.[Kit Number],
        kit.[Model]
    FROM 
        (
            SELECT 
                [ser_pack_code] AS [Pack Code],
                [ser_loc] AS [Location],
                [ser_part] AS [Item Number],
                [ser_qty_pck] AS [Standard Pack Qty],
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
                [ser_qty_pck]
        ) AS ser
    INNER JOIN 
        (
            SELECT 
                [pt__chr02] AS [Item Type],
                [pt_part] AS [Item Number],
                [pt_desc1] AS [Description],
                [pt_prod_line] AS [Prod Line],
                [pt_group] AS [Group],
                [pt_draw] as [FG Type],
                [pt_status] AS [Item Status],
                [pt_sfty_stk] AS [Safety Stock],
                [pt_dsgn_grp] AS [Project],
                [pt_buyer] AS [Planner],
                [pt_vend] AS [Supplier/Customer],
                [pt_routing] AS [Routing],
                CASE 
                    WHEN [pt_net_wt_um] = 'kg' THEN [pt_net_wt]
                    WHEN [pt_net_wt_um] = 'g' THEN [pt_net_wt] / 1000.0
                    ELSE NULL  
                END AS [Net weight in KG],
                CASE 
                    WHEN DATEDIFF(day, [pt_added], GETDATE()) < 90 THEN 'New Item'
                    ELSE ''
                END AS [New Item]
            FROM 
                [QADEE2798].[dbo].[pt_mstr]
        ) AS pt
        ON ser.[Item Number] COLLATE SQL_Latin1_General_CP1_CI_AS = pt.[Item Number] COLLATE SQL_Latin1_General_CP1_CI_AS
    INNER JOIN 
        (
            SELECT 
                [xxwezoned_area_id] AS [Area],
                [xxwezoned_zone_id] AS [Zone],
                [xxwezoned_loc] AS [Location]
            FROM 
                [QADEE2798].[dbo].[xxwezoned_det]
        ) AS loc
        ON ser.[Location] COLLATE SQL_Latin1_General_CP1_CI_AS = loc.[Location] COLLATE SQL_Latin1_General_CP1_CI_AS
    LEFT JOIN 
        (
            SELECT
                d.[dtbin_ifs] as [Item Number],
                d.[dtbin_pi] as [Kit Number],
                dm.[data_model] as [Model]
            FROM 
                [TCIS_MES2_265].[dbo].[dtset_mstr] m
            INNER JOIN 
                [TCIS_MES2_265].[dbo].[dtbin_det] d
                ON m.[dtset_nbr] COLLATE SQL_Latin1_General_CP1_CI_AS = d.[dtbin_set_nbr] COLLATE SQL_Latin1_General_CP1_CI_AS
            INNER JOIN 
                [TCIS_MES2_265].[dbo].[data_mstr] dm
                ON m.[dtset_nbr] COLLATE SQL_Latin1_General_CP1_CI_AS = dm.[data_set] COLLATE SQL_Latin1_General_CP1_CI_AS 
                AND d.[dtbin_ifs] COLLATE SQL_Latin1_General_CP1_CI_AS = dm.[data_ifs] COLLATE SQL_Latin1_General_CP1_CI_AS
            WHERE 
                m.[dtset_last] = 1
        ) AS kit
        ON ser.[Item Number] COLLATE SQL_Latin1_General_CP1_CI_AS = kit.[Item Number] COLLATE SQL_Latin1_General_CP1_CI_AS
),
PODData AS (
    SELECT 
        pod.[pod_part],
        po.[po_vend]
    FROM 
        [QADEE2798].[dbo].[pod_det] AS pod
    LEFT JOIN 
        [QADEE2798].[dbo].[po_mstr] AS po ON pod.[pod_nbr] = po.[po_nbr]
    WHERE pod.[pod_end_eff[1]]] = '2049-12-31'  -- Fixed syntax error here
)
SELECT 
    MainData.*,
    PODData.[po_vend],
    CASE
        WHEN PODData.[po_vend] = '2848' AND MainData.[Prod Line] = 'V_RM' THEN 1.0/6.0
        WHEN PODData.[po_vend] = '2848' AND MainData.[Prod Line] <> 'V_RM' THEN 1.0
        WHEN MainData.[Prod Line] IN ('A_FG','F_FG','J_FG','N_FG','S_FG','P_FG','Q_FG','Z_FG') THEN 1.0/6.0
        -- Fixed condition to handle NULL FG Type properly
        WHEN MainData.[Prod Line] IN ('C_FG','G_FG','E_FG','L_FG','O_FG','R_FG','T_FG','U_FG','M_FG','V_FG','K_FG') 
             AND (MainData.[FG Type] <> 'Other' OR MainData.[FG Type] IS NULL) 
        THEN 1.0
        WHEN MainData.[Prod Line] = 'B_FG' AND MainData.[FG Type] = 'Center' THEN 1.0/12.0
        -- Simplified condition to handle all other B_FG cases including NULL
        WHEN MainData.[Prod Line] = 'B_FG' THEN 1.0/6.0
        ELSE NULL
    END AS [Warehouse_Packaging_Ratio],
    CASE
        WHEN PODData.[po_vend] = '2848' AND MainData.[Prod Line] = 'V_RM' THEN 2
        WHEN PODData.[po_vend] = '2848' AND MainData.[Prod Line] <> 'V_RM' THEN 5
        WHEN MainData.[Prod Line] IN ('A_FG','F_FG','J_FG','N_FG','S_FG','K_FG','P_FG','Q_FG','Z_FG','B_FG') THEN 2
        WHEN MainData.[Prod Line] IN ('G_FG','T_FG') THEN 5
        WHEN MainData.[Prod Line] IN ('E_FG','R_FG','U_FG') AND MainData.[FG Type] = 'FB' THEN 4
        WHEN MainData.[Prod Line] IN ('E_FG','R_FG','U_FG') AND MainData.[FG Type] = 'FC' THEN 5
		WHEN MainData.[Prod Line] IN ('E_FG','R_FG','U_FG') AND MainData.[FG Type] = 'other' THEN null
        WHEN MainData.[Prod Line] = 'V_FG' THEN 3
        WHEN MainData.[Prod Line] IN ('L_FG','K_FG','C_FG','O_FG','M_FG') THEN 4
        ELSE NULL
    END AS [Stackability],
    CASE
        WHEN PODData.[po_vend] = '2848' THEN 'Poiana'
        WHEN MainData.[Prod Line] = '0000' THEN 'check prod line'
        WHEN MainData.[Prod Line] = 'A_FG' THEN 'SFG'
        WHEN MainData.[Prod Line] = 'B_FG' THEN 'Lucenec'
        WHEN MainData.[Prod Line] = 'C_FG' THEN 'CDPO'
        WHEN MainData.[Prod Line] = 'E_FG' THEN 'CV'
        WHEN MainData.[Prod Line] = 'F_FG' THEN 'Lucenec'
        WHEN MainData.[Prod Line] = 'G_FG' THEN 'Nissan'
        WHEN MainData.[Prod Line] = 'J_FG' THEN 'SFG'
        WHEN MainData.[Prod Line] = 'K_FG' THEN 'KIA'
        WHEN MainData.[Prod Line] = 'L_FG' THEN 'LAND ROVER'
        WHEN MainData.[Prod Line] = 'M_FG' THEN 'MMA'
        WHEN MainData.[Prod Line] = 'N_FG' THEN 'SFG'
        WHEN MainData.[Prod Line] = 'O_FG' THEN 'Opel'
        WHEN MainData.[Prod Line] = 'P_FG' THEN 'Lucenec'
        WHEN MainData.[Prod Line] = 'Q_FG' THEN 'SFG'
        WHEN MainData.[Prod Line] = 'R_FG' THEN 'CV'
        WHEN MainData.[Prod Line] = 'S_FG' THEN 'SFG'
        WHEN MainData.[Prod Line] = 'T_FG' THEN 'Nissan'
        WHEN MainData.[Prod Line] = 'U_FG' and MainData.[FG Type] <> 'other' THEN 'CV'
        WHEN MainData.[Prod Line] = 'V_FG' THEN 'Volvo'
        WHEN MainData.[Prod Line] = 'Z_FG' THEN 'SFG'
        ELSE 'other'
    END AS [WH_Group]
FROM 
    MainData
LEFT JOIN 
    PODData 
    ON MainData.[Item Number] COLLATE SQL_Latin1_General_CP1_CI_AS = PODData.[pod_part] COLLATE SQL_Latin1_General_CP1_CI_AS
ORDER BY 
    MainData.[Area], 
    MainData.[Zone];