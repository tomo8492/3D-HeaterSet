import xception as E

def sql_dict(lobj: "module.LotObject") -> str:
    '''
    this function assemable sql string parts to a whole sql query string,
    Parameters:
    take an object with below named attributes as arg:
        lotid -> search for key :string
        operator -> iterable with 3 values(select column key, from db key, where condition key) :iterable
        attributes -> select column key iterable:iterable
        cond -> Order condition(ASEC, DEC):string
    return string
    '''
    lotid: str = lobj.lotid
    ope = lobj.operator
    select = ope[0]
    fromdb = ope[1]
    where = ope[2]
    attributes = lobj.attributes
    cond = lobj.cond
    plant = lobj.plant
    datas = lobj.datas

    prefix = 'SELECT '
    if select == 'RelatedID_Tree':
        prefix = f'''
                            WITH target_ancestors AS (
                -- Case 1: Input node exists as child (子指図) - find actual root
                SELECT 
                    relation.子指図 as node_id,
                    relation.子品目 as node_sap,
                    relation.親指図 as parent_id,
                    relation.親品目 as parent_sap,
                    LEVEL as hierarchy_level,
                    relation.子指図 as target_root
                FROM V_指図親子関係一覧 relation
                WHERE CONNECT_BY_ISLEAF = 1  -- Only leaf nodes in upward traversal (roots)
                START WITH relation.子指図 IN ({lotid})
                CONNECT BY NOCYCLE relation.子指図 = PRIOR relation.親指図
                
                UNION ALL
                
                -- Case 2: Input node exists as parent (親指図) - find actual root
                SELECT 
                    relation.子指図 as node_id,
                    relation.子品目 as node_sap,
                    relation.親指図 as parent_id,
                    relation.親品目 as parent_sap,
                    LEVEL as hierarchy_level,
                    relation.子指図 as target_root
                FROM V_指図親子関係一覧 relation
                WHERE relation.親指図 IS NULL  -- Root nodes only
                START WITH relation.親指図 IN ({lotid})
                CONNECT BY NOCYCLE PRIOR relation.子指図 = relation.親指図
                
                UNION ALL
                
                -- Case 3: Input is root itself
                SELECT 
                    relation.子指図 as node_id,
                    relation.子品目 as node_sap,
                    relation.親指図 as parent_id,
                    relation.親品目 as parent_sap,
                    1 as hierarchy_level,
                    relation.子指図 as target_root
                FROM V_指図親子関係一覧 relation
                WHERE relation.子指図 IN ({lotid}) AND relation.親指図 IS NULL
            ),
            filtered_hierarchy AS (
                -- Traverse downward from actual root
                SELECT
                    relation.子指図 as node_id,
                    relation.子品目 as node_sap,
                    relation.親指図 as parent_id,
                    relation.親品目 as parent_sap,
                    LEVEL as hierarchy_level,
                    SYS_CONNECT_BY_PATH(relation.子指図, ' -> ') as hierarchy_path,
                    CONNECT_BY_ROOT relation.子指図 as root_id
                FROM V_指図親子関係一覧 relation
                START WITH relation.親指図 IS NULL 
                    AND relation.子指図 IN (SELECT DISTINCT target_root FROM target_ancestors)
                CONNECT BY NOCYCLE PRIOR relation.子指図 = relation.親指図
            )
        '''

    if select == 'ShelfUpdating':
        prefix = "UPDATE "
    
    if select == "HistoryInsert":
        prefix = "INSERT"

    dict = {"ProcessCheck":{
    "UserCode":"MMEC.ITEM_NAME",
    "LotID":"MLI.LOT_ID",
    "ProductCode":"NMRP.OLD_PRODUCT_CODE",
    "SAPCode":"NMRP.PRODUCT_CODE",
    "CurrentProcess":"MMPL.PROCESS_NAME",
    "Status":"CASE WHEN MLI.HOLD_FLG = '1' THEN '保留' ELSE '-' END", 
    "LotState": "CASE WHEN MLI.STATE_FLG = 'E' THEN '完成' WHEN MLI.STATE_FLG = 'S' THEN '廃棄' ELSE '仕掛' END",
    "LastUpdatedDate":"MLI.UP_DATE",
    #"Location":"TO_SINGLE_BYTE(MLI.LOCATION_CODE)",
    "SerialNumber":"TO_SINGLE_BYTE(ILS.SERIAL_01), TO_SINGLE_BYTE(ILS.SERIAL_02), TO_SINGLE_BYTE(ILS.SERIAL_03), TO_SINGLE_BYTE(ILS.SERIAL_04), TO_SINGLE_BYTE(ILS.SERIAL_05), TO_SINGLE_BYTE(ILS.SERIAL_06)",
    "InProgress":"CASE MLI.STEP WHEN '100' THEN '処理前' WHEN '301' THEN '処理中' WHEN '302' THEN '処理後' END ",
    "OnHoldReason":"COALESCE(MMRL.COMMENTS, '-')",
    "Restriction":"""COALESCE(
    (SELECT CASE WHEN Z.RELEASE_FLG = 0 THEN Z.NOT_SHIPPABLE_TYPE || ' : ' || Z.NOT_SHIPPABLE_REASON ELSE '制限解除済み' END 
    FROM Z_KIBUKURO_INFORMATION Z 
    WHERE MLI.LOT_ID IN (Z.LOT_ID, Z.LOT_ID1, Z.LOT_ID2)
    AND Z.RELEASE_FLG = (SELECT MIN(Z2.RELEASE_FLG) 
                        FROM Z_KIBUKURO_INFORMATION Z2 
                        WHERE MLI.LOT_ID IN (Z2.LOT_ID, Z2.LOT_ID1, Z2.LOT_ID2))
    AND ROWNUM = 1), '-')
    """,#"(SELECT CASE WHEN MIN(Z.RELEASE_FLG) = 0 THEN Z.NOT_SHIPPABLE_TYPE || ' : ' || Z.NOT_SHIPPABLE_REASON ELSE '-' END FROM Z_KIBUKURO_INFORMATION Z WHERE TO_SINGLE_BYTE(MLI.LOT_ID) IN (Z.LOT_ID, Z.LOT_ID1, Z.LOT_ID2))",
    "flowName" : "MMFL.FLOW_NAME"
    },

    "RelatedID":{
    "ParentID":"V_指図親子関係一覧.親指図",
    "ParentSAP":"V_指図親子関係一覧.親品目",
    "ChildID":"V_指図親子関係一覧.子指図",
    "ChildSAP":"V_指図親子関係一覧.子品目",
    "DB":"DISTINCT Relation.子品目, NMRP_P.OLD_PRODUCT_CODE, Relation.親品目"
    },

    "RelatedID_Tree":{
    "NodeID": f"""
    SELECT 
    fh.node_id,
    (SELECT OLD_PRODUCT_CODE 
     FROM NHK_MASTER_RECEIVE_PRODUCT 
     WHERE PRODUCT_CODE = fh.parent_sap 
     AND ROWNUM = 1) as P_DB,
    fh.parent_id,
    (SELECT OLD_PRODUCT_CODE 
     FROM NHK_MASTER_RECEIVE_PRODUCT 
     WHERE PRODUCT_CODE = fh.node_sap 
     AND ROWNUM = 1) as C_DB,
    fh.hierarchy_level,
    fh.hierarchy_path,
    fh.root_id,
    CASE WHEN fh.node_id IN ({lotid}) THEN 'INPUT_NODE' ELSE 'RELATED_NODE' END as node_type
    """
    },

    "Schedule":{"ScStart":"V_ASP_生産計画一覧.計画日_着手",
    "ScFinish":"V_ASP_生産計画一覧.計画日_完了",
    "LotID":"V_ASP_生産計画一覧.指図NO",
    "SAPCode":"V_ASP_生産計画一覧.品目",
    "DBCode":"V_ASP_生産計画一覧.ＤＢ図番",
    "FlowCode":"V_ASP_生産計画一覧.フローCD",
    "FlowName":"V_ASP_生産計画一覧.フロー名称",
    "ManCode":"V_ASP_生産計画一覧.管理CD",
    "UserName":"V_ASP_生産計画一覧.客先名",
    "UserPartsNum":"V_ASP_生産計画一覧.客先品番",
    "Quantity":"V_ASP_生産計画一覧.数量",
    "Resource":"V_ASP_生産計画一覧.資源",
    "Cat1":"V_ASP_生産計画一覧.大区分",
    "Cat1Name":"V_ASP_生産計画一覧.大区分名称",
    "Cat2":"V_ASP_生産計画一覧.中区分",
    "Cat2Name":"V_ASP_生産計画一覧.中区分名称",
    "ProcessCode":"V_ASP_生産計画一覧.工程CD",
    "ProcessName":"V_ASP_生産計画一覧.工程名",
    "Start":"V_ASP_生産計画一覧.着手予定日",
    "Finish":"V_ASP_生産計画一覧.完了予定日",
    "ProcessOrder":"V_ASP_生産計画一覧.工順",
    "Status":"V_ASP_生産計画一覧.ステータス",
    "UpdateTime":"V_ASP_生産計画一覧.LV更新日時",
    "DeliveryDateByStaff":"V_ASP_生産計画一覧.担当者設定納期",
    "StandardOpearteTime":"V_ASP_生産計画一覧.標準手作業時間",
    "StandardMachiningTime":"V_ASP_生産計画一覧.標準装置時間",
    "StandardTime":"V_ASP_生産計画一覧.標準時間",
    "SUMStandardTime":"P.指図NO, MIN(W.工程名) KEEP (DENSE_RANK FIRST ORDER BY W.工程名) AS CurrentProcess, SUM(P.標準時間) AS RemainTime",
    "SUMStandardTimeRevised":"SUM(CASE WHEN V_ASP_生産計画一覧.工程名 NOT LIKE '%受入%' THEN V_ASP_生産計画一覧.標準時間 ELSE 0 END)",
    "SAPParentCode":"V_ASP_生産計画一覧.中間品目",
    "CurrentProcess!":"MMPL.PROCESS_NAME AS 現工程"},

    '''
    Add average lead time to each process code to enable the app to calculate the estimate complete date by sum up the
    lead time of each process averaged from process history, cache the process code and average lead time into a cache table to avoid
    putting great load onto DB server, this table should have a timestamp column as a flag to trigger refreshment of the average lead time
    after a set amount of time(for example once a year).
    SQL to get the average lead time:
    SELECT V_工程別指図リードタイム_完成.NHK図番, V_工程別指図リードタイム_完成.工程コード, V_工程別指図リードタイム_完成.工程名, Avg(V_工程別指図リードタイム_完成.リードタイム)
    FROM LV23_TEST.V_工程別指図リードタイム_完成 V_工程別指図リードタイム_完成
    WHERE (V_工程別指図リードタイム_完成.NHK図番='DB20083')
    GROUP BY V_工程別指図リードタイム_完成.NHK図番, V_工程別指図リードタイム_完成.工程コード, V_工程別指図リードタイム_完成.工程名
    '''

    "WIP":{"WIP":"V_指図別工程別仕掛.指図NO, V_指図別工程別仕掛.DB図番, COALESCE(V_指図別工程別仕掛.フロー名称, '-'), V_指図別工程別仕掛.工程名, V_指図別工程別仕掛.処理区分, V_指図別工程別仕掛.更新日, V_指図別工程別仕掛.SEQ, COALESCE(V_指図別工程別仕掛.保留, '-'), COALESCE(V_指図別工程別仕掛.保留理由, '-'), CASE WHEN V_指図別工程別仕掛.保留 IS NOT NULL THEN 525600 ELSE CEIL(SUM(CASE WHEN COALESCE(V_ASP_生産計画一覧.工程名, '-') NOT LIKE '%受入%' THEN COALESCE(V_ASP_生産計画一覧.標準時間, 0) ELSE 1440 END) ) END"},
    #V_指図別工程別仕掛.指図NO
    #'YYYY-MM-DD HH24:MI:SS'
    "WIP-I":{"WIP-I":"""
    MLI.LOT_ID,
    MMFL.FLOW_NAME,
    MMPL.PROCESS_NAME,
    CASE WHEN MLI.STEP = '100' THEN '処理前' WHEN MLI.STEP = '301' THEN '処理中' WHEN MLI.STEP = '302' THEN '処理後' ELSE MLI.STEP END,
    TO_CHAR(MLI.UP_DATE, 'YYYY-MM-DD HH24:MI:SS'),
    MMF.FLOW_SEQ_NO,
    COALESCE(MMRL.COMMENTS, '-'),
    CASE WHEN ENDD.COMPDATE IS NULL THEN '-' ELSE TO_CHAR(ENDD.COMPDATE, 'YYYY-MM-DD') END,
    CASE WHEN MLI.USER_CODE = 'RECEIVE_LOT' THEN '投入前' ELSE '-' END,
    COALESCE(
             (SELECT CASE WHEN Z.RELEASE_FLG = 0 THEN Z.NOT_SHIPPABLE_TYPE || ' : ' || Z.NOT_SHIPPABLE_REASON ELSE '制限解除済み' END 
    FROM Z_KIBUKURO_INFORMATION Z 
    WHERE MLI.LOT_ID IN (Z.LOT_ID, Z.LOT_ID1, Z.LOT_ID2)
    AND Z.RELEASE_FLG = (SELECT MIN(Z2.RELEASE_FLG) 
                        FROM Z_KIBUKURO_INFORMATION Z2 
                        WHERE MLI.LOT_ID IN (Z2.LOT_ID, Z2.LOT_ID1, Z2.LOT_ID2))
    AND ROWNUM = 1), '-'),
    NMRP.OLD_PRODUCT_CODE
    """
    },

    "Records":{
    "RecordsCluster":"""
    MLH.LOT_ID,
	MLH.FLOW_CODE,
	MMPL.PROCESS_NAME,
	CASE WHEN MLH.STEP = '100' THEN '処理前' WHEN MLH.STEP = '301' THEN '処理中' WHEN MLH.STEP = '302' THEN '処理後' ELSE MLH.STEP END,
	COALESCE(MLH.EQP_CODE, '-'),
	CASE WHEN MLH.HOLD_FLG = '1' THEN '保留' ELSE '-' END,
	COALESCE(MMRL.COMMENTS, '-'),
	CASE WHEN MLH.STATE_FLG = 'E' THEN '完成' WHEN MLH.STATE_FLG = 'S' THEN '廃棄' ELSE '仕掛' END,
	MLH.UP_DATE,
	MLH.USER_CODE,
    MMFL.FLOW_NAME
    """
    },

    "YellowBag":{
    "YellowBagCluster":"""
    LOT_ID,
	DB_NO,
	LOT_ID1,
	DB_NO1,
	LOT_ID2,
	DB_NO2,
	NOT_SHIPPABLE_TYPE,
	NOT_SHIPPABLE_REASON,
	CASE WHEN RELEASE_FLG = 1 THEN '解除済' ELSE '出荷制限中' END,
	COALESCE(RELEASE_BY, '-'),
	CASE WHEN RELEASE_DATE IS NULL THEN '-' ELSE RELEASE_DATE END,
	COALESCE(WRITTEN_BY, '-')
    """
    },

    "DBSearch" : {
        "DB" : "DISTINCT OLD_PRODUCT_CODE"
    },

    "Components" :{
        "DB" : "DISTINCT NMRP.OLD_PRODUCT_CODE AS SEARCH_KEY",
        "SAPCode" : "NMRP.PRODUCT_CODE AS KEY_SAP",
        "CompoName" : "PROSAP.OLD_PRODUCT_CODE AS COMPO_NAME",
        "ProcessName" : "NMMP.PROCESS_CODE AS PROCESS_NAME",
        "ComponentCode" : "NMMP.COMPONENT AS COMPONENT_CODE",
        "Count" : "NMMP.COUNT AS COUNT",
        "StrName" : "PROSAP.PRODUCT_NO, PROSAP.PRODUCT_NAME"
    },

    "Flow" : {
        "flowitem" : "V1.品目コード, V1.DB図番, V1.品名, MMF.FLOW_CODE AS フローコード, MMFL.FLOW_NAME AS フロー名称, CASE WHEN MMF.FLOW_CODE LIKE '%A' THEN '内製' WHEN MMF.FLOW_CODE LIKE '%B' THEN '外製' WHEN MMF.FLOW_CODE LIKE '%B_%' THEN '外製' WHEN MMF.FLOW_CODE LIKE '%C' THEN '3次以降外製' WHEN MMF.FLOW_CODE LIKE '%C_%' THEN '3次以降外製' ELSE 'その他' END AS 内外製区分, MMF.FLOW_SEQ_NO AS 工順, V1.工程コード, V1.工程名称, V1.使用設備, V1.手作業時間, V1.機械時間, V1.合計時間"
    },

    "ShelfExisting" : {
        "existing" : "COUNT(*)",
        "all_data" : "INFO.LOT_ID, INFO.SHELF_NO, MST.SHELF_NAME, INFO.USER_CODE, INFO.UP_DATE, INFO.URGENT_FLG, INFO.DESIGNATED_SHELF_FLAG, NMRP.OLD_PRODUCT_CODE, MMPL.PROCESS_NAME, CASE WHEN MLI.HOLD_FLG = '1' THEN '保留' ELSE '-' END"
    },

    "ShelfUpdating" : {"updating" : 
                       "Z_SHELF_LOT_INFORMATION SET SHELF_NO = :1, IN_OUT_FLG = :2, USER_CODE = :3, TRANSPORT_DATE = '', UP_DATE = SYSDATE"},

    "HistoryInsert" : {"history":
                       " INTO Z_SHELF_LOT_HISTORY VALUES(:1, :2, :3, :4, '', SYSDATE, '5501', :5)"},

    "DataViewerUni" : {"dataviewer": "*"},

    "ProcessOverview" : {"pov" : f""" 
            MLI.LOT_ID, 
            NMRP.OLD_PRODUCT_CODE, 
            MLH.FLOW_CODE, 
            MMPL.PROCESS_NAME, 
            MMRL.COMMENTS, 
            NULL AS SC_PROCESS_NAME,
            NULL AS SC_DATE,
            NULL AS LEADTIME,
            MLI.STATE_FLG, 
            MLH.UP_DATE,
            'HI' AS CAT 
        FROM LV23_TEST.MES_LOT_INFORMATION MLI
        INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP ON NMRP.PRODUCT_CODE = MLI.PRODUCT_CODE AND NMRP.PLANT = MLI.LOCATION_CODE AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE)
            IN (
            SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE)
            FROM NHK_MASTER_RECEIVE_PRODUCT
            GROUP BY PLANT, PRODUCT_CODE)
        INNER JOIN MES_LOT_HISTORY MLH ON MLI.LOT_ID = MLH.LOT_ID
        LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLH.PROCESS_CODE 
            AND MMPL.LOCATION_CODE = MLH.LOCATION_CODE
        LEFT JOIN MES_MASTER_REASON_LC MMRL ON MLH.REASON_CODE = MMRL.REASON_CODE
        WHERE MLI.STATE_FLG NOT IN ('E', 'S') 
            AND NMRP.OLD_PRODUCT_CODE IN ({lotid})

        UNION ALL

        SELECT 
            MLI.LOT_ID, 
            VA.ＤＢ図番, 
            NULL, 
            NULL, 
            NULL,
            VA.工程名,
            VA.着手予定日,
            1,
            MLI.STATE_FLG, 
            NULL,
            'SC'  
    """},

    "FromDB":{
    "ProcessRelated":"""FROM MES_LOT_INFORMATION MLI
    LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE AND MMPL.LOCATION_CODE = MLI.LOCATION_CODE
    INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP ON NMRP.PRODUCT_CODE = MLI.PRODUCT_CODE AND NMRP.PLANT = MLI.LOCATION_CODE AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE) IN (SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE) FROM NHK_MASTER_RECEIVE_PRODUCT GROUP BY PLANT, PRODUCT_CODE)
    LEFT JOIN MES_MASTER_CODE_EXC MMEC ON SUBSTR(NMRP.DOC_TYPE, 1, 2) = MMEC.ITEM_CODE AND MMEC.EXC_GROUP_CODE = 'CUSTOMER'
    LEFT JOIN MES_MASTER_REASON_LC MMRL ON MLI.REASON_CODE = MMRL.REASON_CODE
    LEFT JOIN MES_MASTER_FLOW_LC MMFL ON MMFL.FLOW_CODE = MLI.FLOW_CODE
    """,
    "Serial":"FROM MES_LOT_INFORMATION MLI LEFT JOIN ICM_LOT_SERIAL ILS ON ILS.LOT_ID = MLI.LOT_ID",
    "RelatedSearch":"FROM LV23_NHK.V_指図親子関係一覧",
    "RelatedSearch_tree":"""FROM filtered_hierarchy fh
                            ORDER BY fh.hierarchy_level, fh.node_id""",
    "ScheduleView":"FROM LV23_NHK.V_ASP_生産計画一覧 V_ASP_生産計画一覧",
    "Comparetest!":"FROM MES_LOT_INFORMATION MLI LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE AND MMPL.LOCATION_CODE = MLI.LOCATION_CODE LEFT JOIN LV23_NHK.V_ASP_生産計画一覧 V_ASP_生産計画一覧 ON V_ASP_生産計画一覧.指図NO = MLI.LOT_ID",
    "WIP":"FROM LV23_NHK.V_指図別工程別仕掛 LEFT JOIN V_ASP_生産計画一覧 ON V_指図別工程別仕掛.指図NO = V_ASP_生産計画一覧.指図NO",
    #"WIP":"FROM MES_LOT_INFORMATION MLI LEFT OUTER JOIN LV23_NHK.V_指図別工程別仕掛 ON MLI.LOT_ID = V_指図別工程別仕掛.指図NO LEFT OUTER JOIN V_ASP_生産計画一覧 ON V_指図別工程別仕掛.指図NO = V_ASP_生産計画一覧.指図NO",
    "WIP-I":f"""
    FROM
    MES_LOT_INFORMATION MLI
    LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE AND MMPL.LOCATION_CODE = MLI.LOCATION_CODE
    INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP ON NMRP.PRODUCT_CODE = MLI.PRODUCT_CODE AND NMRP.PLANT = MLI.LOCATION_CODE AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE)
    IN (
    SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE)
    FROM NHK_MASTER_RECEIVE_PRODUCT
    GROUP BY PLANT, PRODUCT_CODE)
    LEFT JOIN MES_MASTER_CODE_EXC MMEC ON SUBSTR(NMRP.DOC_TYPE, 1, 2) = MMEC.ITEM_CODE AND MMEC.EXC_GROUP_CODE = 'CUSTOMER'
    LEFT JOIN MES_MASTER_FLOW_LC MMFL ON MMFL.FLOW_CODE = MLI.FLOW_CODE
    LEFT JOIN MES_MASTER_FLOW MMF ON MLI.PROCESS_CODE = MMF.PROCESS_CODE AND MLI.FLOW_CODE = MMF.FLOW_CODE
    LEFT JOIN MES_MASTER_REASON_LC MMRL ON MLI.REASON_CODE = MMRL.REASON_CODE
    LEFT JOIN (SELECT STRING_ATTRIBUTE2, STRING_ATTRIBUTE3, MAX(WORK_END_DATE) AS COMPDATE FROM ASP_OPERATION WHERE is_assigned != 0 AND STATUS IN ('T','D','A') GROUP BY STRING_ATTRIBUTE2,STRING_ATTRIBUTE3) ENDD ON MLI.LOT_ID = ENDD.STRING_ATTRIBUTE2 AND MLI.PRODUCT_CODE = ENDD.STRING_ATTRIBUTE3
    """,
    "RemainTime":"FROM V_指図別工程別仕掛 W JOIN V_ASP_生産計画一覧 P ON P.指図NO = W.指図NO ",
    "Records":"""
    FROM MES_LOT_HISTORY MLH
	LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLH.PROCESS_CODE AND MMPL.LOCATION_CODE = MLH.LOCATION_CODE INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP ON NMRP.PRODUCT_CODE = MLH.PRODUCT_CODE AND NMRP.PLANT = MLH.LOCATION_CODE AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE) IN (
    SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE)
    FROM NHK_MASTER_RECEIVE_PRODUCT
    GROUP BY PLANT, PRODUCT_CODE)
	LEFT JOIN MES_MASTER_REASON_LC MMRL ON MLH.REASON_CODE = MMRL.REASON_CODE
	LEFT JOIN MES_MASTER_FLOW_LC MMFL ON MMFL.FLOW_CODE = MLH.FLOW_CODE
    """,
    "DBRelation":"FROM LV23_NHK.V_指図親子関係一覧 Relation INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP_P ON NMRP_P.PRODUCT_CODE = Relation.親品目 INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP_C ON NMRP_C.PRODUCT_CODE = Relation.子品目",
    "YellowBag":"FROM Z_KIBUKURO_INFORMATION",
    "DBSearch":"FROM NHK_MASTER_RECEIVE_PRODUCT",
    "Components": "FROM LV23_NHK.NHK_MASTER_MATERIAL_PROCESS NMMP LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP ON NMRP.PRODUCT_CODE = NMMP.PRODUCT_CODE LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT PROSAP ON NMMP.COMPONENT = PROSAP.PRODUCT_CODE",
    "Flow" : """FROM V_製品別工程別設備別標準時間 V1
            INNER JOIN ICM_CURRENT_PRODUCT_AND_FLOW ICPF
            ON V1.品目コード = SUBSTR(ICPF.FINAL_ITEM_CODE, 1, 10)
            AND V1.プラント = ICPF.LOCATION_CODE
            INNER JOIN MES_MASTER_FLOW MMF
            ON SUBSTR(ICPF.FINAL_ITEM_CODE, 12, 99) = MMF.FLOW_CODE
            AND V1.工程コード = MMF.PROCESS_CODE
            AND V1.プラント = MMF.LOCATION_CODE
            LEFT JOIN MES_MASTER_FLOW_LC MMFL
            ON MMF.FLOW_CODE = MMFL.FLOW_CODE
            AND MMF.LOCATION_CODE = MMFL.LOCATION_CODE""",
    "existing" : "Z_SHELF_LOT_INFORMATION",
    "all_data" : """Z_SHELF_LOT_INFORMATION 
                    INFO LEFT JOIN Z_MES_MASTER_SHELF_LC MST ON INFO.SHELF_NO = MST.SHELF_NO 
                    LEFT JOIN MES_LOT_INFORMATION MLI ON MLI.LOT_ID = INFO.LOT_ID 
                    LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE AND MMPL.LOCATION_CODE = MLI.LOCATION_CODE 
                    INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP ON NMRP.PRODUCT_CODE = MLI.PRODUCT_CODE AND NMRP.PLANT = MLI.LOCATION_CODE AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE) IN (SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE) FROM NHK_MASTER_RECEIVE_PRODUCT GROUP BY PLANT, PRODUCT_CODE) 
                    """,
    "updating" : "",
    "ksystem" : "FROM Z_RETURN_PRODUCT ",
    "product_inventory" : "FROM LV23_NHK.V_製品倉庫データ WHERE (DB図番 Like '%DB%')",
    "pov" : """
    FROM LV23_TEST.MES_LOT_INFORMATION MLI
    INNER JOIN V_ASP_生産計画一覧 VA ON MLI.LOT_ID = VA.指図NO
    """
    },

    "WhereCondition":{
    "ProcessRelated":f" WHERE MLI.LOT_ID IN ({lotid})",
    "RelatedSearch":f"WHERE (V_指図親子関係一覧.親指図 IN ({lotid}) AND V_指図親子関係一覧.子指図 IS NOT NULL) OR (V_指図親子関係一覧.子指図 IN ({lotid}) AND V_指図親子関係一覧.親指図 IS NOT NULL)",
    "RelatedSearch_tree":"",
    "ScheduleView": f"WHERE V_ASP_生産計画一覧.指図NO IN ({lotid})",
    "WIP":f"WHERE (V_指図別工程別仕掛.DB図番 IN ({lotid}))", #
    "WIP-I":f"""WHERE NMRP.OLD_PRODUCT_CODE IN ({lotid}) AND MLI.STATE_FLG NOT IN ('E', 'S') AND MLI.LOCATION_CODE IN ({plant})""",
    "WIP-IC":f"""WHERE NMRP.OLD_PRODUCT_CODE IN ({lotid}) AND MLI.STATE_FLG NOT IN ('E', 'S') AND MLI.USER_CODE <> 'RECEIVE_LOT' AND MLI.LOCATION_CODE IN ({plant})""",
    "RemainTime":f"WHERE V_指図別工程別仕掛.DB図番 IN ({lotid}) GROUP BY V_指図別工程別仕掛.指図NO, V_指図別工程別仕掛.DB図番, V_指図別工程別仕掛.フロー名称, V_指図別工程別仕掛.工程名, V_指図別工程別仕掛.処理区分, V_指図別工程別仕掛.更新日, V_指図別工程別仕掛.SEQ, V_指図別工程別仕掛.保留, V_指図別工程別仕掛.保留理由",
    "Records":f"WHERE (MLH.LOT_ID IN {lotid})",
    "DBRelation":f"WHERE Relation.子品目 IN (SELECT Relation_K.子品目 FROM LV23_NHK.V_指図親子関係一覧 Relation_K INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP_K ON NMRP_K.PRODUCT_CODE = Relation_K.親品目 WHERE NMRP_K.OLD_PRODUCT_CODE = {lotid} AND ROWNUM = 1)",
    "DBRelation_C":f"WHERE NMRP_C.OLD_PRODUCT_CODE = {lotid}",
    "YellowBag_Lot":f"WHERE LOT_ID IN {lotid} OR LOT_ID1 IN {lotid} OR LOT_ID2 IN {lotid}",
    "YellowBag_DB":f"WHERE DB_NO IN {lotid} OR DB_NO1 IN {lotid} OR DB_NO2 IN {lotid}",
    "DBSearch":f"WHERE OLD_PRODUCT_CODE LIKE {lotid}",
    "Components":f"WHERE NMRP.OLD_PRODUCT_CODE IN ({lotid}) OR PROSAP.OLD_PRODUCT_CODE IN ({lotid}) AND  NMMP.IS_ABOLISHED=0",
    "Flow" : f"WHERE V1.プラント = {plant} AND V1.DB図番 = {lotid} ORDER BY V1.品目コード, 内外製区分 DESC, フロー名称, 工順, V1.使用設備",
    "existing" : f"WHERE LOT_ID IN ({lotid})",
    "all_data" : "WHERE IN_OUT_FLG = 1",
    "updating" : f"WHERE LOT_ID = :4",
    "history" : "",
    "dataview_none" : "",
    "schedule_view_all" : f"WHERE (V_ASP_生産計画一覧.計画日_着手 = TO_DATE({lotid},'YYYY-MM-DD'))",
    "pov" : f"WHERE MLI.STATE_FLG NOT IN ('E', 'S') AND VA.ＤＢ図番 IN ({lotid}) ORDER BY 1, 10, 7"
    },
    "Misc.":{
    "Order":f"ORDER BY {cond[0]} {cond[1]}"
    }
    }

    if dict["WhereCondition"][where] and not lotid:
        raise E.NoKeyExc()

    sql = prefix + ",".join([dict[select][i] for i in attributes]) + " " + dict["FromDB"][fromdb] + " " + dict["WhereCondition"][where]
    #print(sql)
    if cond[0] != 0:
        sql += f" ORDER BY {cond[0]} {cond[1]}"
    #print(sql)
    return sql

def predefined(key: str) -> str:
    pred_sql = {
        'lot_status' : '''
    SELECT
        MLI.LOT_ID,
        MMPL.PROCESS_NAME,
        CASE WHEN MLI.HOLD_FLG = '1' THEN '保留' ELSE '-' END,
        MLI.PROCESS_CODE,
        ZSLI.SHELF_NO
    FROM
        MES_LOT_INFORMATION MLI LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE LEFT JOIN Z_SHELF_LOT_INFORMATION ZSLI ON ZSLI.LOT_ID = MLI.LOT_ID
    WHERE
        MLI.LOT_ID = :1
    ''',

    'inventory_combined': '''
SELECT 'CURRENT'              AS ROW_TYPE,
       NMRP.OLD_PRODUCT_CODE  AS DISPLAY_CODE,
       SUM(INV.QTY)           AS QTY,
       INV.LOCATION_CODE      AS LOCATION_CODE
FROM (
    SELECT DISTINCT PRODUCT_CODE, OLD_PRODUCT_CODE
    FROM   NHK_MASTER_RECEIVE_PRODUCT
    WHERE  OLD_PRODUCT_CODE = :1
) NMRP
JOIN V_ERP_INVENTORY_BEFORE_DAY INV
       ON INV.PRODUCT_CODE = NMRP.PRODUCT_CODE
GROUP BY NMRP.OLD_PRODUCT_CODE, INV.LOCATION_CODE
UNION ALL
SELECT DISTINCT
        'WIP'                    AS ROW_TYPE,
        PROSAP.OLD_PRODUCT_CODE  AS DISPLAY_CODE,
        NVL(INV.QTY, 0)         AS QTY,
        INV.LOCATION_CODE        AS LOCATION_CODE
 FROM   LV23_NHK.NHK_MASTER_MATERIAL_PROCESS NMMP
 JOIN (
     SELECT DISTINCT PRODUCT_CODE
     FROM   NHK_MASTER_RECEIVE_PRODUCT
     WHERE  OLD_PRODUCT_CODE = :2
 ) NMRP ON NMRP.PRODUCT_CODE = NMMP.PRODUCT_CODE
 JOIN (
     SELECT PRODUCT_CODE,
            MAX(OLD_PRODUCT_CODE) AS OLD_PRODUCT_CODE
     FROM   NHK_MASTER_RECEIVE_PRODUCT
     WHERE  PRODUCT_CODE LIKE '102%'
     GROUP  BY PRODUCT_CODE
 ) PROSAP ON NMMP.COMPONENT = PROSAP.PRODUCT_CODE
 LEFT JOIN V_ERP_INVENTORY_BEFORE_DAY INV
        ON INV.PRODUCT_CODE = PROSAP.PRODUCT_CODE
 WHERE  NMMP.IS_ABOLISHED = 0
''',

    'inventory_components': '''
SELECT DISTINCT
       PROSAP.PRODUCT_NO,
       COALESCE(
           CASE WHEN LENGTH(TRIM(PROSAP.OLD_PRODUCT_CODE)) > 2 THEN PROSAP.OLD_PRODUCT_CODE END,
           CASE WHEN LENGTH(TRIM(PROSAP.PRODUCT_NO)) > 2 THEN PROSAP.PRODUCT_NO END,
           PROSAP.PRODUCT_CODE
       ) AS COMPO_NAME,
       NVL(INV.QTY, 0)         AS QTY,
       INV.LOCATION_CODE        AS LOCATION_CODE
FROM   LV23_NHK.NHK_MASTER_MATERIAL_PROCESS NMMP
JOIN (
    SELECT DISTINCT PRODUCT_CODE
    FROM   NHK_MASTER_RECEIVE_PRODUCT
    WHERE  OLD_PRODUCT_CODE = :1
) NMRP ON NMRP.PRODUCT_CODE = NMMP.PRODUCT_CODE
JOIN (
    SELECT PRODUCT_CODE,
           MAX(OLD_PRODUCT_CODE) AS OLD_PRODUCT_CODE,
           MAX(PRODUCT_NO)       AS PRODUCT_NO
    FROM   NHK_MASTER_RECEIVE_PRODUCT
    WHERE  PRODUCT_CODE LIKE '103%'
    GROUP  BY PRODUCT_CODE
) PROSAP ON NMMP.COMPONENT = PROSAP.PRODUCT_CODE
LEFT JOIN V_ERP_INVENTORY_BEFORE_DAY INV
       ON INV.PRODUCT_CODE = PROSAP.PRODUCT_CODE
WHERE  NMMP.IS_ABOLISHED = 0
ORDER  BY COMPO_NAME
''',

    'NOKEY_vendor_master': '''
SELECT VENDOR_CODE, VENDOR_NAME FROM PROMAN_VENDOR_MASTER ORDER BY VENDOR_NAME
''',

    'NOKEY_asp_supply_schedule': '''
SELECT * FROM V_ASP_外注支給日程表
''',

    'vendor_processes': '''
SELECT PROCESS_CODE, PROCESS_NAME
FROM PROMAN_VENDOR_PROCESS_TYPE
WHERE VENDOR_CODE = :vendor_code
ORDER BY PROCESS_CODE
''',

    'vendor_commit_update': '''
UPDATE PROMAN_VENDOR_SCHEDULE
SET VENDOR_COMMIT_DATE = TO_DATE(:commit_date, 'YYYY-MM-DD')
WHERE SCHEDULE_ID = :schedule_id
''',

    'lot_info': '''
SELECT
    nmrp.OLD_PRODUCT_CODE  AS db_drawing,
    mmp.PRODUCT_NAME       AS product_name,
    CASE
        WHEN mli.HOLD_FLG = '1' THEN mmpl.PROCESS_NAME || ' [保留]'
        ELSE mmpl.PROCESS_NAME
    END AS lot_status
FROM MES_LOT_INFORMATION mli
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT nmrp
    ON nmrp.PRODUCT_CODE = mli.PRODUCT_CODE
   AND nmrp.PLANT = mli.LOCATION_CODE
   AND (nmrp.PLANT, nmrp.PRODUCT_CODE, nmrp.UP_DATE)
       IN (SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE)
           FROM NHK_MASTER_RECEIVE_PRODUCT
           GROUP BY PLANT, PRODUCT_CODE)
LEFT JOIN MES_MASTER_PRODUCT_LC mmp
    ON mmp.PRODUCT_CODE = mli.PRODUCT_CODE
LEFT JOIN MES_MASTER_PROCESS_LC mmpl
    ON mmpl.PROCESS_CODE = mli.PROCESS_CODE
WHERE mli.LOT_ID = :lot_id
''',

    'vendor_schedule_insert': '''
MERGE INTO PROMAN_VENDOR_SCHEDULE tgt
USING DUAL
ON (    tgt.LOT_ID       = :lot_id
    AND tgt.VENDOR_CODE  = :vendor_code
    AND tgt.PROCESS_CODE = :process_code)
WHEN NOT MATCHED THEN
    INSERT (LOT_ID, VENDOR_CODE, PROCESS_CODE, SHIP_DATE, NBD)
    VALUES (:lot_id, :vendor_code, :process_code,
            TO_DATE(:ship_date, 'YYYY-MM-DD'),
            TO_DATE(:nbd, 'YYYY-MM-DD'))
WHEN MATCHED THEN
    UPDATE SET SHIP_DATE = TO_DATE(:ship_date, 'YYYY-MM-DD'),
               NBD       = TO_DATE(:nbd, 'YYYY-MM-DD')
''',

    'vendor_schedule_delete': '''
DELETE FROM PROMAN_VENDOR_SCHEDULE WHERE SCHEDULE_ID = :schedule_id
''',

    'vendor_schedule': '''
SELECT
    vs.LOT_ID                                            AS lot_id,
    vpt.PROCESS_NAME                                     AS process_name,
    TO_CHAR(vs.SHIP_DATE,          'YYYY-MM-DD')         AS ship_date,
    TO_CHAR(vs.NBD,                'YYYY-MM-DD')         AS nbd,
    TO_CHAR(vs.VENDOR_COMMIT_DATE, 'YYYY-MM-DD')         AS vendor_commit_date,
    vs.SCHEDULE_ID                                       AS schedule_id,
    CASE
        WHEN mli.HOLD_FLG = '1' THEN mmpl.PROCESS_NAME || ' [保留]'
        ELSE mmpl.PROCESS_NAME
    END                                                  AS lot_status,
    nmrp.OLD_PRODUCT_CODE                                AS db_drawing,
    mmp.PRODUCT_NAME                                     AS product_name
FROM PROMAN_VENDOR_SCHEDULE vs
INNER JOIN PROMAN_VENDOR_PROCESS_TYPE vpt
    ON vs.VENDOR_CODE  = vpt.VENDOR_CODE
   AND vs.PROCESS_CODE = vpt.PROCESS_CODE
LEFT JOIN MES_LOT_INFORMATION mli
    ON mli.LOT_ID = vs.LOT_ID
LEFT JOIN MES_MASTER_PROCESS_LC mmpl
    ON mmpl.PROCESS_CODE = mli.PROCESS_CODE
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT nmrp
    ON nmrp.PRODUCT_CODE = mli.PRODUCT_CODE
   AND nmrp.PLANT = mli.LOCATION_CODE
   AND (nmrp.PLANT, nmrp.PRODUCT_CODE, nmrp.UP_DATE)
       IN (SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE)
           FROM NHK_MASTER_RECEIVE_PRODUCT
           GROUP BY PLANT, PRODUCT_CODE)
LEFT JOIN MES_MASTER_PRODUCT_LC mmp
    ON mmp.PRODUCT_CODE = mli.PRODUCT_CODE
WHERE vs.VENDOR_CODE = :vendor_code
  AND (   vs.SHIP_DATE          BETWEEN TO_DATE(:date_from, 'YYYY-MM-DD') AND TO_DATE(:date_to, 'YYYY-MM-DD')
       OR vs.NBD                BETWEEN TO_DATE(:date_from, 'YYYY-MM-DD') AND TO_DATE(:date_to, 'YYYY-MM-DD')
       OR vs.VENDOR_COMMIT_DATE BETWEEN TO_DATE(:date_from, 'YYYY-MM-DD') AND TO_DATE(:date_to, 'YYYY-MM-DD')
      )
ORDER BY vs.LOT_ID, vpt.PROCESS_CODE
''',

    }

    return pred_sql.get(key, None)


'''
Full length process map sql
SELECT 
    MLI.LOT_ID, 
    MLI.PRODUCT_CODE, 
    MLH.FLOW_CODE, 
    MMPL.PROCESS_NAME, 
    MLH.HOLD_FLG,
    NULL ,
    NULL AS SCH_DATE,
    NULL ,
    MMRL.COMMENTS, 
    MLI.STATE_FLG, 
    MLI.REASON_CODE, 
    MLH.UP_DATE AS UP_DATE,
    'HISTORY' 
FROM LV23_TEST.MES_LOT_INFORMATION MLI
INNER JOIN MES_LOT_HISTORY MLH ON MLI.LOT_ID = MLH.LOT_ID
LEFT JOIN MES_MASTER_PROCESS_LC MMPL ON MMPL.PROCESS_CODE = MLH.PROCESS_CODE 
    AND MMPL.LOCATION_CODE = MLH.LOCATION_CODE
LEFT JOIN MES_MASTER_REASON_LC MMRL ON MLH.REASON_CODE = MMRL.REASON_CODE
WHERE 
    MLI.STATE_FLG NOT IN ('E', 'S') 
    AND MLI.LOT_ID = '4235048073'

UNION ALL

SELECT 
    MLI.LOT_ID, 
    MLI.PRODUCT_CODE, 
    NULL , 
    NULL , 
    NULL ,
    VA.工程名,
    VA.着手予定日 AS SCH_DATE,
    PAL.LEADTIME,
    NULL , 
    MLI.STATE_FLG, 
    MLI.REASON_CODE, 
    NULL AS UP_DATE ,
    'SCHEDULE' 
FROM LV23_TEST.MES_LOT_INFORMATION MLI
INNER JOIN V_ASP_生産計画一覧 VA ON MLI.LOT_ID = VA.指図NO
LEFT JOIN PROMAN_AVG_LEADTIME PAL ON VA.ＤＢ図番 = PAL.DB AND VA.工程CD = PAL.PROCESS_CODE
WHERE MLI.STATE_FLG NOT IN ('E', 'S') 
    AND MLI.LOT_ID = '4235048073'

ORDER BY MLI.LOT_ID, UP_DATE, SCH_DATE
'''