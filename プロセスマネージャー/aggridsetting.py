import functools

@functools.lru_cache(maxsize=None)
def generate_col_settings(mode: str, pds: bool = False) -> dict:
        '''
        # pds
        - True : Used as header in Pandas Dataframe generation, will omit the first column when mode is 'i' or 'w'.
        - False : Used as header in AgGrid grid generation, will output all avaliable columns according to key. 
        Default as False
        '''

        mode = 'w' if mode == 'e' else mode

        date_filter_parms = '''(filterLocalDateAtMidnight, cellValue) => {
                    const dateParts = cellValue.split(" ");
                    const splited_dateParts = dateParts[0].split("-");
                    const cellDate = new Date(Number(splited_dateParts[0]), Number(splited_dateParts[1]) - 1, Number(splited_dateParts[2]));
                    return cellDate < filterLocalDateAtMidnight ? -1 : cellDate > filterLocalDateAtMidnight ? 1 : 0;
                }'''
        

        aggrid_columns_settings = {
        'i' :
                [
                {'field' : 'Selected',
                'headerName' : 'S',
                'hide' : 'True'},

                {'field' : 'lotid',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'headerComponent': 'CheckboxHeader',
                'flex': 2,
                        },
                {'field' : 'customer',
                'headerName' : '客先',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"]
                        },
                'flex': 1,
                        },
                {'field' : 'productcode',
                'headerName' : '図番',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 2,
                        },
                {'field' : 'currentprocess',
                'headerName' : '現在工程',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 2,
                        },
                {'field' : 'lotstate',
                'headerName' : '指図状態',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 1,
                        },
                {'field' : 'onholdreason',
                'headerName' : '保留/理由',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 3,
                        },
                {'field' : 'lastupdatedate',
                'headerName' : '最終更新日時',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        },
                'flex': 2,
                        },
                {'field' : 'state',
                'headerName' : '処理ステータス',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 1,
                        },
                {'field' : 'restriction',
                'headerName' : '出荷制限',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 3,
                        },
                {'field' : 'flowname',
                'headerName' : 'フロー',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 2,
                        }]
                ,
        'w' : [
                {'field' : 'Selected',
                'headerName' : 'S',
                'hide' : 'True'
                },
                {'field' : 'lotid',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 2,
                        },
                {'field' : 'flowname',
                'headerName' : 'フロー',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 2,
                        },
                {'field' : 'currentprocess',
                'headerName' : '現在工程',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 3,
                        },
                {'field' : 'inprogress',
                'headerName' : '処理ステータス',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 1,
                        },
                {'field' : 'lastupdatedate',
                'headerName' : '最終更新日時',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        },
                'flex': 2,
                        },
                {'field' : 'SEQ',
                'headerName' : '工順(SEQ)',
                'filter' : 'agNumberColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 1,
                        },
                {'field' : 'onholdreason',
                'headerName' : '保留/理由',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 3,
                        },
                {'field' : 'completedate',
                'headerName' : '完成予定日',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        },
                'flex': 1,
                        },
                {'field' : 'StartStatus',
                'headerName' : '投入状態',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 1,
                        },
                {'field' : 'restriction',
                'headerName' : '出荷制限',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 3,
                        },
                {'field' : 'productcode',
                'headerName' : '図番',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        },
                'flex': 1,
                        }
                        ],
        'h' : [
                {'field' : '指図',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'フローコード',
                'headerName' : 'フローコード',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 90},
                {'field' : '工程',
                'headerName' : '工程',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '処理ステータス',
                'headerName' : '処理ステータス',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 50},
                {'field' : '設備コード',
                'headerName' : '設備コード',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 50},
                {'field' : '保留',
                'headerName' : '保留',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 40},
                {'field' : '保留理由',
                'headerName' : '保留理由',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '状態',
                'headerName' : '状態',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 50},
                {'field' : '処理日時',
                'headerName' : '処理日時',
                'filter' : 'agDateColumnFilter',
                'minWidth' : 50},
                {'field' : '処理者氏名コード',
                'headerName' : '処理者氏名コード',
                'filter' : 'agDateColumnFilter',
                'minWidth' : 50},
                {'field' : 'フロー',
                'headerName' : 'フロー',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 90}
                ],
        'y' : [
                {'field' : '指図',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '図番',
                'headerName' : '図番',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 80},
                {'field' : '中間指図 1',
                'headerName' : '中間指図 1',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '中間図番 1',
                'headerName' : '中間図番 1',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 80},
                {'field' : '中間指図 2',
                'headerName' : '中間指図 2',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '中間図番 2',
                'headerName' : '中間図番 2',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 80},
                {'field' : '出荷制限区分',
                'headerName' : '出荷制限区分',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 40},
                {'field' : '出荷制限理由',
                'headerName' : '出荷制限理由',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '制限ステータス',
                'headerName' : '制限ステータス',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 40},
                {'field' : '制限解除者',
                'headerName' : '制限解除者',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 40},
                {'field' : '制限解除日',
                'headerName' : '制限解除日',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 40},
                {'field' : '制限設定者',
                'headerName' : '制限設定者',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 40}
                ],
        'v' : 
                [
                {'field' : '指図',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '現在工程',
                'headerName' : '現在工程',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : '開始時間',
                'headerName' : '開始時間',
                'filter' : 'agDateColumnFilter',
                'minWidth' : 100},
                {'field' : '完了予定時間',
                'headerName' : '完了予定時間',
                'filter' : 'agDateColumnFilter',
                'minWidth' : 100},
                {'field' : '資源',
                'headerName' : '資源',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 50},
                {'field' : '状態',
                'headerName' : '状態',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 60},],
        's' : 
                [
                {'field' : '指図',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'シリアル1',
                'headerName' : 'シリアル1',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'シリアル2',
                'headerName' : 'シリアル2',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'シリアル3',
                'headerName' : 'シリアル3',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'シリアル4',
                'headerName' : 'シリアル4',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'シリアル5',
                'headerName' : 'シリアル5',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},
                {'field' : 'シリアル6',
                'headerName' : 'シリアル6',
                'filter' : 'agTextColumnFilter',
                'minWidth' : 100},],
        'A' : 
                [

                {'field' : 'lotid',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'shelf_no',
                'headerName' : '棚番コード',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'shelf_name',
                'headerName' : 'エリア名',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'user_code',
                'headerName' : '氏名コード',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'up_date',
                'headerName' : '搬入日時',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'urgent_flg',
                'headerName' : '特急対応品',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'designated_flg',
                'headerName' : '所定棚番コード',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'db',
                'headerName' : 'DB図番',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'process',
                'headerName' : '工程名',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'on_hold_status',
                'headerName' : '保留状態',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        }        
                        ],
        'D' : 
                [
                {'field' : 'index_id',
                'headerName' : 'Index',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },

                {'field' : 'KIBUKURO_NO',
                'headerName' : '黄袋番号',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'LOCATION_CODE',
                'headerName' : 'プラント',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'LOCATION_SEQ_NO',
                'headerName' : 'LOCATION_SEQ_NO',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'APPLICATION_NO',
                'headerName' : 'APPLICATION_NO',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'PUBLISH_DATE',
                'headerName' : '登録日',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'SERIAL',
                'headerName' : 'シリアル',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'LOT_ID',
                'headerName' : '指図',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'DB_NO',
                'headerName' : 'DB図番',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'PRODUCT_CODE',
                'headerName' : '品目コード',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'LOT_ID1',
                'headerName' : '関連指図1',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'DB_NO1',
                'headerName' : '関連DB1',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'PRODUCT_CODE1',
                'headerName' : '品目コード1',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'LOT_ID2',
                'headerName' : '関連指図2',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'DB_NO2',
                'headerName' : '関連DB2',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'PRODUCT_CODE2',
                'headerName' : '品目コード2',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'NOT_SHIPPABLE_TYPE',
                'headerName' : '黄袋分類',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'NOT_SHIPPABLE_REASON',
                'headerName' : '黄袋理由',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'WRITTEN_BY',
                'headerName' : '発行者',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'REQUEST_BY',
                'headerName' : '依頼者',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'DISTRIBUTE_DATE',
                'headerName' : '発行日',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'RELEASE_DATE',
                'headerName' : '解除日',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'RELEASE_BY',
                'headerName' : '解除者',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'RELEASE_TYPE',
                'headerName' : '解除タイプ',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'RELEASE_REASON',
                'headerName' : '解除理由',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'COLLECTION_DATE',
                'headerName' : 'COLLECTION_DATE',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'COLLECTION_REVIEWER',
                'headerName' : 'COLLECTION_REVIEWER',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'REMARKS1',
                'headerName' : 'REMARKS1',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'REMARKS2',
                'headerName' : 'REMARKS2',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'CREATE_DATE',
                'headerName' : 'CREATE_DATE',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'UP_DATE',
                'headerName' : 'UP_DATE',
                'filter' : 'agDateColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        ':comparator' : date_filter_parms,
                        'maxNumConditions' : 1
                        }
                        },
                {'field' : 'RELEASE_FLG',
                'headerName' : '解除フラグ',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {'field' : 'SPECIAL_NEMBER',
                'headerName' : 'SPECIAL_NEMBER',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                        
                        ],

        'K' :   [
                {'field' : 'index_id',
                'headerName' : 'Index',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },

                {
                        'field': 'ORDER_NO',
                        'headerName': '指図番号',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'ACCEPT_DATE',
                        'headerName': '受入日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'ITEM',
                        'headerName': '品目',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'ITEM_DESCRIPTION',
                        'headerName': '品番注記',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'K_LOT_ID',
                        'headerName': 'KロットID',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'DUTY_DIVISION',
                        'headerName': '担当区分',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'REASON',
                        'headerName': '理由',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'CUSTOMERS',
                        'headerName': '顧客番号',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'CUSTOMERS_DESCRIPTION',
                        'headerName': '顧客名',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'CUSTOMER_NUMBER',
                        'headerName': '顧客注番',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'SLIP_TYPE',
                        'headerName': '伝票種別',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'SERIAL_NUMBER',
                        'headerName': '書類シリアル番号',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'CONF_SERIAL_NUMBER',
                        'headerName': '現品確認シリアル番号',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'APPROVAL',
                        'headerName': '承認',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'APPROVAL_DATE',
                        'headerName': '承認日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'FA_COMP_DATE',
                        'headerName': 'FA完了日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'SURVEY',
                        'headerName': '調査',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'SURVEY_RESULT',
                        'headerName': '調査結果',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MIY_APPROVAL',
                        'headerName': 'MIY承認',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MIY_APPROVAL_DATE',
                        'headerName': 'MIY承認日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'MIY_SHIP',
                        'headerName': 'MIY出荷',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'NEXT_LOT_ID',
                        'headerName': '後続ロットID',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'FLOW_CODE',
                        'headerName': 'フローコード',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'REMARKS',
                        'headerName': '備考',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'K_LOT_USER',
                        'headerName': 'Kロット担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'K_LOT_DATE',
                        'headerName': 'Kロット日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'STORE_USER',
                        'headerName': '入庫担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'STORE_DATE',
                        'headerName': '入庫日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'SHIP_USER',
                        'headerName': '出荷担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'SHIP_DATE',
                        'headerName': '出荷日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'ARRIVE_USER',
                        'headerName': '到着担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'ARRIVE_DATE',
                        'headerName': '到着日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'MIY_K_LOT_USER',
                        'headerName': 'MIY Kロット担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MIY_K_LOT_DATE',
                        'headerName': 'MIY Kロット日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'MIY_STORE_USER',
                        'headerName': 'MIY入庫担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MIY_STORE_DATE',
                        'headerName': 'MIY入庫日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'MIY_SHIP_USER',
                        'headerName': 'MIY出荷担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MIY_SHIP_DATE',
                        'headerName': 'MIY出荷日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'MIY_ARRIVE_USER',
                        'headerName': 'MIY到着担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MIY_ARRIVE_DATE',
                        'headerName': 'MIY到着日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'FA_COMP_INP_USER',
                        'headerName': 'FA完了入力者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'FA_COMP_INP_DATE',
                        'headerName': 'FA完了入力日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'SURVEY_COMP_USER',
                        'headerName': '調査完了担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'SURVEY_COMP_DATE',
                        'headerName': '調査完了日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'LV_START_USER',
                        'headerName': 'LV開始担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'LV_START_DATE',
                        'headerName': 'LV開始日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'LV_END_USER',
                        'headerName': 'LV終了担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'LV_END_DATE',
                        'headerName': 'LV終了日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'SCRAP_USER',
                        'headerName': '廃棄担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'SCRAP_DATE',
                        'headerName': '廃棄日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'RETURN_USER',
                        'headerName': '返却担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'RETURN_DATE',
                        'headerName': '返却日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'CANCEL_USER',
                        'headerName': 'キャンセル担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'CANCEL_DATE',
                        'headerName': 'キャンセル日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'COMPLETE_USER',
                        'headerName': '完了担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'COMPLETE_DATE',
                        'headerName': '完了日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'RECONFIRM_USER',
                        'headerName': '再確認担当者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'RECONFIRM_DATE',
                        'headerName': '再確認日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'CREATED_DATE',
                        'headerName': '作成日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'MODIFIED_USER',
                        'headerName': '更新者',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'MODIFIED_DATE',
                        'headerName': '更新日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': 'STATUS',
                        'headerName': 'ステータス',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'REL_ORDER_NO',
                        'headerName': '関連指図番号',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'COMP_ORDER_NO',
                        'headerName': '部品指図番号',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'PLANT',
                        'headerName': '工場',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                ],
        'F' :   [
                {
                        'field': 'ロケーション',
                        'headerName': 'ロケーション',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '品目コード',
                        'headerName': '品目コード',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'DB図番',
                        'headerName': 'DB図番',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '指図',
                        'headerName': '指図',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '入庫日時',
                        'headerName': '入庫日時',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': '最終検査日時(MES)',
                        'headerName': '最終検査日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': date_filter_parms,
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': '指図ステータス',
                        'headerName': '指図ステータス',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報1',
                        'headerName': '製品情報1',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報2',
                        'headerName': '製品情報2',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報3',
                        'headerName': '製品情報3',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報4',
                        'headerName': '製品情報4',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報5',
                        'headerName': '製品情報5',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報6',
                        'headerName': '製品情報6',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報7',
                        'headerName': '製品情報7',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報8',
                        'headerName': '製品情報8',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報9',
                        'headerName': '製品情報9',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '製品情報10',
                        'headerName': '製品情報10',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                }
                ],
        'S' : [
                {'field' : 'index_id',
                'headerName' : 'Index',
                'filter' : 'agTextColumnFilter',
                'filterParams' : {
                        'buttons': ["reset", "apply"],
                        }
                        },
                {
                        'field': '計画日_着手',
                        'headerName': '計画日_着手',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '計画日_完了',
                        'headerName': '計画日_完了',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '指図NO',
                        'headerName': '指図NO',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '品目',
                        'headerName': '品目',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'ＤＢ図番',
                        'headerName': 'DB図番',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'フローCD',
                        'headerName': 'フローCD',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'フロー名称',
                        'headerName': 'フロー名称',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '管理CD',
                        'headerName': '管理CD',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '客先名',
                        'headerName': '客先名',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '客先品番',
                        'headerName': '客先品番',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '数量',
                        'headerName': '数量',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '資源',
                        'headerName': '資源',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '大区分',
                        'headerName': '大区分',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '大区分名称',
                        'headerName': '大区分名称',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '中区分',
                        'headerName': '中区分',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '中区分名称',
                        'headerName': '中区分名称',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '工程CD',
                        'headerName': '工程CD',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '工程名',
                        'headerName': '工程名',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '着手予定日',
                        'headerName': '着手予定日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': 'date_filter_parms',
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': '完了予定日',
                        'headerName': '完了予定日',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': 'date_filter_parms',
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': '工順',
                        'headerName': '工順',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'ステータス',
                        'headerName': 'ステータス',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': 'LV更新日時',
                        'headerName': 'LV更新日時',
                        'filter': 'agDateColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        ':comparator': 'date_filter_parms',
                        'maxNumConditions': 1
                        }
                },
                {
                        'field': '担当者設定納期',
                        'headerName': '担当者設定納期',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '標準手作業時間',
                        'headerName': '標準手作業時間',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '標準装置時間',
                        'headerName': '標準装置時間',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '標準時間',
                        'headerName': '標準時間',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '中間品目',
                        'headerName': '中間品目',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                },
                {
                        'field': '試作量産区分',
                        'headerName': '試作量産区分',
                        'filter': 'agTextColumnFilter',
                        'filterParams': {
                        'buttons': ["reset", "apply"],
                        }
                }
                ],

        'p': [
                {'field': 'vendor_name',        'headerName': '仕入先名',   'width': 160,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': 'lot_id',             'headerName': '指図番号',   'width': 130,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': 'process_name',       'headerName': '工程名',     'width': 130,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': 'lot_status',         'headerName': '工程状態',   'width': 150,
                 'filter': 'agTextColumnFilter'},
                {'field': 'ship_date',          'headerName': 'NHK支給日',  'width': 120,
                 'filter': 'agTextColumnFilter',},
                 #':cellClassRules': "{ 'bg-orange-200': params => params.data.pull_in_date != null && params.data.pull_in_date !== '' }"},
                {'field': 'nbd',                'headerName': '納期',       'width': 120,
                 'filter': 'agTextColumnFilter'},
                {'field': 'vendor_commit_date', 'headerName': '回答納期',   'width': 120,
                 'filter': 'agTextColumnFilter', 'editable': True},
                {'field': 'comments1',          'headerName': '備考1',      'width': 120, 'editable': True},
                {'field': 'comments2',          'headerName': '備考2',      'width': 120, 'editable': True},
                {'field': 'comments3',          'headerName': '備考3',      'width': 120, 'editable': True},
                {'field': 'schedule_id',        'headerName': 'ID',         'hide': True},
                {'field': 'pull_in_date',       'headerName': '前倒可能日', 'hide': True},
                {'field': 'row_type',           'headerName': 'type',       'hide': True},
        ],

        'asp': [
                {'field': '客先',           'headerName': '客先',           'width': 100,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': 'DB図番',         'headerName': 'DB図番',         'width': 140,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '客先品番',       'headerName': '客先品番',       'width': 130,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '品名',           'headerName': '品名',           'width': 160,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '指図',           'headerName': '指図',           'width': 120,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': 'フローコード',   'headerName': 'フローコード',   'width': 120,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '業者',           'headerName': '業者',           'width': 150,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '支給工程名称',   'headerName': '支給工程名称',   'width': 150,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '支給日時',       'headerName': '支給日時',       'width': 140,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '支給ステータス', 'headerName': '支給ステータス', 'width': 130,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '外注工程名称',   'headerName': '外注工程名称',   'width': 150,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '納期',           'headerName': '納期',           'width': 120,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '外注ステータス', 'headerName': '外注ステータス', 'width': 130,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '保留ステータス', 'headerName': '保留ステータス', 'width': 130,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '外注工順',       'headerName': '外注工順',       'width': 100,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
                {'field': '資源',           'headerName': '資源',           'width': 120,
                 'filter': 'agTextColumnFilter', 'filterParams': {'buttons': ['reset', 'apply']}},
        ],

        }
        if mode in 'iw' and pds:
                return aggrid_columns_settings[mode][1:]
        else:
                return aggrid_columns_settings[mode]