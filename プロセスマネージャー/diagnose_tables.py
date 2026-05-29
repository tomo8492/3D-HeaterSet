# diagnose_tables.py
# NHK_LOT_PROCESS_HISTORY / Z_DUEOUT_HISTORY の構造・性能・データ確認
#
# 目的:
#   「期間中に工程が進んだ指図」を高速に取得できるか判定するための調査。
#   END_TIME / DUEOUT_DATE_TIME のインデックス有無と件数規模を確認する。
#
# 実行:
#   cd 出庫漏れ確認
#   python diagnose_tables.py

import sys
import os
import oracledb
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db_config as CFG

_TIMEOUT_MS = 300_000   # 5分


def connect():
    print(f'接続中: {CFG.DB_USER}@{CFG.DB_DSN} ...')
    conn = oracledb.connect(
        user=CFG.DB_USER,
        password=CFG.DB_PASSWORD,
        dsn=CFG.DB_DSN,
        config_dir=CFG.DB_CONFIG_DIR,
    )
    conn.call_timeout = _TIMEOUT_MS
    print('接続成功\n')
    return conn


def run(conn, sql, params=None, label=''):
    print(f'=== {label} ===')
    with conn.cursor() as c:
        c.arraysize = 1000
        c.execute(sql, parameters=params or [])
        cols = [d[0] for d in c.description]
        rows = c.fetchall()
    col_w = [max(len(str(col)), max((len(str(r[i])) for r in rows), default=0))
             for i, col in enumerate(cols)]
    header = '  ' + ' | '.join(str(col).ljust(col_w[i]) for i, col in enumerate(cols))
    print(header)
    print('  ' + '-' * len(header))
    for row in rows:
        print('  ' + ' | '.join(str(v if v is not None else 'NULL').ljust(col_w[i])
                                 for i, v in enumerate(row)))
    print(f'  → {len(rows)} 件\n')
    return rows


def main():
    conn = connect()

    lot_id = input(
        '確認する指図NOを入力 (空Enterで 4235962024 を使用): '
    ).strip() or '4235962024'
    print(f'→ 指図NO: {lot_id}\n')

    # 先月・今月の日付を計算
    today = date.today()
    first_of_this_month = today.replace(day=1)
    first_of_last_month = (first_of_this_month - timedelta(days=1)).replace(day=1)
    print(f'先月: {first_of_last_month} 〜 {first_of_this_month - timedelta(days=1)}')
    print(f'今月: {first_of_this_month} 〜 {today}\n')

    # ══════════════════════════════════════════════════════════════
    # ブロック1: NHK_LOT_PROCESS_HISTORY の構造確認
    # ══════════════════════════════════════════════════════════════

    # T1: カラム一覧
    run(conn, '''
SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
FROM ALL_TAB_COLUMNS
WHERE OWNER      = UPPER(:1)
  AND TABLE_NAME = 'NHK_LOT_PROCESS_HISTORY'
ORDER BY COLUMN_ID
''',
        params=[CFG.DB_USER],
        label='T1: NHK_LOT_PROCESS_HISTORY カラム一覧')

    # T2: インデックス一覧（END_TIME にインデックスがあるか確認）
    run(conn, '''
SELECT I.INDEX_NAME, C.COLUMN_NAME, C.COLUMN_POSITION, I.UNIQUENESS
FROM ALL_INDEXES     I
JOIN ALL_IND_COLUMNS C
    ON  C.INDEX_NAME  = I.INDEX_NAME
    AND C.TABLE_OWNER = I.TABLE_OWNER
    AND C.TABLE_NAME  = I.TABLE_NAME
WHERE I.TABLE_OWNER = UPPER(:1)
  AND I.TABLE_NAME  = 'NHK_LOT_PROCESS_HISTORY'
ORDER BY I.INDEX_NAME, C.COLUMN_POSITION
''',
        params=[CFG.DB_USER],
        label='T2: NHK_LOT_PROCESS_HISTORY インデックス一覧')

    # T3: 指図NOの履歴（LOT_ID で絞るので高速）
    run(conn, '''
SELECT LOT_ID, PROCESS_CODE, STATE_FLG, END_TIME
FROM NHK_LOT_PROCESS_HISTORY
WHERE LOT_ID = :1
ORDER BY END_TIME DESC NULLS LAST
FETCH FIRST 20 ROWS ONLY
''',
        params=[lot_id],
        label=f'T3: {lot_id} の NHK_LOT_PROCESS_HISTORY 履歴')

    # T4: 先月1か月の件数（END_TIME インデックスが効くか測定）
    #     ★タイムアウト注意: インデックスがない場合は全件スキャンになる
    print(f'T4: 先月({first_of_last_month}〜)の件数取得中... (最大5分)')
    run(conn, '''
SELECT COUNT(*) AS 件数
FROM NHK_LOT_PROCESS_HISTORY
WHERE END_TIME >= :1
  AND END_TIME <  :2
  AND STATE_FLG = 'N'
''',
        params=[first_of_last_month, first_of_this_month],
        label=f'T4: NHK_LOT_PROCESS_HISTORY 先月の件数（STATE_FLG=N）')

    # T5: 先月の件数がわかった場合、指図単位のDISTINCT件数も確認
    print('T5: 先月の指図件数（DISTINCT LOT_ID）取得中...')
    run(conn, '''
SELECT COUNT(DISTINCT LOT_ID) AS 指図件数
FROM NHK_LOT_PROCESS_HISTORY
WHERE END_TIME >= :1
  AND END_TIME <  :2
  AND STATE_FLG = 'N'
''',
        params=[first_of_last_month, first_of_this_month],
        label=f'T5: NHK_LOT_PROCESS_HISTORY 先月の指図件数（DISTINCT）')

    # ══════════════════════════════════════════════════════════════
    # ブロック2: Z_DUEOUT_HISTORY の構造確認
    # ══════════════════════════════════════════════════════════════

    # T6: カラム一覧
    run(conn, '''
SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
FROM ALL_TAB_COLUMNS
WHERE OWNER      = UPPER(:1)
  AND TABLE_NAME = 'Z_DUEOUT_HISTORY'
ORDER BY COLUMN_ID
''',
        params=[CFG.DB_USER],
        label='T6: Z_DUEOUT_HISTORY カラム一覧')

    # T7: インデックス一覧
    run(conn, '''
SELECT I.INDEX_NAME, C.COLUMN_NAME, C.COLUMN_POSITION, I.UNIQUENESS
FROM ALL_INDEXES     I
JOIN ALL_IND_COLUMNS C
    ON  C.INDEX_NAME  = I.INDEX_NAME
    AND C.TABLE_OWNER = I.TABLE_OWNER
    AND C.TABLE_NAME  = I.TABLE_NAME
WHERE I.TABLE_OWNER = UPPER(:1)
  AND I.TABLE_NAME  = 'Z_DUEOUT_HISTORY'
ORDER BY I.INDEX_NAME, C.COLUMN_POSITION
''',
        params=[CFG.DB_USER],
        label='T7: Z_DUEOUT_HISTORY インデックス一覧')

    # T8: 指図NOの出庫履歴
    run(conn, '''
SELECT LOT_ID, COMPONENT, COUNT, MOVEMENT_TYPE, DUEOUT_DATE_TIME
FROM Z_DUEOUT_HISTORY
WHERE LOT_ID = :1
ORDER BY DUEOUT_DATE_TIME DESC NULLS LAST
FETCH FIRST 20 ROWS ONLY
''',
        params=[lot_id],
        label=f'T8: {lot_id} の Z_DUEOUT_HISTORY 出庫履歴')

    # T9: 先月1か月の出庫件数（DUEOUT_DATE_TIME インデックスが効くか）
    print(f'T9: 先月の出庫件数取得中...')
    run(conn, '''
SELECT COUNT(*) AS 件数, COUNT(DISTINCT LOT_ID) AS 指図件数
FROM Z_DUEOUT_HISTORY
WHERE DUEOUT_DATE_TIME >= :1
  AND DUEOUT_DATE_TIME <  :2
  AND MOVEMENT_TYPE = '261'
''',
        params=[first_of_last_month, first_of_this_month],
        label=f'T9: Z_DUEOUT_HISTORY 先月の件数（MOVEMENT_TYPE=261）')

    # ══════════════════════════════════════════════════════════════
    # ブロック4: NHK_MASTER_MATERIAL_PROCESS (BOM) 確認
    # ══════════════════════════════════════════════════════════════

    # 指図NOから品目コードを取得
    sap_rows = run(conn, '''
SELECT DISTINCT PRODUCT_CODE FROM NHK_LOT_PROCESS_HISTORY
WHERE LOT_ID = :1
FETCH FIRST 1 ROWS ONLY
''',
        params=[lot_id],
        label=f'T11: {lot_id} の品目コード（NHK_LOT_PROCESS_HISTORY）')

    if sap_rows:
        sap_code = str(sap_rows[0][0])

        # T12: BOM登録状況
        run(conn, '''
SELECT
    BOM.COMPONENT,
    BOM.COUNT,
    BOM.IS_ABOLISHED,
    NMRP.OLD_PRODUCT_CODE AS DB図番
FROM NHK_MASTER_MATERIAL_PROCESS BOM
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP
    ON  NMRP.PRODUCT_CODE = BOM.COMPONENT
    AND NMRP.PLANT        = '5501'
WHERE BOM.PRODUCT_CODE = :1
ORDER BY BOM.COMPONENT
''',
            params=[sap_code],
            label=f'T12: 品目 {sap_code} の NHK_MASTER_MATERIAL_PROCESS BOM一覧（IS_ABOLISHED含む）')

        # T13: IS_ABOLISHED=1 以外（現行BOM）のみ
        run(conn, '''
SELECT COUNT(*) AS 有効BOM件数
FROM NHK_MASTER_MATERIAL_PROCESS
WHERE PRODUCT_CODE = :1
  AND (IS_ABOLISHED IS NULL OR TO_CHAR(IS_ABOLISHED) <> '1')
''',
            params=[sap_code],
            label=f'T13: 品目 {sap_code} の有効BOM件数（IS_ABOLISHED≠1）')

    # T14: Z_DUEOUT_HISTORY の個別行確認（集計前の生データ）
    #      出庫数量が異常に大きい場合、集計前に重複行がないか確認する
    run(conn, '''
SELECT LOCATION_CODE, FLOW_CODE, PROCESS_CODE, COMPONENT,
       COUNT AS 数量, DATA_DIVISION, DUEOUT_DATE_TIME
FROM Z_DUEOUT_HISTORY
WHERE LOT_ID       = :1
  AND MOVEMENT_TYPE = '261'
ORDER BY COMPONENT, DUEOUT_DATE_TIME
FETCH FIRST 30 ROWS ONLY
''',
        params=[lot_id],
        label=f'T14: {lot_id} の Z_DUEOUT_HISTORY 生データ（集計前・COMPONENT単位）')

    conn.close()
    print('調査完了。')


if __name__ == '__main__':
    main()
