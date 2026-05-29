# diagnose_active_orders.py
# 流動中指図の検索漏れ原因を調査するための診断スクリプト
#
# 実行:
#   cd 出庫漏れ確認
#   python diagnose_active_orders.py
#
# タイムアウト方針:
#   ビュー同士のJOINは避ける。単一ビューの集計 or 指図NO指定のポイントクエリのみ使用。

import sys
import os
import oracledb

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db_config as CFG

_TIMEOUT_MS = 300_000   # 5分（重いビューに対応）


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

    # ──────────────────────────────────────────────────────────
    # Q0: 特定の指図NOで直接診断（最優先・最速）
    #     検索できなかった指図NOを1件入力してもらい、3段階で原因を絞る
    # ──────────────────────────────────────────────────────────
    lot_id = input(
        'Q0: 検索できなかった指図NOを入力 (空Enterでスキップ): '
    ).strip()
    if lot_id:
        # Q0a: V_仕掛指図別未出庫部品一覧に存在するか（単純WHEREなので高速）
        run(conn, '''
SELECT COUNT(*) AS 行数, COUNT(DISTINCT 工程コード) AS 工程数
FROM V_仕掛指図別未出庫部品一覧
WHERE 指図NO = :1
''',
            params=[lot_id],
            label=f'Q0a: {lot_id} が V_仕掛指図別未出庫部品一覧 に存在するか')

        # Q0b: V_ASP_生産計画一覧のステータス一覧（指図NO指定なので高速）
        run(conn, '''
SELECT 工程CD, 工程名, 工順, ステータス, 品目
FROM V_ASP_生産計画一覧
WHERE 指図NO = :1
ORDER BY 工順
''',
            params=[lot_id],
            label=f'Q0b: {lot_id} の V_ASP ステータス一覧')

        # Q0c: NMRP登録状況（品目マスタ）
        run(conn, '''
SELECT P.品目, N.PLANT, N.OLD_PRODUCT_CODE
FROM (
    SELECT DISTINCT 品目 FROM V_ASP_生産計画一覧 WHERE 指図NO = :1
) P
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT N ON 4235962023
.PRODUCT_CODE = P.品目
ORDER BY N.PLANT NULLS LAST
''',
            params=[lot_id],
            label=f'Q0c: {lot_id} の品目 NMRP 登録状況（全PLANT）')

    # ──────────────────────────────────────────────────────────
    # Q1: V_ASP_生産計画一覧 のステータス種別内訳
    #     ※ 単一ビューのGROUP BYのみ → 高速
    #     ここで '着手' 以外のステータスが多ければ、そちらも検索対象にする必要がある
    # ──────────────────────────────────────────────────────────
    run(conn, '''
SELECT ステータス, COUNT(DISTINCT 指図NO) AS 指図件数
FROM V_ASP_生産計画一覧
GROUP BY ステータス
ORDER BY 指図件数 DESC
''',
        label='Q1: V_ASP_生産計画一覧 のステータス別 指図件数（全件）')

    # ──────────────────────────────────────────────────────────
    # Q2: 着手ステータス指図 × NMRP結合状況（PLANT=5501）
    #     ★NMRP未登録が多い場合 → PLANT フィルターを外す or LEFT JOINのまま
    #     ※ V_ASP単体 + NHK_MASTER JOIN → タイムアウトしにくい
    # ──────────────────────────────────────────────────────────
    run(conn, '''
SELECT
    CASE
        WHEN N.OLD_PRODUCT_CODE IS NULL
            THEN '★NMRP未登録（5501でDB図番なし）'
        WHEN LENGTH(TRIM(N.OLD_PRODUCT_CODE)) <= 2
            THEN '★OLD_PRODUCT_CODE 短すぎ'
        ELSE '正常（DB図番あり）'
    END AS NMRP状態,
    COUNT(DISTINCT P.指図NO) AS 指図件数
FROM V_ASP_生産計画一覧 P
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT N
    ON  N.PRODUCT_CODE = P.品目
    AND N.PLANT        = '5501'
WHERE P.ステータス = '着手'
GROUP BY
    CASE
        WHEN N.OLD_PRODUCT_CODE IS NULL
            THEN '★NMRP未登録（5501でDB図番なし）'
        WHEN LENGTH(TRIM(N.OLD_PRODUCT_CODE)) <= 2
            THEN '★OLD_PRODUCT_CODE 短すぎ'
        ELSE '正常（DB図番あり）'
    END
ORDER BY 指図件数 DESC
''',
        label='Q2: 着手指図 × NMRP結合状況（PLANT=5501）')

    # ──────────────────────────────────────────────────────────
    # Q3: 着手指図のサンプル10件（DB図番・品目コード確認用）
    # ──────────────────────────────────────────────────────────
    run(conn, '''
SELECT
    P.指図NO,
    MAX(P.品目)           AS 品目コード,
    MAX(N.OLD_PRODUCT_CODE) AS DB図番_5501
FROM V_ASP_生産計画一覧 P
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT N
    ON  N.PRODUCT_CODE = P.品目
    AND N.PLANT        = '5501'
    AND LENGTH(TRIM(N.OLD_PRODUCT_CODE)) > 2
WHERE P.ステータス = '着手'
GROUP BY P.指図NO
ORDER BY P.指図NO
FETCH FIRST 10 ROWS ONLY
''',
        label='Q3: 着手指図サンプル10件（品目コード・DB図番）')

    # ──────────────────────────────────────────────────────────
    # Q4: DB図番キーワードで着手指図を検索（現行SQLの再現）
    # ──────────────────────────────────────────────────────────
    keyword = input(
        'Q4: 確認したいDB図番キーワードを入力 (例: DB24061, 空Enterでスキップ): '
    ).strip()
    if keyword:
        # 4a: PLANT=5501 固定（現行アプリと同条件）
        run(conn, '''
SELECT
    P.指図NO,
    MAX(P.品目)             AS 品目コード,
    MAX(N.OLD_PRODUCT_CODE) AS DB図番_5501
FROM V_ASP_生産計画一覧 P
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT N
    ON  N.PRODUCT_CODE = P.品目
    AND N.PLANT        = '5501'
    AND LENGTH(TRIM(N.OLD_PRODUCT_CODE)) > 2
WHERE P.ステータス = '着手'
  AND UPPER(N.OLD_PRODUCT_CODE) LIKE UPPER(:1)
GROUP BY P.指図NO
ORDER BY P.指図NO
FETCH FIRST 30 ROWS ONLY
''',
            params=[f'%{keyword}%'],
            label=f'Q4a: PLANT=5501 着手 DB図番 "{keyword}"')

        # 4b: PLANT問わず（DB図番がどのPLANTに登録されているか確認）
        run(conn, '''
SELECT
    P.指図NO,
    MAX(P.品目)             AS 品目コード,
    MAX(N.OLD_PRODUCT_CODE) AS DB図番,
    MAX(N.PLANT)            AS PLANT
FROM V_ASP_生産計画一覧 P
LEFT JOIN NHK_MASTER_RECEIVE_PRODUCT N
    ON  N.PRODUCT_CODE = P.品目
    AND LENGTH(TRIM(N.OLD_PRODUCT_CODE)) > 2
WHERE P.ステータス = '着手'
  AND UPPER(N.OLD_PRODUCT_CODE) LIKE UPPER(:1)
GROUP BY P.指図NO
ORDER BY P.指図NO
FETCH FIRST 30 ROWS ONLY
''',
            params=[f'%{keyword}%'],
            label=f'Q4b: 全PLANT 着手 DB図番 "{keyword}"')

    conn.close()
    print('診断完了。')


if __name__ == '__main__':
    main()
