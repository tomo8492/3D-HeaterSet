-- =====================================================================
-- 三次元測定 処理前(到着済・未処理 / STEP=100) 抽出 SQL
-- ---------------------------------------------------------------------
-- VBA(M_SanjigenDaisha.bas / BuildSql)が動的に組み立てる SQL の雛形。
-- 単体での動作確認・調整用に同じ内容をここに置く。
--
-- 接続(既存の本番マクロと同一):
--   Provider=OraOLEDB.Oracle; Data Source=LV2_PROD; User ID=LV23_NHK; Password=LV23_NHK;
--   ※接続ユーザー=LV23_NHK のため表/ビューは無修飾で参照可({SCHEMA}は空)。
--
-- バインド箇所(VBA が文字列置換する):
--   :PLANT   … プラント(例 5501)
--   :KEYWORD … 三次元測定の工程名キーワード(例 三次元測定)
--   :INLIST  … 対象 DB図番 の IN リスト(例 'DB11048','DB13039', ...)
--   {SCHEMA} … V_ASP_生産計画一覧 のスキーマ接頭辞(通常は空。例 LV23_NHK.)
--
-- データの考え方:
--   targets … MES_LOT_INFORMATION 上で「現在いる工程=三次元測定」かつ
--             STEP='100'(処理前=到着済・未処理)の対象指図。
--             現工程・DB図番・品目コード・フロー名称をここで確定。
--   plan    … V_ASP_生産計画一覧 から、三次元測定行の着手予定日(=次工程開始日)、
--             および工順が1つ手前の行(=前工程)とそのステータス(=前工程着手状況)。
--             LAG() で直前工程を取得し、三次元測定行(RN=1)に紐付ける。
--   並べ替え … 次工程開始日(三次元測定 着手予定日)の早い順。
--             台車組み(同一品番2台ずつ)は VBA 側で実施。
-- =====================================================================
WITH targets AS (
    SELECT MLI.LOT_ID,
           MLI.PRODUCT_CODE,
           NMRP.OLD_PRODUCT_CODE AS DB_NO,
           MMPL.PROCESS_NAME     AS CUR_PROC,
           MMFL.FLOW_NAME        AS FLOW_NAME
    FROM MES_LOT_INFORMATION MLI
    LEFT JOIN MES_MASTER_PROCESS_LC MMPL
        ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE
       AND MMPL.LOCATION_CODE = MLI.LOCATION_CODE
    INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP
        ON NMRP.PRODUCT_CODE = MLI.PRODUCT_CODE
       AND NMRP.PLANT = MLI.LOCATION_CODE
       AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE) IN (
            SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE)
            FROM NHK_MASTER_RECEIVE_PRODUCT
            GROUP BY PLANT, PRODUCT_CODE)
    LEFT JOIN MES_MASTER_FLOW_LC MMFL
        ON MMFL.FLOW_CODE = MLI.FLOW_CODE
    WHERE MLI.LOCATION_CODE = '5501'              -- :PLANT
      AND MLI.STATE_FLG NOT IN ('E','S')          -- 完成/廃棄を除外
      AND MLI.STEP = '100'                        -- 処理前(到着済・未処理)
      AND MMPL.PROCESS_NAME LIKE '%三次元測定%'    -- :KEYWORD (現工程=三次元測定)
      AND NMRP.OLD_PRODUCT_CODE IN (
            'DB11048','DB13039','DB11041','DB23061','DB17114',
            'DB17052','DB21088','DB18118','DB20025','DB20084',
            'DB10015','DB12042','DB18085','DB03185','DB24009'
          )                                       -- :INLIST
),
plan AS (
    SELECT 指図NO,
           着手予定日,
           PREV_PROC,
           PREV_STATUS,
           ROW_NUMBER() OVER (PARTITION BY 指図NO ORDER BY 工順) AS RN
    FROM (
        SELECT 指図NO,
               工順,
               工程名,
               着手予定日,
               LAG(工程名)     OVER (PARTITION BY 指図NO ORDER BY 工順) AS PREV_PROC,
               LAG(ステータス) OVER (PARTITION BY 指図NO ORDER BY 工順) AS PREV_STATUS
        FROM V_ASP_生産計画一覧                    -- {SCHEMA}V_ASP_生産計画一覧
        WHERE 指図NO IN (SELECT LOT_ID FROM targets)
    )
    WHERE 工程名 LIKE '%三次元測定%'                -- :KEYWORD (三次元測定行に限定)
)
SELECT
    TO_SINGLE_BYTE(t.LOT_ID) AS 指図,
    t.DB_NO                  AS DB図番,
    t.CUR_PROC               AS 現工程,
    p.着手予定日             AS 次工程開始日,
    t.PRODUCT_CODE           AS 品目コード,
    t.FLOW_NAME              AS フロー名称,
    p.PREV_PROC              AS 前工程,
    NVL(p.PREV_STATUS, '-')  AS 前工程着手状況
FROM targets t
LEFT JOIN plan p
    ON p.指図NO = t.LOT_ID
   AND p.RN = 1
ORDER BY p.着手予定日 ASC NULLS LAST, t.DB_NO, t.LOT_ID;
