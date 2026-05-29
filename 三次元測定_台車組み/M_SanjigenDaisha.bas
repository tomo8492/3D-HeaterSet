Attribute VB_Name = "M_SanjigenDaisha"
'==================================================================================
'  三次元測定 処理前 台車組み出力ツール (Excel VBA)
'  ----------------------------------------------------------------------------
'  「プロセスマネージャー」(oracledb + TNS で Oracle に直接接続する Python アプリ)
'  が参照しているのと同じビュー／テーブルを ADODB で直接参照し、
'  三次元測定の「処理前(到着済・未処理 / STEP=100)」の指図を抽出して
'  台車組み(同一品番・1台車2台まで・三次元測定開始日が早い順)を行い
'  「出力」シートへ書き出す。
'
'  使い方:
'    1) マクロ有効ブック(.xlsm)へ本モジュールを取り込む(コピペ推奨。下記注意参照)
'    2) シート「設定」に接続情報(TNS別名・ユーザー・パスワード)を入力
'    3) シート「対象品番」A列に DB図番 を列挙(後から追加するだけで対象が増える)
'    4) 実行_三次元測定台車組み を実行(図形ボタンに割り当て推奨)
'
'  ※ 文字コード注意:
'     本ファイルは UTF-8 です。日本語環境で File>Import すると日本語が文字化け
'     する場合があります。その場合はテキストとして開いてコードをコピーし、
'     VBE の標準モジュールへ「貼り付け」てください(貼り付けは Unicode を保持)。
'
'  ※ ビット数注意:
'     OraOLEDB.Oracle (Oracle Provider for OLE DB) は Excel と同じビット数
'     (32bit / 64bit)のものをインストールしておく必要があります。
'==================================================================================
Option Explicit

' ---- シート名 ----------------------------------------------------------------
Private Const SHEET_SETTING As String = "設定"
Private Const SHEET_TARGET  As String = "対象品番"
Private Const SHEET_OUTPUT  As String = "出力"

' ---- 既定値(設定シートが空欄のときに使用) -----------------------------------
Private Const DEFAULT_PROVIDER As String = "OraOLEDB.Oracle"
Private Const DEFAULT_KEYWORD  As String = "三次元測定"
Private Const DEFAULT_PLANT    As String = "5501"      ' 伊勢原
Private Const CMD_TIMEOUT      As Long = 300           ' 秒(重いビュー対策)
Private Const CART_CAPACITY    As Long = 2             ' 1台車に載る台数

' ---- ログ ---------------------------------------------------------------------
'   True にすると処理状況を VBE の「イミディエイト ウィンドウ」(Ctrl+G)へ出力。
'   生成SQL・接続結果(パスワードはマスク)・件数・処理時間・エラー詳細が見られる。
Private Const DEBUG_LOG As Boolean = True

' ---- 出力列(0始まり) --------------------------------------------------------
'   0:指図 1:DB図番 2:現工程 3:次工程開始日 4:品目コード 5:フロー名称
'   6:前工程 7:前工程着手状況
Private Const FLD_COUNT As Long = 8


'==================================================================================
' エントリポイント(日本語名・ASCII名の両方を用意)
'==================================================================================
Public Sub 実行_三次元測定台車組み()
    RunSanjigenDaisha
End Sub

Public Sub RunSanjigenDaisha()
    Dim provider As String, dataSource As String, uid As String, pwd As String
    Dim schemaPrefix As String, keyword As String, plant As String
    Dim targets As Collection
    Dim inList As String
    Dim sql As String
    Dim rows As Collection      ' 抽出結果(着手予定日 昇順)
    Dim t0 As Double: t0 = Timer

    On Error GoTo EH

    LogMsg String$(70, "=")
    LogMsg "▼ 三次元測定 台車組み 開始"

    ' --- 各シートが無ければ雛形を作成 ---
    EnsureSheets

    ' --- 設定読み込み ---
    provider = GetSetting("プロバイダ", DEFAULT_PROVIDER)
    dataSource = GetSetting("データソース", "")
    uid = GetSetting("ユーザーID", "")
    pwd = GetSetting("パスワード", "")
    schemaPrefix = GetSetting("スキーマ接頭辞", "")
    keyword = GetSetting("工程名キーワード", DEFAULT_KEYWORD)
    plant = GetSetting("プラント", DEFAULT_PLANT)
    If Len(Trim$(keyword)) = 0 Then keyword = DEFAULT_KEYWORD
    If Len(Trim$(plant)) = 0 Then plant = DEFAULT_PLANT
    If Len(Trim$(provider)) = 0 Then provider = DEFAULT_PROVIDER

    LogMsg "設定: プロバイダ=" & provider & " / データソース=" & dataSource & _
           " / ユーザー=" & uid & " / パスワード=" & MaskPwd(pwd)
    LogMsg "設定: スキーマ接頭辞=[" & schemaPrefix & "] / 工程名キーワード=[" & keyword & _
           "] / プラント=" & plant

    If Len(Trim$(dataSource)) = 0 Then
        LogMsg "中断: データソース未入力"
        MsgBox "シート「" & SHEET_SETTING & "」にデータソース(TNS別名 または EZConnect" & vbCrLf & _
               "例: host:1521/service)を入力してください。", vbExclamation, "設定不足"
        ThisWorkbook.Sheets(SHEET_SETTING).Activate
        Exit Sub
    End If

    ' --- 対象品番(DB図番)読み込み ---
    Set targets = ReadTargets()
    LogMsg "対象品番: " & targets.Count & " 件 → " & JoinCollection(targets, ", ")
    If targets.Count = 0 Then
        LogMsg "中断: 対象品番なし"
        MsgBox "シート「" & SHEET_TARGET & "」のA列に対象の DB図番 を入力してください。" & vbCrLf & _
               "(2行目以降。例: DB11048)", vbExclamation, "対象品番なし"
        ThisWorkbook.Sheets(SHEET_TARGET).Activate
        Exit Sub
    End If
    inList = BuildInList(targets)

    ' --- SQL 組み立て & 実行 ---
    sql = BuildSql(inList, schemaPrefix, keyword, plant)
    LogBlock "生成SQL", sql
    Set rows = QueryRows(provider, dataSource, uid, pwd, sql)
    LogMsg "抽出結果: " & rows.Count & " 件"

    If rows.Count = 0 Then
        ' 出力はクリアしておく(戻り値は使わない)
        Call WriteResults(New Collection, keyword, plant)
        LogMsg "該当なし(0件)。処理時間 " & Format$(Timer - t0, "0.0") & " 秒"
        MsgBox "対象(三次元測定・処理前)に該当する指図はありませんでした。" & vbCrLf & _
               "工程名キーワード=「" & keyword & "」 / プラント=" & plant, _
               vbInformation, "該当なし"
        Exit Sub
    End If

    ' --- 出力(台車組み) ---
    Dim cartCount As Long
    cartCount = WriteResults(rows, keyword, plant)

    ThisWorkbook.Sheets(SHEET_OUTPUT).Activate
    LogMsg "▲ 完了: 指図 " & rows.Count & " 件 / 台車 " & cartCount & " 台 / " & _
           Format$(Timer - t0, "0.0") & " 秒"
    MsgBox "完了しました。" & vbCrLf & _
           "対象指図: " & rows.Count & " 件 / 台車: " & cartCount & " 台" & vbCrLf & _
           "処理時間: " & Format$(Timer - t0, "0.0") & " 秒" & vbCrLf & vbCrLf & _
           "詳細ログは VBE のイミディエイト ウィンドウ(Ctrl+G)を確認してください。", _
           vbInformation, "三次元測定 台車組み"
    Exit Sub

EH:
    LogMsg "★エラー [" & Err.Number & "] " & Err.Description & _
           IIf(Len(Err.Source) > 0, " (Source: " & Err.Source & ")", "")
    If Len(sql) > 0 Then LogBlock "失敗時SQL", sql
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "[" & Err.Number & "] " & Err.Description & vbCrLf & vbCrLf & _
           "詳細は VBE のイミディエイト ウィンドウ(Ctrl+G)を確認してください。", _
           vbCritical, "エラー"
End Sub


'==================================================================================
' SQL 組み立て
'   ・targets : MES上で「現工程=三次元測定 かつ 処理前(STEP=100)」の対象指図
'   ・plan    : V_ASP_生産計画一覧から、三次元測定行の着手予定日(=次工程開始日)と
'               直前工程(前工程)・そのステータス(前工程着手状況)を取得
'==================================================================================
Private Function BuildSql(ByVal inList As String, ByVal schemaPrefix As String, _
                          ByVal keyword As String, ByVal plant As String) As String
    Dim vasp As String
    Dim kw As String
    Dim p As String
    vasp = schemaPrefix & "V_ASP_生産計画一覧"
    kw = Replace(keyword, "'", "''")        ' 念のためエスケープ
    p = Replace(plant, "'", "''")

    Dim s As String
    s = "WITH targets AS (" & vbCrLf
    s = s & "  SELECT MLI.LOT_ID, MLI.PRODUCT_CODE, NMRP.OLD_PRODUCT_CODE AS DB_NO," & vbCrLf
    s = s & "         MMPL.PROCESS_NAME AS CUR_PROC, MMFL.FLOW_NAME AS FLOW_NAME" & vbCrLf
    s = s & "  FROM MES_LOT_INFORMATION MLI" & vbCrLf
    s = s & "  LEFT JOIN MES_MASTER_PROCESS_LC MMPL" & vbCrLf
    s = s & "    ON MMPL.PROCESS_CODE = MLI.PROCESS_CODE AND MMPL.LOCATION_CODE = MLI.LOCATION_CODE" & vbCrLf
    s = s & "  INNER JOIN NHK_MASTER_RECEIVE_PRODUCT NMRP" & vbCrLf
    s = s & "    ON NMRP.PRODUCT_CODE = MLI.PRODUCT_CODE AND NMRP.PLANT = MLI.LOCATION_CODE" & vbCrLf
    s = s & "   AND (NMRP.PLANT, NMRP.PRODUCT_CODE, NMRP.UP_DATE) IN" & vbCrLf
    s = s & "       (SELECT PLANT, PRODUCT_CODE, MAX(UP_DATE) FROM NHK_MASTER_RECEIVE_PRODUCT GROUP BY PLANT, PRODUCT_CODE)" & vbCrLf
    s = s & "  LEFT JOIN MES_MASTER_FLOW_LC MMFL ON MMFL.FLOW_CODE = MLI.FLOW_CODE" & vbCrLf
    s = s & "  WHERE MLI.LOCATION_CODE = '" & p & "'" & vbCrLf
    s = s & "    AND MLI.STATE_FLG NOT IN ('E','S')" & vbCrLf            ' 完成/廃棄を除外
    s = s & "    AND MLI.STEP = '100'" & vbCrLf                          ' 処理前(到着済・未処理)
    s = s & "    AND MMPL.PROCESS_NAME LIKE '%" & kw & "%'" & vbCrLf      ' 現工程=三次元測定
    s = s & "    AND NMRP.OLD_PRODUCT_CODE IN (" & inList & ")" & vbCrLf
    s = s & ")," & vbCrLf
    s = s & "plan AS (" & vbCrLf
    s = s & "  SELECT 指図NO, 着手予定日, PREV_PROC, PREV_STATUS," & vbCrLf
    s = s & "         ROW_NUMBER() OVER (PARTITION BY 指図NO ORDER BY 工順) AS RN" & vbCrLf
    s = s & "  FROM (" & vbCrLf
    s = s & "    SELECT 指図NO, 工順, 工程名, 着手予定日," & vbCrLf
    s = s & "           LAG(工程名)   OVER (PARTITION BY 指図NO ORDER BY 工順) AS PREV_PROC," & vbCrLf
    s = s & "           LAG(ステータス) OVER (PARTITION BY 指図NO ORDER BY 工順) AS PREV_STATUS" & vbCrLf
    s = s & "    FROM " & vasp & vbCrLf
    s = s & "    WHERE 指図NO IN (SELECT LOT_ID FROM targets)" & vbCrLf
    s = s & "  )" & vbCrLf
    s = s & "  WHERE 工程名 LIKE '%" & kw & "%'" & vbCrLf                ' 三次元測定行に限定
    s = s & ")" & vbCrLf
    s = s & "SELECT" & vbCrLf
    s = s & "  t.LOT_ID                    AS 指図," & vbCrLf
    s = s & "  t.DB_NO                     AS DB図番," & vbCrLf
    s = s & "  t.CUR_PROC                  AS 現工程," & vbCrLf
    s = s & "  p.着手予定日                 AS 次工程開始日," & vbCrLf
    s = s & "  t.PRODUCT_CODE              AS 品目コード," & vbCrLf
    s = s & "  t.FLOW_NAME                 AS フロー名称," & vbCrLf
    s = s & "  p.PREV_PROC                 AS 前工程," & vbCrLf
    s = s & "  NVL(p.PREV_STATUS, '-')     AS 前工程着手状況" & vbCrLf
    s = s & "FROM targets t" & vbCrLf
    s = s & "LEFT JOIN plan p ON p.指図NO = t.LOT_ID AND p.RN = 1" & vbCrLf
    s = s & "ORDER BY p.着手予定日 ASC NULLS LAST, t.DB_NO, t.LOT_ID"

    BuildSql = s
End Function


'==================================================================================
' ADODB で接続し、結果を「行(Variant配列)のコレクション」として返す(着手予定日 昇順)
'==================================================================================
Private Function QueryRows(ByVal provider As String, ByVal dataSource As String, _
                           ByVal uid As String, ByVal pwd As String, _
                           ByVal sql As String) As Collection
    Dim conn As Object, rs As Object
    Dim cs As String
    Dim out As Collection
    Set out = New Collection

    cs = "Provider=" & provider & ";Data Source=" & dataSource & ";"
    If Len(uid) > 0 Then cs = cs & "User Id=" & uid & ";"
    If Len(pwd) > 0 Then cs = cs & "Password=" & pwd & ";"

    Dim tc As Double: tc = Timer
    LogMsg "接続文字列: " & MaskConn(cs)
    LogMsg "接続中... (CommandTimeout=" & CMD_TIMEOUT & "秒)"

    Set conn = CreateObject("ADODB.Connection")
    conn.CommandTimeout = CMD_TIMEOUT
    conn.Open cs
    LogMsg "接続成功 (" & Format$(Timer - tc, "0.0") & " 秒)。クエリ実行中..."

    Dim tq As Double: tq = Timer
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 0, 1   ' adOpenForwardOnly, adLockReadOnly
    LogMsg "クエリ実行完了 (" & Format$(Timer - tq, "0.0") & " 秒)。取得中..."

    Dim n As Long: n = 0
    Do While Not rs.EOF
        Dim r() As Variant
        ReDim r(0 To FLD_COUNT - 1)
        Dim i As Long
        For i = 0 To FLD_COUNT - 1
            r(i) = rs.Fields(i).Value
        Next i
        out.Add r
        n = n + 1
        ' 先頭数件は中身もログ(動作確認用)
        If n <= 5 Then
            LogMsg "  行" & n & ": 指図=" & NzStr(r(0)) & " / DB図番=" & NzStr(r(1)) & _
                   " / 現工程=" & NzStr(r(2)) & " / 次工程開始日=" & NzStr(r(3)) & _
                   " / 前工程=" & NzStr(r(6)) & " / 前工程着手状況=" & NzStr(r(7))
        End If
        rs.MoveNext
    Loop
    LogMsg "取得完了: " & n & " 行 (" & Format$(Timer - tc, "0.0") & " 秒)"

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    Set QueryRows = out
End Function


'==================================================================================
' 台車組み + 出力
'   ・rows は着手予定日 昇順で渡される前提
'   ・同一 DB図番 で 2台ずつ台車に割り当て(早い順)。端数は1台。
'   ・台車Noは生成順(=着手予定日が早い順)になるので、台車Noで並べると
'     「早い順・同一品番ペア」で整列する。
'   戻り値: 台車数
'==================================================================================
Private Function WriteResults(ByVal rows As Collection, ByVal keyword As String, _
                              ByVal plant As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_OUTPUT)

    Application.ScreenUpdating = False

    ' --- クリア ---
    ws.Cells.Clear

    ' --- 見出し ---
    Dim headers As Variant
    headers = Array("台車No", "指図", "DB図番", "現工程", "次工程開始日", _
                    "品目コード", "フロー名称", "前工程", "前工程着手状況")
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    ' 抽出条件メモ(右側)
    ws.Cells(1, UBound(headers) + 3).Value = _
        "抽出: 三次元測定 処理前 / キーワード=" & keyword & " / プラント=" & plant & _
        " / " & Format$(Now, "yyyy/mm/dd hh:nn")

    If rows.Count = 0 Then
        Application.ScreenUpdating = True
        WriteResults = 0
        Exit Function
    End If

    ' --- 台車割り当て ---
    Dim openCart As Object   ' DB図番 -> 現在の(未満杯)台車No
    Dim cartFill As Object   ' 台車No -> 積載数
    Dim cartRows As Object   ' 台車No -> 行インデックスの Collection
    Set openCart = CreateObject("Scripting.Dictionary")
    Set cartFill = CreateObject("Scripting.Dictionary")
    Set cartRows = CreateObject("Scripting.Dictionary")

    Dim nextCart As Long: nextCart = 0
    Dim idx As Long
    Dim db As String
    Dim cartNo As Long

    For idx = 1 To rows.Count
        Dim row As Variant
        row = rows(idx)
        db = NzStr(row(1))   ' DB図番

        If openCart.Exists(db) Then
            cartNo = openCart(db)
        Else
            nextCart = nextCart + 1
            cartNo = nextCart
            openCart(db) = cartNo
            cartFill(cartNo) = 0
            Set cartRows(cartNo) = New Collection
        End If

        cartFill(cartNo) = cartFill(cartNo) + 1
        cartRows(cartNo).Add idx
        If cartFill(cartNo) >= CART_CAPACITY Then openCart.Remove db  ' 満杯→締切
    Next idx

    ' 端数(1台のみ)の台車数をログ
    Dim partial As Long: partial = 0
    Dim kk As Variant
    For Each kk In cartFill.Keys
        If cartFill(kk) < CART_CAPACITY Then partial = partial + 1
    Next kk
    LogMsg "台車割当: " & nextCart & " 台 (満載 " & (nextCart - partial) & _
           " 台 / 端数1台 " & partial & " 台)"

    ' --- 台車No順に出力 ---
    Dim outRow As Long: outRow = 2
    Dim cn As Long
    For cn = 1 To nextCart
        Dim col As Collection
        Set col = cartRows(cn)
        Dim k As Long
        For k = 1 To col.Count
            Dim rr As Variant
            rr = rows(col(k))
            ws.Cells(outRow, 1).Value = cn
            ws.Cells(outRow, 2).Value = NzStr(rr(0))      ' 指図
            ws.Cells(outRow, 3).Value = NzStr(rr(1))      ' DB図番
            ws.Cells(outRow, 4).Value = NzStr(rr(2))      ' 現工程
            WriteDateCell ws.Cells(outRow, 5), rr(3)      ' 次工程開始日
            ws.Cells(outRow, 6).Value = NzStr(rr(4))      ' 品目コード
            ws.Cells(outRow, 7).Value = NzStr(rr(5))      ' フロー名称
            ws.Cells(outRow, 8).Value = NzStr(rr(6))      ' 前工程
            ws.Cells(outRow, 9).Value = NzStr(rr(7))      ' 前工程着手状況
            ' 台車ごとに薄く色分け(視認性)
            If (cn Mod 2) = 0 Then
                ws.Range(ws.Cells(outRow, 1), ws.Cells(outRow, 9)).Interior.Color = RGB(222, 235, 247)
            End If
            outRow = outRow + 1
        Next k
    Next cn

    ' --- 体裁 ---
    ws.Columns("A:I").AutoFit
    ws.Range("A1:I1").AutoFilter
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    Application.ScreenUpdating = True
    WriteResults = nextCart
End Function


'==================================================================================
' 補助関数群
'==================================================================================

' 対象品番シートのA列(2行目以降)から DB図番 を読み取り(重複・空白除去・大文字化)
Private Function ReadTargets() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_TARGET)
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim out As Collection
    Set out = New Collection

    Dim r As Long, v As String
    For r = 2 To last
        v = UCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        v = SanitizeCode(v)
        If Len(v) > 0 Then
            If Not seen.Exists(v) Then
                seen(v) = True
                out.Add v
            End If
        End If
    Next r
    Set ReadTargets = out
End Function

' DB図番として安全な文字(英数字とハイフン)のみ残す
Private Function SanitizeCode(ByVal s As String) As String
    Dim i As Long, ch As String, res As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "-" Then
            res = res & ch
        End If
    Next i
    SanitizeCode = res
End Function

' IN句リスト 'AAA','BBB',... を生成
Private Function BuildInList(ByVal targets As Collection) As String
    Dim parts() As String
    ReDim parts(0 To targets.Count - 1)
    Dim i As Long
    For i = 1 To targets.Count
        parts(i - 1) = "'" & targets(i) & "'"
    Next i
    BuildInList = Join(parts, ",")
End Function

' 設定シートの値を取得(ラベルはA列、値はB列)。見つからない/空欄なら既定値。
Private Function GetSetting(ByVal label As String, ByVal defaultValue As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_SETTING)
    Dim last As Long
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To last
        If Trim$(CStr(ws.Cells(r, 1).Value)) = label Then
            Dim v As String
            v = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(v) > 0 Then
                GetSetting = v
            Else
                GetSetting = defaultValue
            End If
            Exit Function
        End If
    Next r
    GetSetting = defaultValue
End Function

' Null/空を "-" に
Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Then
        NzStr = "-"
    Else
        NzStr = Trim$(CStr(v))
        If Len(NzStr) = 0 Then NzStr = "-"
    End If
End Function

' 日付セル書き込み(Date型なら yyyy/mm/dd 表示。Null/不正は "-")
Private Sub WriteDateCell(ByVal cell As Range, ByVal v As Variant)
    If IsNull(v) Then
        cell.Value = "-"
    ElseIf IsDate(v) Then
        cell.Value = CDate(v)
        cell.NumberFormatLocal = "yyyy/mm/dd"
    Else
        cell.Value = NzStr(v)
    End If
End Sub

' 必要シートが無ければ雛形を作成(設定の既定値・対象品番の初期15件を投入)
Private Sub EnsureSheets()
    EnsureSettingSheet
    EnsureTargetSheet
    EnsureOutputSheet
End Sub

Private Sub EnsureSettingSheet()
    If SheetExists(SHEET_SETTING) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = SHEET_SETTING
    ws.Range("A1").Value = "項目": ws.Range("B1").Value = "値"
    ws.Range("A1:B1").Font.Bold = True
    Dim items As Variant, vals As Variant
    items = Array("プロバイダ", "データソース", "ユーザーID", "パスワード", _
                  "スキーマ接頭辞", "工程名キーワード", "プラント")
    vals = Array(DEFAULT_PROVIDER, "", "", "", "", DEFAULT_KEYWORD, DEFAULT_PLANT)
    Dim i As Long
    For i = 0 To UBound(items)
        ws.Cells(i + 2, 1).Value = items(i)
        ws.Cells(i + 2, 2).Value = vals(i)
    Next i
    ws.Range("D2").Value = "■ 入力ガイド"
    ws.Range("D3").Value = "データソース: TNS別名 または EZConnect(例 host:1521/SERVICE)"
    ws.Range("D4").Value = "スキーマ接頭辞: 通常は空欄。必要時のみ 例 LV23_NHK. を入力"
    ws.Range("D5").Value = "工程名キーワード: 既定は 三次元測定 (部分一致)"
    ws.Range("D6").Value = "プラント: 伊勢原=5501 / 宮田=5401"
    ws.Range("D2").Font.Bold = True
    ws.Columns("A:B").AutoFit
End Sub

Private Sub EnsureTargetSheet()
    If SheetExists(SHEET_TARGET) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_TARGET
    ws.Range("A1").Value = "DB図番"
    ws.Range("A1").Font.Bold = True
    Dim codes As Variant
    codes = Array("DB11048", "DB13039", "DB11041", "DB23061", "DB17114", _
                  "DB17052", "DB21088", "DB18118", "DB20025", "DB20084", _
                  "DB10015", "DB12042", "DB18085", "DB03185", "DB24009")
    Dim i As Long
    For i = 0 To UBound(codes)
        ws.Cells(i + 2, 1).Value = codes(i)
    Next i
    ws.Range("C2").Value = "※ A列(2行目以降)に追加するだけで対象が増えます。"
    ws.Columns("A").AutoFit
End Sub

Private Sub EnsureOutputSheet()
    If SheetExists(SHEET_OUTPUT) Then Exit Sub
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = SHEET_OUTPUT
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(nm)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function


'==================================================================================
' ログ出力(イミディエイト ウィンドウ / Ctrl+G)
'==================================================================================
Private Sub LogMsg(ByVal msg As String)
    If DEBUG_LOG Then Debug.Print Format$(Now, "yyyy/mm/dd hh:nn:ss") & "  " & msg
End Sub

' 複数行(SQL等)を1行ずつ出力(Debug.Print の長文切り詰め対策)
Private Sub LogBlock(ByVal title As String, ByVal body As String)
    If Not DEBUG_LOG Then Exit Sub
    LogMsg title & ":"
    Dim lines() As String, i As Long
    lines = Split(body, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        Debug.Print "    " & lines(i)
    Next i
End Sub

' 接続文字列のパスワードをマスク
Private Function MaskConn(ByVal cs As String) As String
    Dim i As Long, j As Long
    i = InStr(1, cs, "Password=", vbTextCompare)
    If i = 0 Then MaskConn = cs: Exit Function
    j = InStr(i, cs, ";")
    If j = 0 Then j = Len(cs) + 1
    MaskConn = Left$(cs, i + Len("Password=") - 1) & "****" & Mid$(cs, j)
End Function

' パスワードの長さだけ伏字に(空なら (空))
Private Function MaskPwd(ByVal pwd As String) As String
    If Len(pwd) = 0 Then
        MaskPwd = "(空)"
    Else
        MaskPwd = String$(Len(pwd), "*")
    End If
End Function

' Collection を区切り文字で連結(ログ表示用)
Private Function JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim parts() As String
    If c.Count = 0 Then JoinCollection = "": Exit Function
    ReDim parts(0 To c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        parts(i - 1) = CStr(c(i))
    Next i
    JoinCollection = Join(parts, sep)
End Function
