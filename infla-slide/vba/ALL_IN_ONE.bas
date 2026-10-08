'==============================================================
' インフレスライド計算表 作成マクロ  ―  全部入り1ファイル版
'--------------------------------------------------------------
' このファイル1つを標準モジュールに貼り付ければ動きます。
' モジュール名（オブジェクト名）は何でも構いません。
'
' 手順
'   1) Excelを「マクロ有効ブック(.xlsm)」で保存する
'   2) Alt+F11 → 挿入 → 標準モジュール
'   3) このファイルを全部コピーして貼り付ける
'   4) Alt+F8 →「設定シート作成」→「スライド計算表作成」
'
' 作られるシート
'   ・スライド計算表（明細＋経費計算）
'   ・スライド調書（様式4-2号）… スライド計算表からリンク
'
' 新工種を考慮します（設定シートの「新工種を考慮する」＝する）。
' 考慮しない場合はその設定を「しない」に変えるか、新工種なし版を貼ってください。
'
' CSVの前提：①列＝変更設計、②列＝当初設計
' 金額は 単価×数量 を1円未満切り捨て（ROUNDDOWN）で計算します。
'
' そのほかのマクロ
'   ・特殊集計区分CSV取込 … 特殊集計区分一覧表CSVから処分費〇・管材費〇・処分費額を入れる
'   ・特殊集計区分CSV構造チェック … そのCSVの列の並びだけを書き出す（金額は出しません）
'   ・CSV構造チェック … CSVのコードと名称だけを書き出す（金額は出しません）
'   ・テスト用CSV作成 … 動作確認用のサンプルCSVを2本作る
'
' 参照設定の追加は不要です（すべて CreateObject の遅延バインディング）。
'==============================================================
Option Explicit

'==============================================================
' 宣言部
'==============================================================
'--- M00_Main の宣言 ---
Public Const SH_CONFIG As String = "設定"

' CSVの列番号（エスティマ「スライド用csv出力」／見出し行なし）
Public Const CSV_CODE   As Long = 1    ' コード
Public Const CSV_NAME   As Long = 2    ' 施工単価名称
Public Const CSV_TANI   As Long = 3    ' 単位
Public Const CSV_KIKAKU1 As Long = 4   ' 規格1
Public Const CSV_KIKAKU2 As Long = 5   ' 規格2
Public Const CSV_TEKIYO As Long = 6    ' 摘要
Public Const CSV_TANKA1 As Long = 7    ' 設計単価①
Public Const CSV_TANKA2 As Long = 8    ' 設計単価②
Public Const CSV_SURYO1 As Long = 9    ' 数量①
Public Const CSV_SURYO2 As Long = 10   ' 数量②
Public Const CSV_KIN1   As Long = 11   ' 設計金額①
Public Const CSV_KIN2   As Long = 12   ' 設計金額②

' 明細レコードの配列インデックス
Public Const R_CODE   As Long = 0
Public Const R_NAME   As Long = 1
Public Const R_TANI   As Long = 2
Public Const R_KIKAKU As Long = 3
Public Const R_TEKIYO As Long = 4
Public Const R_TANKA  As Long = 5
Public Const R_SURYO  As Long = 6
Public Const R_KIN    As Long = 7
Public Const R_LEVEL  As Long = 8      ' "費目" "L1".."L4" "" "Z" "YZ" "G"
Public Const R_SURYO2 As Long = 9      ' 採用しなかった側の数量（新工種の自動判定に使う）
Public Const R_COUNT  As Long = 10

'--- M10_CsvIO の宣言 ---
Public Const F_KOSHU    As Long = 0   ' 工種
Public Const F_SHUBETSU As Long = 1   ' 種別
Public Const F_SAIBETSU As Long = 2   ' 細別
Public Const F_KIKAKU   As Long = 3   ' 規格
Public Const F_TANI     As Long = 4   ' 単位
Public Const F_SURYO    As Long = 5   ' 数量
Public Const F_TANKA    As Long = 6   ' 単価
Public Const F_KINGAKU  As Long = 7   ' 金額
Public Const F_TEKIYO   As Long = 8   ' 摘要
Public Const F_COUNT    As Long = 9

'--- M20_Slide の宣言 ---
Public Const SC_CODE     As Long = 1    ' A  コード（作業用・非表示可）
Public Const SC_HIMOKU   As Long = 2    ' B  費目
Public Const SC_KUBUN    As Long = 3    ' C  工事区分
Public Const SC_KOSHU    As Long = 4    ' D  工種
Public Const SC_SHUBETSU As Long = 5    ' E  種別
Public Const SC_SAIBETSU As Long = 6    ' F  細別
Public Const SC_NAME     As Long = 7    ' G  施工単価名称
Public Const SC_KIKAKU   As Long = 8    ' H  規格
Public Const SC_TANKA_O  As Long = 9    ' I  単価 スライド前
Public Const SC_TANKA_N  As Long = 10   ' J  単価 スライド後
Public Const SC_Q_ALL    As Long = 11   ' K  数量 全体
Public Const SC_Q_DONE   As Long = 12   ' L  数量 出来形   ★手入力
Public Const SC_Q_REST   As Long = 13   ' M  数量 残工事   =K-L
Public Const SC_TANI     As Long = 14   ' N  単位
Public Const SC_MK_SHOBU As Long = 15   ' O  処分費〇      ★手入力
Public Const SC_MK_KANZA As Long = 16   ' P  管材費〇      ★手入力
Public Const SC_MK_SHIN  As Long = 17   ' Q  新工種〇      ★手入力
Public Const SC_ZKUBUN   As Long = 18   ' R  区分（経費計算部のZコード項目でだけ使う）
Public Const SC_A_ALL    As Long = 19   ' S  スライド前・全体      =K*I
Public Const SC_A_DONE   As Long = 20   ' T  スライド前・出来形    =L*I
Public Const SC_A_REST   As Long = 21   ' U  スライド前・残工事    =M*I
Public Const SC_B_ALL    As Long = 22   ' V  スライド後・全体      =K*J
Public Const SC_B_REST   As Long = 23   ' W  スライド後・残工事    =M*J
Public Const SC_N_ALL    As Long = 24   ' X  新工種抜き・全体      =IF(Q="〇",0,S)
Public Const SC_N_DONE   As Long = 25   ' Y  新工種抜き・出来形    =IF(Q="〇",0,T)
Public Const SC_TEKIYO   As Long = 26   ' Z  摘要
Public Const SC_SHA_O    As Long = 27   ' AA 処分費額（全数量分）スライド前 ★手入力
Public Const SC_SHA_N    As Long = 28   ' AB 処分費額（全数量分）スライド後 ★手入力
Public Const SC_SA_ALL   As Long = 29   ' AC 処分費額 前・全体
Public Const SC_SA_DONE  As Long = 30   ' AD 処分費額 前・出来形
Public Const SC_SA_REST  As Long = 31   ' AE 処分費額 前・残工事
Public Const SC_SB_ALL   As Long = 32   ' AF 処分費額 後・全体
Public Const SC_SB_REST  As Long = 33   ' AG 処分費額 後・残工事
Public Const SC_SN_ALL   As Long = 34   ' AH 処分費額 新抜き・全体
Public Const SC_SN_DONE  As Long = 35   ' AI 処分費額 新抜き・出来形
Public Const SC_LAST     As Long = 35

Public Const SH_SLIDE    As String = "スライド計算表"
Public Const SH_TANKA    As String = "_単価表"
Public Const FIRST_ROW   As Long = 8

Public Const CLR_INPUT   As Long = 65535          ' 黄色（手入力）
Public Const CLR_HEAD    As Long = 15849925       ' 薄い青
Public Const CLR_TOTAL   As Long = 49407          ' オレンジ

' ---- 他モジュールが参照する行 ----
Public gDirectFirst As Long     ' 直接工事費部の先頭行
Public gDirectLast  As Long     ' 直接工事費部の最終行
Public gZFirst      As Long     ' 共通仮設費 積上げ部の先頭行（0＝なし）
Public gZLast       As Long     ' 　　　　　　　　　　最終行
Public gZList       As Collection ' Zコード項目（Array(名称, コード, 先頭行, 最終行)）を出現順に
Public gDetailLast  As Long     ' 明細の最終行
Public gRowKakaku2  As Long     ' ㉑'工事価格（スクラップ込み）の行
Public gRowZei      As Long     ' ㉓消費税相当額の行
Public gRowKoujihi  As Long     ' ㉔工事費の行
Public gLastRow     As Long

' 経費計算部の作業用
Private mKWs As Worksheet
Private mKCfg As Object

' 新工種を自動判定するか
Private mAutoShin As Boolean

' 新工種を考慮するか（しないとき新工種の列・系列を出さない）
Public gUseShin As Boolean

'--- M40_Chosho の宣言 ---
Public Const SH_CHOSHO As String = "スライド調書（様式4-2号）"

'--- M50_Tokushu の宣言 ---
Private Const TK_TANI_LIST As String = "式|個|本|箇所|か所|ヶ所|カ所|枚|組|台|人|日|月|回|基|面|条|袋|缶|丁|冊|" & _
    "m|m2|m3|㎡|㎥|ｍ|ｍ2|ｍ3|km|cm|mm|t|kg|g|L|l|ℓ|ha|a|人日|台日|日当|空m3|区間|社|戸|世帯|箇|ケ|個所"

Public Const SH_TK_RESULT As String = "_特殊集計取込結果"
Public Const SH_TK_CHECK  As String = "_特殊集計CSV構造"

' 突合結果（計算表の行ごと）の配列の添字
Private Const A_KIND  As Long = 0     ' 判定（処分／管材）
Private Const A_BEF   As Long = 1     ' 金額（スライド前）の合計
Private Const A_AFT   As Long = 2     ' 金額（スライド後）の合計
Private Const A_CSVB  As Long = 3     ' 元になったCSVの行（前）
Private Const A_CSVA  As Long = 4     ' 元になったCSVの行（後）
Private Const A_NAMES As Long = 5     ' 処分費等の名称
Private Const A_NB    As Long = 6     ' 件数（前）
Private Const A_NA    As Long = 7     ' 件数（後）
Private Const A_NOTE  As Long = 8     ' 注意書き
Private Const A_KUBUN As Long = 9     ' 集計区分名称
Private Const A_CODE  As Long = 10    ' 突合に使ったコード
Private Const A_WHICH As Long = 11    ' どのコードで当てたか
Private Const A_HASB  As Long = 12    ' 前の金額が入っていたか
Private Const A_HASA  As Long = 13    ' 後の金額が入っていたか
Private Const A_COUNT As Long = 14

'--- M90_SampleData の宣言 ---
Private mSb As String


'==============================================================
' 《M00_Main》 メイン処理・CSV読み込み・2本の突合
'==============================================================


'==============================================================
' メイン
'==============================================================
Public Sub スライド計算表作成()
    Dim cfg As Object
    Dim pathOld As String, pathNew As String
    Dim recOld As Collection, recNew As Collection
    Dim labelOld As String, labelNew As String
    Dim warn As String

    On Error GoTo ErrHandler

    If Not SheetExists(SH_CONFIG) Then
        MsgBox "「" & SH_CONFIG & "」シートがありません。" & vbCrLf & _
               "先に［設定シート作成］を実行してください。", vbExclamation
        Exit Sub
    End If

    ' 古い設定シートに足りない項目を補ってから読む
    If 設定シート更新() > 0 Then
        MsgBox "「" & SH_CONFIG & "」シートに、このバージョンで増えた設定を追加しました。" & vbCrLf & _
               "赤字の項目を確認してから、もう一度実行してください。", vbInformation
        ThisWorkbook.Worksheets(SH_CONFIG).Activate
        Exit Sub
    End If

    Set cfg = GetConfig()

    pathOld = AskCsvPath(CStr(CfgVal(cfg, "旧単価CSVパス", "")), _
                         "【旧単価】当初の単価適用日で出力したCSVを選択してください")
    If pathOld = "" Then Exit Sub
    pathNew = AskCsvPath(CStr(CfgVal(cfg, "新単価CSVパス", "")), _
                         "【新単価】基準日の単価適用日で出力したCSVを選択してください")
    If pathNew = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "旧単価CSVを読み込んでいます…"
    Set recOld = LoadEstima(pathOld, cfg, labelOld)

    Application.StatusBar = "新単価CSVを読み込んでいます…"
    Set recNew = LoadEstima(pathNew, cfg, labelNew)

    If recOld.Count = 0 Then
        GoTo NoData
    End If

    Application.StatusBar = "スライド計算表を作成しています…"
    warn = BuildSlideSheet(cfg, recOld, recNew)

    Application.StatusBar = "スライド調書（様式4-2号）を作成しています…"
    BuildChosho cfg

    ThisWorkbook.Worksheets(SH_SLIDE).Activate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "2つのシートを作成しました。" & vbCrLf & _
           "　・スライド計算表（明細＋経費計算）" & vbCrLf & _
           "　・スライド調書（様式4-2号）" & vbCrLf & vbCrLf & _
           "旧単価CSV：" & recOld.Count & " 行（採用系列：" & labelOld & "）" & vbCrLf & _
           "新単価CSV：" & recNew.Count & " 行（採用系列：" & labelNew & "）" & vbCrLf & vbCrLf & _
           IIf(warn = "", "突合はすべて一致しました。", warn) & vbCrLf & vbCrLf & _
           "次に黄色セルを手入力してください。" & vbCrLf & _
           "【スライド計算表・明細部】" & vbCrLf & _
           "　・L列 出来形数量" & vbCrLf & _
           "　・O列 処分費〇／P列 管材費〇" & IIf(gUseShin, "／Q列 新工種〇", "（新工種は考慮しない設定）") & vbCrLf & _
           "　・AA列/AB列 処分費額（全数量分の金額。処分費と運搬費の合算額から処分費分だけ）" & vbCrLf & _
           "【スライド計算表・経費計算部】" & vbCrLf & _
           "　・⑦⑪⑮ 経費率（積算システムの値があれば上書き）" & vbCrLf & _
           "　・⑯契約保証費", vbInformation
    Exit Sub

NoData:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "明細を1行も読み取れませんでした。" & vbCrLf & _
           "CSVがエスティマの「スライド用csv出力」形式か確認してください。", vbExclamation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "内容：" & Err.Description & vbCrLf & "番号：" & Err.Number, vbCritical
End Sub


'==============================================================
' エスティマCSVの読み込み
'   ①列＝変更設計、②列＝当初設計。usedLabel に採用した側を返す
'==============================================================
Private Function LoadEstima(ByVal filePath As String, ByVal cfg As Object, _
                            ByRef usedLabel As String) As Collection
    Dim rows As Collection, recs As Collection
    Dim arr() As String
    Dim i As Long
    Dim useFirst As Boolean
    Dim seriesName As String
    Dim rec As Variant
    Dim code As String

    Set recs = New Collection
    Set rows = ParseCsvText( _
                    ReadTextFile(filePath, CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))), _
                    CStr(CfgVal(cfg, "区切り文字", ",")))

    If rows.Count = 0 Then
        Set LoadEstima = recs
        Exit Function
    End If

    '--- ①＝変更、②＝当初（積算システムの出力仕様）---
    seriesName = NormText(CStr(CfgVal(cfg, "使用系列", "変更")))
    If seriesName = NormText("当初") Then
        useFirst = False
        usedLabel = "当初"
    Else
        useFirst = True
        usedLabel = "変更"
    End If

    '--- 明細を組み立てる ---
    For i = 1 To rows.Count
        arr = rows(i)
        code = Trim$(ColVal(arr, CSV_CODE))

        ReDim rec(0 To R_COUNT - 1)
        rec(R_CODE) = code
        rec(R_NAME) = Trim$(ColVal(arr, CSV_NAME))
        rec(R_TANI) = Trim$(ColVal(arr, CSV_TANI))
        rec(R_KIKAKU) = JoinKikaku(Trim$(ColVal(arr, CSV_KIKAKU1)), Trim$(ColVal(arr, CSV_KIKAKU2)))
        rec(R_TEKIYO) = Trim$(ColVal(arr, CSV_TEKIYO))

        If useFirst Then
            rec(R_TANKA) = ToNum(ColVal(arr, CSV_TANKA1))
            rec(R_SURYO) = ToNum(ColVal(arr, CSV_SURYO1))
            rec(R_KIN) = ToNum(ColVal(arr, CSV_KIN1))
            rec(R_SURYO2) = ToNum(ColVal(arr, CSV_SURYO2))
        Else
            rec(R_TANKA) = ToNum(ColVal(arr, CSV_TANKA2))
            rec(R_SURYO) = ToNum(ColVal(arr, CSV_SURYO2))
            rec(R_KIN) = ToNum(ColVal(arr, CSV_KIN2))
            rec(R_SURYO2) = ToNum(ColVal(arr, CSV_SURYO1))
        End If

        rec(R_LEVEL) = LevelOf(code)

        ' 名称もコードも空の行は捨てる
        If Not (code = "" And CStr(rec(R_NAME)) = "") Then recs.Add rec
    Next i

    Set LoadEstima = recs
End Function


'--------------------------------------------------------------
' コードからレベルを判定する
'   （単価調整VBA ①様式整理レベル表記 と同じ規則）
'--------------------------------------------------------------
Public Function LevelOf(ByVal code As String) As String
    Dim c As String
    c = UCase$(Trim$(code))

    If c = "" Then Exit Function                      ' 明細行

    If c Like "*X1000*" Then LevelOf = "費目": Exit Function

    If c Like "*Y23*" Then
        Select Case Len(c)
            Case 5:  LevelOf = "L1"
            Case 7:  LevelOf = "L2"
            Case 9:  LevelOf = "L3"
            Case 11: LevelOf = "L4"
            Case Else: LevelOf = "L4"
        End Select
        Exit Function
    End If

    If c Like "*Y10*" Then LevelOf = "L1": Exit Function
    If c Like "*Y18*" Then LevelOf = "L2": Exit Function   ' 工L2
    If c Like "*Y20*" Then LevelOf = "L2": Exit Function
    If c Like "*Y30*" Then LevelOf = "L3": Exit Function
    If c Like "*Y40*" Then LevelOf = "L4": Exit Function

    If c Like "YZ*" Then LevelOf = "YZ": Exit Function
    If c Like "Z0*" Then LevelOf = "Z": Exit Function
    If c Like "G*" Then LevelOf = "G": Exit Function
End Function


Private Function JoinKikaku(ByVal k1 As String, ByVal k2 As String) As String
    If k1 = "" Then
        JoinKikaku = k2
    ElseIf k2 = "" Then
        JoinKikaku = k1
    Else
        JoinKikaku = k1 & "　" & k2
    End If
End Function


'--------------------------------------------------------------
' 突合キー（旧CSVと新CSVの行を対応づける）
'--------------------------------------------------------------
Public Function RecKey(ByVal rec As Variant) As String
    RecKey = NormText(CStr(rec(R_CODE))) & "|" & NormText(CStr(rec(R_NAME))) & "|" & _
             NormText(CStr(rec(R_KIKAKU))) & "|" & NormText(CStr(rec(R_TANI)))
End Function


'==============================================================
' CSVパスの決定
'==============================================================
Private Function AskCsvPath(ByVal presetPath As String, ByVal caption As String) As String
    Dim v As Variant

    If presetPath <> "" Then
        If Dir(presetPath) <> "" Then
            AskCsvPath = presetPath
            Exit Function
        End If
    End If

    v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , caption)
    If VarType(v) = vbBoolean Then Exit Function
    AskCsvPath = CStr(v)
End Function


'==============================================================
' 設定シートの読み取り
'==============================================================
Public Function GetConfig() As Object
    Dim ws As Worksheet, d As Object
    Dim r As Long, lastR As Long, k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 1 Then lastR = 1

    For r = 1 To lastR
        k = NormText(CStr(ws.Cells(r, 1).Value))
        If k <> "" Then
            If Not d.Exists(k) Then d.Add k, ws.Cells(r, 2).Value
        End If
    Next r

    Set GetConfig = d
End Function


Public Function CfgVal(ByVal d As Object, ByVal key As String, ByVal defaultValue As Variant) As Variant
    Dim k As String
    k = NormText(key)

    If Not d.Exists(k) Then
        CfgVal = defaultValue
        Exit Function
    End If

    If IsEmpty(d(k)) Or CStr(d(k)) = "" Then
        CfgVal = defaultValue
    Else
        CfgVal = d(k)
    End If
End Function


'==============================================================
' シート操作の共通処理
'==============================================================
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function


Public Function FreshSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    If SheetExists(sheetName) Then ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Set FreshSheet = ws
End Function


'==============================================================
' 《M10_CsvIO》 CSV読み書き・数値／文字列変換
'==============================================================


'--------------------------------------------------------------
' テキストファイルを指定文字コードで読み込む
'--------------------------------------------------------------
Public Function ReadTextFile(ByVal filePath As String, ByVal charSetName As String) As String
    Dim st As Object, s As String

    If charSetName = "" Then charSetName = "Shift_JIS"

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                      ' adTypeText
    st.Charset = charSetName
    st.Open
    st.LoadFromFile filePath
    s = st.ReadText(-1)              ' adReadAll
    st.Close

    ' UTF-8 BOM が残る場合があるので除去
    If Len(s) > 0 Then
        If Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    End If

    ReadTextFile = s
End Function


'--------------------------------------------------------------
' テキストファイルを指定文字コードで書き出す
'--------------------------------------------------------------
Public Sub WriteTextFile(ByVal filePath As String, ByVal s As String, ByVal charSetName As String)
    Dim st As Object

    If charSetName = "" Then charSetName = "Shift_JIS"

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                      ' adTypeText
    st.Charset = charSetName
    st.Open
    st.WriteText s
    st.SaveToFile filePath, 2        ' adSaveCreateOverWrite
    st.Close
End Sub

'--------------------------------------------------------------
' CSVテキストを解析して Collection(String配列) を返す
'   ・ダブルクォート囲み、"" によるエスケープ、囲み内の改行に対応
'--------------------------------------------------------------
Public Function ParseCsvText(ByVal s As String, ByVal delim As String) As Collection
    Dim rows As Collection
    Dim fields() As String
    Dim nField As Long
    Dim buf As String
    Dim inQuote As Boolean
    Dim i As Long, n As Long
    Dim ch As String

    If delim = "" Then delim = ","
    Set rows = New Collection

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    n = Len(s)

    ReDim fields(0 To 63)
    nField = 0
    i = 1

    Do While i <= n
        ch = Mid$(s, i, 1)

        If inQuote Then
            If ch = """" Then
                If i < n And Mid$(s, i + 1, 1) = """" Then
                    buf = buf & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                buf = buf & ch
            End If

        ElseIf ch = """" Then
            inQuote = True

        ElseIf ch = delim Then
            If nField > UBound(fields) Then ReDim Preserve fields(0 To nField + 63)
            fields(nField) = buf
            buf = ""
            nField = nField + 1

        ElseIf ch = vbLf Then
            If nField > UBound(fields) Then ReDim Preserve fields(0 To nField + 63)
            fields(nField) = buf
            rows.Add TrimFieldArray(fields, nField)
            buf = ""
            nField = 0
            ReDim fields(0 To 63)

        Else
            buf = buf & ch
        End If

        i = i + 1
    Loop

    ' 最終行（末尾に改行が無い場合）
    If nField > 0 Or Len(buf) > 0 Then
        If nField > UBound(fields) Then ReDim Preserve fields(0 To nField + 63)
        fields(nField) = buf
        rows.Add TrimFieldArray(fields, nField)
    End If

    Set ParseCsvText = rows
End Function


Private Function TrimFieldArray(ByRef fields() As String, ByVal lastIdx As Long) As String()
    Dim r() As String, k As Long
    ReDim r(0 To lastIdx)
    For k = 0 To lastIdx
        r(k) = fields(k)
    Next k
    TrimFieldArray = r
End Function


'--------------------------------------------------------------
' 配列から列番号(1始まり)の値を安全に取り出す。範囲外は空文字。
'--------------------------------------------------------------
Public Function ColVal(ByRef arr() As String, ByVal colNo As Long) As String
    If colNo <= 0 Then Exit Function
    If colNo - 1 > UBound(arr) Then Exit Function
    ColVal = arr(colNo - 1)
End Function


'--------------------------------------------------------------
' 文字列を数値へ（カンマ・円記号・全角数字・空白を吸収）
'--------------------------------------------------------------
Public Function ToNum(ByVal v As String) As Double
    Dim t As String

    t = Trim$(v)
    If t = "" Then Exit Function

    On Error Resume Next
    t = StrConv(t, vbNarrow)          ' 全角 → 半角
    On Error GoTo 0

    t = Replace(t, ",", "")
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")
    t = Replace(t, "\", "")
    t = Replace(t, "¥", "")
    t = Replace(t, "円", "")

    If t = "" Or t = "-" Then Exit Function
    If Not IsNumeric(t) Then Exit Function

    ToNum = CDbl(t)
End Function


'--------------------------------------------------------------
' 突合キー用の正規化（前後空白・空白文字を除去し半角化）
'--------------------------------------------------------------
Public Function NormText(ByVal s As String) As String
    Dim t As String

    t = Trim$(s)
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")

    On Error Resume Next
    t = StrConv(t, vbNarrow)
    On Error GoTo 0

    NormText = t
End Function


'--------------------------------------------------------------
' 金額の端数処理
'   mode : "切り捨て" / "切り上げ" / "四捨五入"
'   unit : 丸め単位（1 / 10 / 100 / 1000 …）
'--------------------------------------------------------------
Public Function RoundAmt(ByVal v As Double, ByVal mode As String, ByVal unit As Double) As Double
    If unit <= 0 Then unit = 1

    Select Case mode
        Case "切り上げ"
            RoundAmt = -Int(-v / unit) * unit
        Case "四捨五入"
            RoundAmt = Int(v / unit + 0.5) * unit
        Case Else                       ' 切り捨て
            RoundAmt = Int(v / unit) * unit
    End Select
End Function


'==============================================================
' 《M20_Slide》 設定シート生成・スライド計算表（明細＋経費計算）
'==============================================================


'==============================================================
' 設定シート
'==============================================================
Public Sub 設定シート作成()
    Dim ws As Worksheet
    Dim r As Long

    If SheetExists(SH_CONFIG) Then
        If MsgBox("「" & SH_CONFIG & "」シートを作り直します。入力済みの内容は消えます。" & _
                  vbCrLf & "よろしいですか？", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = FreshSheet(SH_CONFIG)

    With ws.Range("A1")
        .Value = "インフレスライド計算表 作成設定"
        .Font.Size = 14
        .Font.Bold = True
    End With

    r = 3
    PutHead ws, r, "【工事情報】":                                                     r = r + 1
    PutItem ws, r, "工事名", "":                                                        r = r + 1
    PutItem ws, r, "工事場所", "":                                                      r = r + 1
    PutItem ws, r, "工期（自）", "":                                                    r = r + 1
    PutItem ws, r, "工期（至）", "":                                                    r = r + 1
    PutItem ws, r, "請負代金額", 0, "税込。様式4-2号の請負率の算出に使います":          r = r + 1
    PutItem ws, r, "基準日", "", "インフレスライドの請求日":                            r = r + 1
    PutItem ws, r, "発注者", "神戸市":                                                  r = r + 1
    PutItem ws, r, "受注者", "":                                                        r = r + 2

    PutHead ws, r, "【CSV設定】", "エスティマ 機能→CSV連携→抽出指示ファイル「スライド用csv出力」": r = r + 1
    PutItem ws, r, "旧単価CSVパス", "", "当初の単価適用日で出力したCSV。空欄なら実行時に選択":   r = r + 1
    PutItem ws, r, "新単価CSVパス", "", "基準日の単価適用日で出力したCSV。空欄なら実行時に選択": r = r + 1
    PutItem ws, r, "文字コード", "Shift_JIS", "UTF-8の場合は UTF-8 と入力":              r = r + 1
    PutItem ws, r, "区切り文字", ",":                                                   r = r + 1
    PutItem ws, r, "使用系列", "変更", "変更／当初。CSVの①列＝変更設計、②列＝当初設計":      r = r + 1
    PutItem ws, r, "新工種を考慮する", DefaultUseShin(), _
            "する／しない。しないにすると新工種の列・系列を出さず、受注者負担1%の母数は残工事そのものになります": r = r + 1
    PutItem ws, r, "新工種の自動判定", "する", _
            "する／しない。採用しなかった側の数量が0で、採用側に数量がある明細にQ列の〇を付けます（考慮する場合のみ）": r = r + 2

    PutHead ws, r, "【特殊集計区分CSVの設定】", "エスティマの「特殊集計区分一覧表」CSV。処分費〇・管材費〇を自動で入れます": r = r + 1
    PutItem ws, r, "特殊集計区分CSVパス（スライド前）", "", "空欄なら実行時に選択":              r = r + 1
    PutItem ws, r, "特殊集計区分CSVパス（スライド後）", "", _
            "空欄なら実行時に選択。後の一覧表が無ければキャンセルすればAB列は空のまま":          r = r + 1
    PutItem ws, r, "特殊集計CSV列指定", "自動", _
            "自動／または 区分=14,コード=1,名称=2,単位=3,規格1=4,規格2=5,金額=11 のように列番号（A,Bでも可）": r = r + 1
    PutItem ws, r, "処分費の区分名", "処分", "特殊集計区分の名前にこの文字が入る行をO列の〇にします。カンマ区切りで複数可": r = r + 1
    PutItem ws, r, "管材費の区分名", "水道,管材,管財", "同じくP列の〇にします":              r = r + 1
    PutItem ws, r, "特殊集計CSVから処分費額も入れる", "する", _
            "する／しない。するとCSVの金額をAA列（処分費額・全数量分）に入れます":          r = r + 2

    PutHead ws, r, "【経費計算の設定】":                                                r = r + 1
    PutItem ws, r, "契約保証費", 0, "当初設計時の額で固定":                             r = r + 1
    PutItem ws, r, "処分費控除率", 0.03, "③＝②－ROUNDDOWN((①－①'/2)×この率,0)":     r = r + 1
    PutItem ws, r, "消費税率", 0.1:                                                     r = r + 2

    PutHead ws, r, "【経費率の自動計算】", "スライド計算表の率欄の初期値。積算システムの値で上書きできます": r = r + 1
    PutItem ws, r, "共通仮設費率A", 485.4, "率(%) = A × 対象額 ^ B":                    r = r + 1
    PutItem ws, r, "共通仮設費率B", -0.2231:                                            r = r + 1
    PutItem ws, r, "共通仮設費 地域補正", 1.5:                                          r = r + 1
    PutItem ws, r, "共通仮設費 週休補正", 1.01:                                         r = r + 1
    PutItem ws, r, "現場管理費率A", 202.3, "率(%) = A × 対象額 ^ B":                    r = r + 1
    PutItem ws, r, "現場管理費率B", -0.1034:                                            r = r + 1
    PutItem ws, r, "現場管理費 地域補正", 1.2:                                          r = r + 1
    PutItem ws, r, "現場管理費 週休補正", 1.02:                                         r = r + 1
    PutItem ws, r, "一般管理費率係数", -4.97802, "率(%) = 係数 × LOG10(対象額) + 定数": r = r + 1
    PutItem ws, r, "一般管理費率定数", 56.92101

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 58
    ws.Range("C:C").Font.Color = RGB(120, 120, 120)

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("B4").Select

    MsgBox "「" & SH_CONFIG & "」シートを作成しました。" & vbCrLf & _
           "工事情報を入力してから［スライド計算表作成］を実行してください。", vbInformation
End Sub


Private Sub PutHead(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    Optional ByVal note As String = "")
    With ws.Cells(r, 1)
        .Value = label
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
    End With
    If note <> "" Then ws.Cells(r, 3).Value = note
End Sub


Private Sub PutItem(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByVal defaultValue As Variant, Optional ByVal note As String = "")
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = defaultValue
    ws.Cells(r, 2).Interior.Color = CLR_INPUT
    ws.Cells(r, 2).Borders.LineStyle = xlContinuous
    If note <> "" Then ws.Cells(r, 3).Value = note
End Sub


'==============================================================
' スライド計算表を作る
'==============================================================
' 古いバージョンで作った設定シートに、後から増えた項目を補う
'   既にある項目は触りません。入力済みの値は消えません。
Public Function 設定シート更新() As Long
    Dim ws As Worksheet, cfg As Object
    Dim r As Long, added As Long

    If Not SheetExists(SH_CONFIG) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    Set cfg = GetConfig()
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    added = added + AddIfMissing(ws, cfg, r, "新工種を考慮する", DefaultUseShin(), _
        "する／しない。しないにすると新工種の列・系列を出しません")
    added = added + AddIfMissing(ws, cfg, r, "新工種の自動判定", "する", _
        "する／しない。当初数量0・変更数量ありの明細にQ列の〇を付けます")
    added = added + AddIfMissing(ws, cfg, r, "使用系列", "変更", _
        "変更／当初。CSVの①列＝変更設計、②列＝当初設計")
    added = added + AddIfMissing(ws, cfg, r, "特殊集計区分CSVパス（スライド前）", "", _
        "空欄なら実行時に選択")
    added = added + AddIfMissing(ws, cfg, r, "特殊集計区分CSVパス（スライド後）", "", _
        "空欄なら実行時に選択。無ければキャンセルすればAB列は空のまま")
    added = added + AddIfMissing(ws, cfg, r, "特殊集計CSV列指定", "自動", _
        "自動／または 区分=14,コード=1,名称=2,単位=3,規格1=4,規格2=5,金額=11 のように列番号")
    added = added + AddIfMissing(ws, cfg, r, "処分費の区分名", "処分", _
        "特殊集計区分の名前にこの文字が入る行をO列の〇にします")
    added = added + AddIfMissing(ws, cfg, r, "管材費の区分名", "水道,管材,管財", _
        "同じくP列の〇にします")
    added = added + AddIfMissing(ws, cfg, r, "特殊集計CSVから処分費額も入れる", "する", _
        "する／しない。CSVの金額をAA列（処分費額・全数量分）に入れます")
    added = added + AddIfMissing(ws, cfg, r, "契約保証費", 0, "当初設計時の額で固定")
    added = added + AddIfMissing(ws, cfg, r, "処分費控除率", 0.03, "")
    added = added + AddIfMissing(ws, cfg, r, "消費税率", 0.1, "")
    added = added + AddIfMissing(ws, cfg, r, "共通仮設費率A", 485.4, "")
    added = added + AddIfMissing(ws, cfg, r, "共通仮設費率B", -0.2231, "")
    added = added + AddIfMissing(ws, cfg, r, "共通仮設費 地域補正", 1.5, "")
    added = added + AddIfMissing(ws, cfg, r, "共通仮設費 週休補正", 1.01, "")
    added = added + AddIfMissing(ws, cfg, r, "現場管理費率A", 202.3, "")
    added = added + AddIfMissing(ws, cfg, r, "現場管理費率B", -0.1034, "")
    added = added + AddIfMissing(ws, cfg, r, "現場管理費 地域補正", 1.2, "")
    added = added + AddIfMissing(ws, cfg, r, "現場管理費 週休補正", 1.02, "")
    added = added + AddIfMissing(ws, cfg, r, "一般管理費率係数", -4.97802, "")
    added = added + AddIfMissing(ws, cfg, r, "一般管理費率定数", 56.92101, "")

    設定シート更新 = added
End Function


Private Function AddIfMissing(ByVal ws As Worksheet, ByVal cfg As Object, ByRef r As Long, _
                              ByVal label As String, ByVal defaultValue As Variant, _
                              ByVal note As String) As Long
    If cfg.Exists(NormText(label)) Then Exit Function

    r = r + 1
    If ws.Cells(r, 1).Value <> "" Then r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    PutItem ws, r, label, defaultValue, IIf(note = "", "（このバージョンで追加された設定）", note)
    ws.Cells(r, 1).Font.Color = RGB(192, 0, 0)
    AddIfMissing = 1
End Function


Public Function BuildSlideSheet(ByVal cfg As Object, ByVal recOld As Collection, _
                                ByVal recNew As Collection) As String
    Dim ws As Worksheet
    Dim newTanka() As Double
    Dim warnText As String
    Dim i As Long, r As Long, n As Long
    Dim rec As Variant
    Dim lvl As String
    Dim lvlNum As Long
    Dim rowArr() As Long, lvlArr() As Long
    Dim cnt As Long, firstZ As Long, firstX As Long, skippedG As Long
    Dim zCnt As Long, zRowArr() As Long, zLvlArr() As Long
    Dim curItem As String, curCode As String, curStart As Long

    n = recOld.Count
    ReDim newTanka(1 To n)
    warnText = MatchNewPrices(recOld, recNew, newTanka)

    Set ws = FreshSheet(SH_SLIDE)
    Set gZList = New Collection
    gUseShin = (NormText(CStr(CfgVal(cfg, "新工種を考慮する", DefaultUseShin()))) = "する")
    mAutoShin = gUseShin And (NormText(CStr(CfgVal(cfg, "新工種の自動判定", "する"))) = "する")
    WriteHeader ws, cfg

    ReDim rowArr(1 To n)
    ReDim lvlArr(1 To n)
    cnt = 0
    r = FIRST_ROW - 1
    firstZ = 0
    gDirectFirst = FIRST_ROW

    ' X1000（本工事費）の位置を探す。これより前はGコードの単価表なので内訳には入れない
    firstX = 0
    For i = 1 To n
        rec = recOld(i)
        If CStr(rec(R_LEVEL)) = "費目" Then firstX = i: Exit For
    Next i

    If firstX = 0 Then
        firstX = 1
        warnText = warnText & IIf(warnText = "", "", vbCrLf) & _
                   "★ X1000（本工事費）の行が見つかりませんでした。" & vbCrLf & _
                   "　 単価表部を分けられないので、全行を内訳として扱っています。"
    Else
        skippedG = firstX - 1
    End If

    '--- 直接工事費部（X1000 〜 最初のZコードの直前）---
    For i = firstX To n
        rec = recOld(i)
        lvl = CStr(rec(R_LEVEL))

        If lvl = "Z" Then
            ' 諸経費部の始まり。YZコード・Gコードでは切らない
            firstZ = i
            Exit For
        End If

        r = r + 1
        lvlNum = LevelNum(lvl)
        WriteRow ws, r, rec, newTanka(i), lvlNum, lvl
        cnt = cnt + 1
        rowArr(cnt) = r
        lvlArr(cnt) = lvlNum
    Next i

    WriteSumFormulas ws, rowArr, lvlArr, cnt
    gDirectLast = r

    '--- 共通仮設費 積上げ部（Zコード配下）---
    gZFirst = 0: gZLast = 0
    If firstZ > 0 Then
        gZFirst = r + 1
        ReDim zRowArr(1 To n)
        ReDim zLvlArr(1 To n)
        zCnt = 0
        curItem = "": curCode = "": curStart = 0

        For i = firstZ To n
            rec = recOld(i)
            lvl = CStr(rec(R_LEVEL))

            r = r + 1
            lvlNum = LevelNum(lvl)
            WriteRow ws, r, rec, newTanka(i), lvlNum, lvl

            If lvl = "Z" Then
                If curItem <> "" Then gZList.Add Array(curItem, curCode, curStart, r - 1)
                curItem = CStr(rec(R_NAME))
                curCode = CStr(rec(R_CODE))
                curStart = r
            End If

            zCnt = zCnt + 1
            zRowArr(zCnt) = r
            zLvlArr(zCnt) = lvlNum
        Next i

        If curItem <> "" Then gZList.Add Array(curItem, curCode, curStart, r)
        WriteSumFormulas ws, zRowArr, zLvlArr, zCnt
        gZLast = r
    End If

    gDetailLast = r
    FinishDetail ws, r

    '--- 経費計算部 ---
    r = BuildKeihiSection(ws, cfg, r + 2)
    gLastRow = r
    FinishSheet ws, r

    '--- 単価表部（代価表の中身）を辿れるように索引を作る ---
    '    特殊集計区分CSVの取込で、代価表の中にある処分費から内訳の行へ上がるのに使います
    If firstX > 1 Then WriteTankaIndex recOld, firstX - 1

    If skippedG > 0 Then
        warnText = warnText & IIf(warnText = "", "", vbCrLf) & _
                   "X1000より前の " & skippedG & " 行（Gコードの単価表）は内訳に入れていません。" & vbCrLf & _
                   "　（「" & SH_TANKA & "」シートに索引として残しています。特殊集計区分CSVの取込で使います）"
    End If

    ws.Activate
    BuildSlideSheet = warnText
End Function


'--------------------------------------------------------------
' 単価表部の索引
'   代価表の見出し（Gコードで数量が無い行）から次の見出しまでが1つの代価表。
'   その中の各行に「所属代価表」を付けて書き出しておく。
'--------------------------------------------------------------
Private Sub WriteTankaIndex(ByVal recs As Collection, ByVal lastIdx As Long)
    Dim ws As Worksheet
    Dim rec As Variant
    Dim i As Long, r As Long
    Dim owner As String

    Set ws = FreshSheet(SH_TANKA)
    ws.Range("A1").Value = "単価表部（代価表の中身）の索引　" & _
        "※［特殊集計区分CSV取込］で、代価表の中の処分費から内訳の行へ上がるのに使います。消さないでください"
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "コード"
    ws.Range("B2").Value = "所属代価表"
    ws.Range("C2").Value = "名称"
    ws.Range("D2").Value = "単位"
    ws.Range("E2").Value = "数量"
    ws.Range("F2").Value = "規格"
    With ws.Range("A2:F2")
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With

    r = 2
    owner = ""
    For i = 1 To lastIdx
        rec = recs(i)
        If CStr(rec(R_LEVEL)) = "G" And CDbl(rec(R_SURYO)) = 0 And CDbl(rec(R_SURYO2)) = 0 Then
            ' 代価表の見出し（Gコードで数量が入っていない行）
            owner = CStr(rec(R_CODE))
        ElseIf owner <> "" Then
            r = r + 1
            ws.Cells(r, 1).Value = "'" & CStr(rec(R_CODE))
            ws.Cells(r, 2).Value = "'" & owner
            ws.Cells(r, 3).Value = rec(R_NAME)
            ws.Cells(r, 4).Value = rec(R_TANI)
            ws.Cells(r, 5).Value = rec(R_SURYO)
            ws.Cells(r, 6).Value = rec(R_KIKAKU)
        End If
    Next i

    If r > 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(r, 6)).Borders.LineStyle = xlContinuous
    ws.Columns("A:B").ColumnWidth = 13
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 6
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 26
    On Error Resume Next
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub


'--------------------------------------------------------------
' 旧CSVの各行に対応する新単価を求める
'--------------------------------------------------------------
Private Function MatchNewPrices(ByVal recOld As Collection, ByVal recNew As Collection, _
                                ByRef newTanka() As Double) As String
    Dim i As Long, miss As Long
    Dim positional As Boolean
    Dim d As Object
    Dim k As String
    Dim rec As Variant

    positional = (recOld.Count = recNew.Count)
    If positional Then
        For i = 1 To recOld.Count
            If RecKey(recOld(i)) <> RecKey(recNew(i)) Then
                positional = False
                Exit For
            End If
        Next i
    End If

    If positional Then
        For i = 1 To recOld.Count
            rec = recNew(i)
            newTanka(i) = CDbl(rec(R_TANKA))
        Next i
        Exit Function
    End If

    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To recNew.Count
        rec = recNew(i)
        k = RecKey(rec)
        If Not d.Exists(k) Then d.Add k, CDbl(rec(R_TANKA))
    Next i

    For i = 1 To recOld.Count
        rec = recOld(i)
        k = RecKey(rec)
        If d.Exists(k) Then
            newTanka(i) = d(k)
        Else
            newTanka(i) = CDbl(rec(R_TANKA))
            If CStr(rec(R_LEVEL)) = "" Then miss = miss + 1
        End If
    Next i

    MatchNewPrices = "旧CSVと新CSVで行がずれていたため、名称・規格・単位で突合しました。" & vbCrLf & _
                     "新単価が見つからず旧単価を据え置いた明細：" & miss & " 行"
End Function


' 名称が「スクラップ」を含むか（全角／半角カタカナの差は NormText が吸収する）
Public Function IsScrapName(ByVal s As String) As Boolean
    If s = "" Then Exit Function
    IsScrapName = (InStr(1, NormText(s), NormText("スクラップ")) > 0)
End Function


' 新工種を考慮するかの既定値
'   ※「新工種なし版」ではここが "しない" になっています。
'     設定シートに項目が無いときもこの値が使われます。
Public Function DefaultUseShin() As String
    DefaultUseShin = "する"
End Function


Public Function LevelNum(ByVal lvl As String) As Long
    Select Case lvl
        Case "費目": LevelNum = 0
        Case "L1":   LevelNum = 1
        Case "L2":   LevelNum = 2
        Case "L3":   LevelNum = 3
        Case "L4":   LevelNum = 4
        Case "Z":    LevelNum = 1
        Case "YZ":   LevelNum = 2
        Case Else:   LevelNum = 9      ' 明細（Gコードの代価行もここ）
    End Select
End Function


'==============================================================
' 見出し
'==============================================================
Private Sub WriteHeader(ByVal ws As Worksheet, ByVal cfg As Object)
    ws.Range("B1").Value = CfgVal(cfg, "工事名", "")
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("G1").Value = "基準日：" & CStr(CfgVal(cfg, "基準日", ""))
    ws.Range("K1").Value = "新工種：" & IIf(gUseShin, "考慮する", "考慮しない")
    ws.Range("K1").Font.Bold = True
    ws.Range("K1").Font.Color = IIf(gUseShin, RGB(0, 0, 0), RGB(192, 0, 0))

    ws.Cells(3, SC_CODE).Value = "コード"
    ws.Cells(3, SC_HIMOKU).Value = "費目"
    ws.Cells(3, SC_KUBUN).Value = "工事区分"
    ws.Cells(4, SC_KOSHU).Value = "工種"
    ws.Cells(5, SC_SHUBETSU).Value = "種別"
    ws.Cells(6, SC_SAIBETSU).Value = "細別"
    ws.Cells(3, SC_KIKAKU).Value = "規格"
    ws.Cells(3, SC_TANKA_O).Value = "単価"
    ws.Cells(5, SC_TANKA_O).Value = "スライド前"
    ws.Cells(5, SC_TANKA_N).Value = "スライド後"
    ws.Cells(3, SC_Q_ALL).Value = "数量"
    ws.Cells(5, SC_Q_ALL).Value = "全体"
    ws.Cells(5, SC_Q_DONE).Value = "出来形"
    ws.Cells(5, SC_Q_REST).Value = "残工事"
    ws.Cells(3, SC_TANI).Value = "単位"
    ws.Cells(3, SC_MK_SHOBU).Value = "区分（〇を入力）"
    ws.Cells(5, SC_MK_SHOBU).Value = "処分費"
    ws.Cells(5, SC_MK_KANZA).Value = "管材費"
    If gUseShin Then ws.Cells(5, SC_MK_SHIN).Value = "新工種"
    ws.Cells(3, SC_ZKUBUN).Value = "区分"
    ws.Cells(5, SC_ZKUBUN).Value = "経費計算部用"
    ws.Cells(3, SC_A_ALL).Value = "スライド前（旧単価）"
    ws.Cells(5, SC_A_ALL).Value = "全体"
    ws.Cells(5, SC_A_DONE).Value = "出来形"
    ws.Cells(5, SC_A_REST).Value = "残工事"
    ws.Cells(3, SC_B_ALL).Value = "スライド後（新単価）"
    ws.Cells(5, SC_B_ALL).Value = "全体"
    ws.Cells(5, SC_B_REST).Value = "残工事"
    If gUseShin Then
        ws.Cells(3, SC_N_ALL).Value = "新工種抜き"
        ws.Cells(5, SC_N_ALL).Value = "全体"
        ws.Cells(5, SC_N_DONE).Value = "出来形"
    End If
    ws.Cells(3, SC_TEKIYO).Value = "摘要"
    ws.Cells(3, SC_SHA_O).Value = "処分費額（手入力・全数量分）"
    ws.Cells(5, SC_SHA_O).Value = "スライド前"
    ws.Cells(5, SC_SHA_N).Value = "スライド後"
    ws.Cells(3, SC_SA_ALL).Value = "処分費額"
    ws.Cells(5, SC_SA_ALL).Value = "前・全体"
    ws.Cells(5, SC_SA_DONE).Value = "前・出来形"
    ws.Cells(5, SC_SA_REST).Value = "前・残工事"
    ws.Cells(5, SC_SB_ALL).Value = "後・全体"
    ws.Cells(5, SC_SB_REST).Value = "後・残工事"
    If gUseShin Then
        ws.Cells(5, SC_SN_ALL).Value = "新抜き・全体"
        ws.Cells(5, SC_SN_DONE).Value = "新抜き・出来形"
    End If

    With ws.Range(ws.Cells(3, 1), ws.Cells(6, SC_LAST))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(3, SC_MK_SHOBU), ws.Cells(6, IIf(gUseShin, SC_MK_SHIN, SC_MK_KANZA))).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(3, SC_ZKUBUN), ws.Cells(6, SC_ZKUBUN)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(3, SC_SHA_O), ws.Cells(6, SC_SHA_N)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(6, 1), ws.Cells(6, SC_LAST)).Borders(xlEdgeBottom).LineStyle = xlDouble
End Sub


'==============================================================
' 明細1行
'==============================================================
Private Sub WriteRow(ByVal ws As Worksheet, ByVal r As Long, ByVal rec As Variant, _
                     ByVal tankaNew As Double, ByVal lvlNum As Long, ByVal lvl As String)
    ws.Cells(r, SC_CODE).Value = rec(R_CODE)
    ws.Cells(r, SC_TEKIYO).Value = rec(R_TEKIYO)

    If lvl = "Z" Then
        ws.Cells(r, SC_KUBUN).Value = rec(R_NAME)
    ElseIf lvl = "YZ" Then
        ws.Cells(r, SC_KOSHU).Value = "②"
        ws.Cells(r, SC_SHUBETSU).Value = rec(R_NAME)
    Else
        Select Case lvlNum
            Case 0
                ws.Cells(r, SC_HIMOKU).Value = rec(R_NAME)
            Case 1
                ws.Cells(r, SC_KUBUN).Value = "①"
                ws.Cells(r, SC_KOSHU).Value = rec(R_NAME)
            Case 2
                ws.Cells(r, SC_KOSHU).Value = "②"
                ws.Cells(r, SC_SHUBETSU).Value = rec(R_NAME)
            Case 3
                ws.Cells(r, SC_SHUBETSU).Value = "③"
                ws.Cells(r, SC_SAIBETSU).Value = rec(R_NAME)
            Case 4
                ws.Cells(r, SC_SAIBETSU).Value = "④"
                ws.Cells(r, SC_NAME).Value = rec(R_NAME)
            Case Else
                ws.Cells(r, SC_NAME).Value = rec(R_NAME)
                ws.Cells(r, SC_KIKAKU).Value = rec(R_KIKAKU)
        End Select
    End If

    If lvlNum = 9 Then
        ws.Cells(r, SC_TANKA_O).Value = rec(R_TANKA)
        ws.Cells(r, SC_TANKA_N).Value = tankaNew
        ws.Cells(r, SC_Q_ALL).Value = rec(R_SURYO)
        ws.Cells(r, SC_TANI).Value = rec(R_TANI)
        WriteDetailFormulas ws, r
        ' 採用しなかった側の数量が0なら、この設計で追加された工種とみなす
        If mAutoShin Then
            If CDbl(rec(R_SURYO2)) = 0 And CDbl(rec(R_SURYO)) <> 0 Then
                ws.Cells(r, SC_MK_SHIN).Value = "〇"
            End If
        End If
    Else
        ws.Cells(r, SC_TANI).Value = "式"
    End If
End Sub


Public Sub WriteDetailFormulas(ByVal ws As Worksheet, ByVal r As Long)
    Dim aI As String, aJ As String, aK As String, aL As String, aM As String
    Dim aO As String, aQ As String, aSO As String, aSN As String
    Dim aAll As String, aBall As String
    Dim shO As String, shN As String

    aI = ws.Cells(r, SC_TANKA_O).Address(False, False)
    aJ = ws.Cells(r, SC_TANKA_N).Address(False, False)
    aK = ws.Cells(r, SC_Q_ALL).Address(False, False)
    aL = ws.Cells(r, SC_Q_DONE).Address(False, False)
    aM = ws.Cells(r, SC_Q_REST).Address(False, False)
    aO = ws.Cells(r, SC_MK_SHOBU).Address(False, False)
    aQ = ws.Cells(r, SC_MK_SHIN).Address(False, False)
    aSO = ws.Cells(r, SC_SHA_O).Address(False, False)
    aSN = ws.Cells(r, SC_SHA_N).Address(False, False)
    aAll = ws.Cells(r, SC_A_ALL).Address(False, False)
    aBall = ws.Cells(r, SC_B_ALL).Address(False, False)

    ' 金額は 単価×数量 を1円未満切り捨て
    ws.Cells(r, SC_Q_REST).Formula = "=" & aK & "-" & aL
    ws.Cells(r, SC_A_ALL).Formula = "=ROUNDDOWN(" & aK & "*" & aI & ",0)"
    ws.Cells(r, SC_A_DONE).Formula = "=ROUNDDOWN(" & aL & "*" & aI & ",0)"
    ws.Cells(r, SC_A_REST).Formula = "=ROUNDDOWN(" & aM & "*" & aI & ",0)"
    ws.Cells(r, SC_B_ALL).Formula = "=ROUNDDOWN(" & aK & "*" & aJ & ",0)"
    ws.Cells(r, SC_B_REST).Formula = "=ROUNDDOWN(" & aM & "*" & aJ & ",0)"
    ' 新工種抜き：新工種〇の行を0にする
    If gUseShin Then
        ws.Cells(r, SC_N_ALL).Formula = "=IF(" & aQ & "=""〇"",0," & aAll & ")"
        ws.Cells(r, SC_N_DONE).Formula = "=IF(" & aQ & "=""〇"",0," & _
            ws.Cells(r, SC_A_DONE).Address(False, False) & ")"
    End If

    ' 処分費額：全数量に対する金額を手入力し、出来形・残工事へは数量比で割り振る
    ' 未入力ならその行の金額を全額 処分費とみなす
    shO = "IF(" & aSO & "=""""," & aAll & "," & aSO & ")"
    shN = "IF(" & aSN & "="""",IF(" & aSO & "=""""," & aBall & "," & aSO & ")," & aSN & ")"

    ' 処分費額も数量比で割り振ったあと1円未満切り捨て
    ws.Cells(r, SC_SA_ALL).Formula = "=IF(" & aO & "<>""〇"",0,ROUNDDOWN(" & shO & ",0))"
    ws.Cells(r, SC_SA_DONE).Formula = "=IF(" & aO & "<>""〇"",0,IF(" & aK & "=0,0,ROUNDDOWN(" & _
        shO & "*" & aL & "/" & aK & ",0)))"
    ws.Cells(r, SC_SA_REST).Formula = "=IF(" & aO & "<>""〇"",0,IF(" & aK & "=0,0,ROUNDDOWN(" & _
        shO & "*" & aM & "/" & aK & ",0)))"
    ws.Cells(r, SC_SB_ALL).Formula = "=IF(" & aO & "<>""〇"",0,ROUNDDOWN(" & shN & ",0))"
    ws.Cells(r, SC_SB_REST).Formula = "=IF(" & aO & "<>""〇"",0,IF(" & aK & "=0,0,ROUNDDOWN(" & _
        shN & "*" & aM & "/" & aK & ",0)))"
    If gUseShin Then
        ws.Cells(r, SC_SN_ALL).Formula = "=IF(" & aQ & "=""〇"",0," & _
            ws.Cells(r, SC_SA_ALL).Address(False, False) & ")"
        ws.Cells(r, SC_SN_DONE).Formula = "=IF(" & aQ & "=""〇"",0," & _
            ws.Cells(r, SC_SA_DONE).Address(False, False) & ")"
    End If
End Sub


'==============================================================
' 階層行の集計式（金額7列。処分費額は明細行だけが持つ）
'==============================================================
Public Sub WriteSumFormulas(ByVal ws As Worksheet, ByRef rowArr() As Long, _
                            ByRef lvlArr() As Long, ByVal cnt As Long)
    Dim i As Long, j As Long, k As Long
    Dim parentIdx() As Long
    Dim stackLvl() As Long, stackIdx() As Long, sp As Long
    Dim cols As Variant, c As Variant
    Dim addrs As String
    Dim childRows() As Long, nc As Long
    Dim contiguous As Boolean

    If cnt = 0 Then Exit Sub

    ReDim parentIdx(1 To cnt)
    ReDim stackLvl(0 To cnt)
    ReDim stackIdx(0 To cnt)
    sp = 0

    For i = 1 To cnt
        Do While sp > 0
            If stackLvl(sp) >= lvlArr(i) Then sp = sp - 1 Else Exit Do
        Loop
        If sp > 0 Then parentIdx(i) = stackIdx(sp) Else parentIdx(i) = 0
        If lvlArr(i) < 9 Then
            sp = sp + 1
            stackLvl(sp) = lvlArr(i)
            stackIdx(sp) = i
        End If
    Next i

    If gUseShin Then
        cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST, SC_N_ALL, SC_N_DONE)
    Else
        cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST)
    End If

    For i = 1 To cnt
        If lvlArr(i) < 9 Then
            ReDim childRows(1 To cnt)
            nc = 0
            For j = i + 1 To cnt
                If parentIdx(j) = i Then
                    nc = nc + 1
                    childRows(nc) = rowArr(j)
                ElseIf lvlArr(j) <= lvlArr(i) Then
                    Exit For
                End If
            Next j

            If nc > 0 Then
                contiguous = True
                For k = 2 To nc
                    If childRows(k) <> childRows(k - 1) + 1 Then contiguous = False: Exit For
                Next k

                For Each c In cols
                    If contiguous Then
                        addrs = ws.Cells(childRows(1), c).Address(False, False) & ":" & _
                                ws.Cells(childRows(nc), c).Address(False, False)
                    Else
                        addrs = ""
                        For k = 1 To nc
                            If addrs <> "" Then addrs = addrs & ","
                            addrs = addrs & ws.Cells(childRows(k), c).Address(False, False)
                        Next k
                    End If
                    ws.Cells(rowArr(i), c).Formula = "=SUM(" & addrs & ")"
                Next c
            End If
        End If
    Next i
End Sub


'==============================================================
' 経費計算部（同じシートの明細の下に作る）
'   戻り値：最終行
'==============================================================
Private Function BuildKeihiSection(ByVal ws As Worksheet, ByVal cfg As Object, _
                                   ByVal startRow As Long) As Long
    Dim cols As Variant, shoCols As Variant, rateSrc As Variant
    Dim i As Long, c As Long, r As Long, z As Long
    Dim kojo As String, zei As String
    Dim zi As Variant
    Dim zRow1 As Long, zRow2 As Long
    Dim rDCHOKU As Long, rKANZAI As Long, rSHOBUN As Long, rTAIGAI As Long
    Dim rTSUMI As Long
    Dim rKTAISHO As Long, rKRITSU As Long, rKBUN As Long, rKKEI As Long, rJUN As Long
    Dim rGTAISHO As Long, rGRITSU As Long, rGKEI As Long, rGENKA As Long
    Dim rITAISHO As Long, rIRITSU As Long, rIBUN As Long, rHOSHO As Long
    Dim rIKEI0 As Long, rKAKAKU0 As Long, rHASU As Long, rIKEI As Long, rKAKAKU As Long
    Dim rSCRAP As Long, rZGEN As Long, rZKAK As Long

    Set mKWs = ws
    Set mKCfg = cfg

    If gUseShin Then
        cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST, SC_N_ALL, SC_N_DONE)
        shoCols = Array(SC_SA_ALL, SC_SA_DONE, SC_SA_REST, SC_SB_ALL, SC_SB_REST, SC_SN_ALL, SC_SN_DONE)
        ' 0＝自前の率（手入力可）／-1＝率なし（素材のみ）／1以上＝その系列番号の率を参照
        rateSrc = Array(0, 1, -1, 0, 4, 0, 6)
    Else
        cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST)
        shoCols = Array(SC_SA_ALL, SC_SA_DONE, SC_SA_REST, SC_SB_ALL, SC_SB_REST)
        rateSrc = Array(0, 1, -1, 0, 4)
    End If

    kojo = NumStr(CfgVal(cfg, "処分費控除率", 0.03))
    zei = NumStr(CfgVal(cfg, "消費税率", 0.1))

    '--- 行番号を先に確定させる ---
    r = startRow
    ws.Cells(r, SC_HIMOKU).Value = "【経費計算】　※黄色セルは手入力。経費率は積算システムの値があれば上書きしてください"
    With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_ZKUBUN))
        .Merge
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
    End With
    r = r + 1

    rDCHOKU = r:  KLabel ws, r, "①直接工事費":                                          r = r + 1
    rKANZAI = r:  KLabel ws, r, "①'水道工事における管材費":                             r = r + 1
    rSHOBUN = r:  KLabel ws, r, "②直接工事費内の処分費等　※スクラップは含みません（㉒へ）": r = r + 1
    rTAIGAI = r:  KLabel ws, r, "③②のうち率計算の対象外費　※②－(①－①'/2)×3%":       r = r + 1

    '--- Zコード項目を1行ずつ出す（区分は自動判定。R列で直せます）---
    ws.Cells(r, SC_HIMOKU).Value = "─ 諸経費部のZコード項目　※R列の区分は手直しできます　" & _
        "積＝共通仮設費の積上分(④)／原＝工事原価に加算(⑫')／価＝工事価格に加算(⑳')／ス＝スクラップ(㉒・売却益なのでマイナス)／空欄＝使わない ─"
    With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_ZKUBUN))
        .Merge
        .Font.Color = RGB(120, 120, 120)
    End With
    r = r + 1

    zRow1 = r
    If Not gZList Is Nothing Then
        For z = 1 To gZList.Count
            zi = gZList(z)
            KLabelZ ws, r, "　" & CStr(zi(1)) & "　" & CStr(zi(0))
            ws.Cells(r, SC_ZKUBUN).Value = ZKubun(CStr(zi(1)), CStr(zi(0)))
            ws.Cells(r, SC_ZKUBUN).Interior.Color = CLR_INPUT
            ws.Cells(r, SC_ZKUBUN).HorizontalAlignment = xlCenter
            For i = 0 To UBound(cols)
                ws.Cells(r, CLng(cols(i))).Formula = ZRangeSum(CLng(zi(2)), CLng(zi(3)), CLng(cols(i)))
            Next i
            r = r + 1
        Next z
    End If
    zRow2 = r - 1
    If zRow2 < zRow1 Then zRow2 = zRow1

    rTSUMI = r:   KLabel ws, r, "④共通仮設費 積上分 計　※上の区分が「積」の合計":       r = r + 1
    rKTAISHO = r: KLabel ws, r, "⑥共通仮設費対象額　※①－③－①'/2":                    r = r + 1
    rKRITSU = r:  KLabel ws, r, "⑦共通仮設費率":                                        r = r + 1
    rKBUN = r:    KLabel ws, r, "⑦'共通仮設費率分　※⑥×⑦":                            r = r + 1
    rKKEI = r:    KLabel ws, r, "⑧共通仮設費【合計】　※④＋⑦'":                        r = r + 1
    rJUN = r:     KLabel ws, r, "⑨純工事費　※①＋⑧":                                   r = r + 1
    rGTAISHO = r: KLabel ws, r, "⑩現場管理費対象額　※⑨－③－①'/2":                    r = r + 1
    rGRITSU = r:  KLabel ws, r, "⑪現場管理費率":                                        r = r + 1
    rGKEI = r:    KLabel ws, r, "⑫現場管理費【合計】　※⑩×⑪":                          r = r + 1
    rZGEN = r:    KLabel ws, r, "⑫'工事原価に加算するZ項目　※区分「原」の合計":          r = r + 1
    rGENKA = r:   KLabel ws, r, "⑬工事原価　※⑨＋⑫＋⑫'":                               r = r + 1
    rITAISHO = r: KLabel ws, r, "⑭一般管理費対象額　※⑬－③":                           r = r + 1
    rIRITSU = r:  KLabel ws, r, "⑮一般管理費率":                                        r = r + 1
    rIBUN = r:    KLabel ws, r, "⑮'一般管理費率分　※⑭×⑮":                            r = r + 1
    rHOSHO = r:   KLabel ws, r, "⑯契約保証費　※当初設計時額で固定":                     r = r + 1
    rIKEI0 = r:   KLabel ws, r, "⑰一般管理費【合計】（端数整理前）　※⑮'＋⑯":          r = r + 1
    rKAKAKU0 = r: KLabel ws, r, "⑱工事価格（端数整理前）　※⑬＋⑰":                      r = r + 1
    rHASU = r:    KLabel ws, r, "⑲端数整理":                                            r = r + 1
    rIKEI = r:    KLabel ws, r, "⑳一般管理費【合計】（端数整理後）　※⑰－⑲":           r = r + 1
    rZKAK = r:    KLabel ws, r, "⑳'工事価格に加算するZ項目　※区分「価」の合計":          r = r + 1
    rKAKAKU = r:  KLabel ws, r, "㉑工事価格　※⑬＋⑳＋⑳'":                               r = r + 1
    rSCRAP = r:   KLabel ws, r, "㉒スクラップ　※区分「ス」のZコード項目の合計（1000円単位に切捨）。売却益なのでマイナス": r = r + 1
    gRowKakaku2 = r: KLabel ws, r, "㉑'工事価格（スクラップ込み）　※㉑＋㉒":             r = r + 1
    gRowZei = r:     KLabel ws, r, "㉓消費税相当額　※(㉑＋㉒)×消費税率":                r = r + 1
    gRowKoujihi = r: KLabel ws, r, "㉔工事費　※㉑＋㉒＋㉓":                              r = r + 1

    '--- 系列ごとに数式を入れる ---
    For i = 0 To UBound(cols)
        c = CLng(cols(i))

        ' ①直接工事費（明細行のみ。スクラップ〇は㉒へまわす）
        ws.Cells(rDCHOKU, c).Formula = "=SUMPRODUCT((" & KR(SC_TANKA_O, gDirectFirst, gDirectLast) & _
            "<>"""")*(" & KR(c, gDirectFirst, gDirectLast) & "))"

        ws.Cells(rKANZAI, c).Formula = "=SUMPRODUCT((" & KR(SC_MK_KANZA, gDirectFirst, gDirectLast) & _
            "=""〇"")*(" & KR(c, gDirectFirst, gDirectLast) & "))"

        ws.Cells(rSHOBUN, c).Formula = "=SUMPRODUCT((" & KR(SC_TANKA_O, gDirectFirst, gDirectLast) & _
            "<>"""")*(" & KR(CLng(shoCols(i)), gDirectFirst, gDirectLast) & "))"

        ws.Cells(rTAIGAI, c).Formula = "=MAX(0," & KC(rSHOBUN, c) & "-ROUNDDOWN((" & _
            KC(rDCHOKU, c) & "-" & KC(rKANZAI, c) & "/2)*" & kojo & ",0))"

        ' ④積上分 計 ＝ 区分「積」のZ項目の合計
        ws.Cells(rTSUMI, c).Formula = "=SUMPRODUCT((" & KR(SC_ZKUBUN, zRow1, zRow2) & _
            "=""積"")*(" & KR(c, zRow1, zRow2) & "))"

        If CLng(rateSrc(i)) < 0 Then
            ws.Cells(rKTAISHO, c).Value = "素材のみ"
            ws.Cells(rKTAISHO, c).Font.Color = RGB(120, 120, 120)
            ws.Cells(rKTAISHO, c).HorizontalAlignment = xlCenter
            GoTo NextSeries
        End If

        ws.Cells(rKTAISHO, c).Formula = "=" & KC(rDCHOKU, c) & "-" & KC(rTAIGAI, c) & _
                                        "-" & KC(rKANZAI, c) & "/2"
        KRate ws, rKRITSU, c, CLng(rateSrc(i)), cols, RateKasetsu(KC(rKTAISHO, c))
        ws.Cells(rKBUN, c).Formula = "=ROUNDDOWN(" & KC(rKTAISHO, c) & "*" & KC(rKRITSU, c) & ",-3)"
        ws.Cells(rKKEI, c).Formula = "=" & KC(rTSUMI, c) & "+" & KC(rKBUN, c)
        ws.Cells(rJUN, c).Formula = "=" & KC(rDCHOKU, c) & "+" & KC(rKKEI, c)

        ws.Cells(rGTAISHO, c).Formula = "=" & KC(rJUN, c) & "-" & KC(rTAIGAI, c) & _
                                        "-" & KC(rKANZAI, c) & "/2"
        KRate ws, rGRITSU, c, CLng(rateSrc(i)), cols, RateGenba(KC(rGTAISHO, c))
        ws.Cells(rGKEI, c).Formula = "=ROUNDDOWN(" & KC(rGTAISHO, c) & "*" & KC(rGRITSU, c) & ",-3)"
        ws.Cells(rZGEN, c).Formula = "=SUMPRODUCT((" & KR(SC_ZKUBUN, zRow1, zRow2) & _
            "=""原"")*(" & KR(c, zRow1, zRow2) & "))"
        ws.Cells(rGENKA, c).Formula = "=" & KC(rJUN, c) & "+" & KC(rGKEI, c) & "+" & KC(rZGEN, c)

        ws.Cells(rITAISHO, c).Formula = "=" & KC(rGENKA, c) & "-" & KC(rTAIGAI, c)
        KRate ws, rIRITSU, c, CLng(rateSrc(i)), cols, RateIppan(KC(rITAISHO, c))
        ws.Cells(rIBUN, c).Formula = "=" & KC(rITAISHO, c) & "*" & KC(rIRITSU, c)
        ws.Cells(rHOSHO, c).Value = CDbl(CfgVal(cfg, "契約保証費", 0))
        ws.Cells(rHOSHO, c).Interior.Color = CLR_INPUT
        ws.Cells(rIKEI0, c).Formula = "=" & KC(rIBUN, c) & "+" & KC(rHOSHO, c)
        ws.Cells(rZKAK, c).Formula = "=SUMPRODUCT((" & KR(SC_ZKUBUN, zRow1, zRow2) & _
            "=""価"")*(" & KR(c, zRow1, zRow2) & "))"
        ws.Cells(rKAKAKU0, c).Formula = "=" & KC(rGENKA, c) & "+" & KC(rIKEI0, c) & "+" & KC(rZKAK, c)
        ws.Cells(rHASU, c).Formula = "=" & KC(rKAKAKU0, c) & "-ROUNDDOWN(" & KC(rKAKAKU0, c) & ",-3)"
        ws.Cells(rIKEI, c).Formula = "=" & KC(rIKEI0, c) & "-" & KC(rHASU, c)
        ws.Cells(rKAKAKU, c).Formula = "=" & KC(rGENKA, c) & "+" & KC(rIKEI, c) & "+" & KC(rZKAK, c)

        ' ㉒スクラップ ＝ 区分「ス」のZ項目。売却益なのでマイナスで入り、1000円単位に切り捨てる
        ws.Cells(rSCRAP, c).Formula = "=ROUNDDOWN(SUMPRODUCT((" & KR(SC_ZKUBUN, zRow1, zRow2) & _
            "=""ス"")*(" & KR(c, zRow1, zRow2) & ")),-3)"

        ws.Cells(gRowKakaku2, c).Formula = "=" & KC(rKAKAKU, c) & "+" & KC(rSCRAP, c)
        ws.Cells(gRowZei, c).Formula = "=(" & KC(rKAKAKU, c) & "+" & KC(rSCRAP, c) & ")*" & zei
        ws.Cells(gRowKoujihi, c).Formula = "=" & KC(rKAKAKU, c) & "+" & KC(rSCRAP, c) & _
                                           "+" & KC(gRowZei, c)
NextSeries:
    Next i

    '--- 体裁 ---
    With ws.Range(ws.Cells(startRow + 1, SC_A_ALL), ws.Cells(gRowKoujihi, CLng(cols(UBound(cols)))))
        .NumberFormatLocal = "#,##0"
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(rKRITSU, SC_A_ALL), ws.Cells(rKRITSU, CLng(cols(UBound(cols))))).NumberFormatLocal = "0.0000"
    ws.Range(ws.Cells(rGRITSU, SC_A_ALL), ws.Cells(rGRITSU, CLng(cols(UBound(cols))))).NumberFormatLocal = "0.0000"
    ws.Range(ws.Cells(rIRITSU, SC_A_ALL), ws.Cells(rIRITSU, CLng(cols(UBound(cols))))).NumberFormatLocal = "0.0000"
    KHiLite ws, rTSUMI
    KHiLite ws, rJUN
    KHiLite ws, rGENKA
    KHiLite ws, rKAKAKU
    KHiLite ws, gRowKakaku2
    KHiLite ws, gRowKoujihi

    If gZList Is Nothing Then
        KWarn ws, rTSUMI, "Zコード項目が1つもありません。共通仮設費の積上分は0です"
    ElseIf gZList.Count = 0 Then
        KWarn ws, rTSUMI, "Zコード項目が1つもありません。共通仮設費の積上分は0です"
    End If

    BuildKeihiSection = gRowKoujihi
End Function


'--------------------------------------------------------------
' 経費計算部の補助
'--------------------------------------------------------------
Private Sub KLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, SC_HIMOKU).Value = s
    With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_ZKUBUN))
        .Merge
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
    End With
End Sub


Private Sub KLabelZ(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, SC_HIMOKU).Value = s
    With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_MK_SHIN))
        .Merge
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
    End With
    ws.Cells(r, SC_ZKUBUN).Borders.LineStyle = xlContinuous
End Sub


Private Sub KRate(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, _
                  ByVal src As Long, ByVal cols As Variant, ByVal autoFormula As String)
    If src = 0 Then
        ws.Cells(r, c).Formula = autoFormula
        ws.Cells(r, c).Interior.Color = CLR_INPUT
    Else
        ws.Cells(r, c).Formula = "=" & KC(r, CLng(cols(src - 1)))
    End If
End Sub


Private Sub KHiLite(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Range(ws.Cells(r, SC_A_ALL), ws.Cells(r, IIf(gUseShin, SC_N_DONE, SC_B_REST)))
        .Font.Bold = True
        .Interior.Color = CLR_TOTAL
    End With
End Sub


Private Sub KWarn(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, SC_TEKIYO).Value = "★ " & s
    ws.Cells(r, SC_TEKIYO).Font.Color = RGB(192, 0, 0)
End Sub


Private Function KC(ByVal r As Long, ByVal c As Long) As String
    KC = mKWs.Cells(r, c).Address(False, False)
End Function


Private Function KR(ByVal c As Long, ByVal r1 As Long, ByVal r2 As Long) As String
    If r2 < r1 Then r2 = r1
    KR = mKWs.Cells(r1, c).Address(True, True) & ":" & mKWs.Cells(r2, c).Address(True, True)
End Function


' Zコード項目の区分を自動判定する
'   積 ＝ 共通仮設費の積上分（④へ）　　… Z0001〜Z0039
'   原 ＝ 工事原価に加算（⑫'へ）　　　 … Z0040〜Z0044（工期延長等に伴う増加費用など）
'   価 ＝ 工事価格に加算（⑳'へ）　　　 … Z0045以降（設計委託費など）
'   ス ＝ スクラップ（㉒へ）　　　　　　… 名称に「スクラップ」
'   空 ＝ 使わない　　　　　　　　　　　… 支給品費（共通仮設費計から除く運用のため）
' R列で手直しできます。
Private Function ZKubun(ByVal code As String, ByVal nm As String) As String
    Dim n As Long

    If IsScrapName(nm) Then ZKubun = "ス": Exit Function
    If InStr(1, NormText(nm), NormText("支給品費")) > 0 Then Exit Function

    n = ZCodeNum(code)
    If n = 0 Then ZKubun = "積": Exit Function          ' コードが読めなければ積上分とみなす
    If n < 40 Then
        ZKubun = "積"
    ElseIf n < 45 Then
        ZKubun = "原"
    Else
        ZKubun = "価"
    End If
End Function


' "Z0001" → 1 ／ "Z0040" → 40 ／ 数字が取れなければ 0
Private Function ZCodeNum(ByVal code As String) As Long
    Dim t As String, i As Long, ch As String, digits As String

    t = NormText(code)
    If Len(t) = 0 Then Exit Function
    If UCase$(Left$(t, 1)) <> "Z" Then Exit Function

    For i = 2 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            Exit For
        End If
    Next i

    If digits <> "" Then ZCodeNum = CLng(digits)
End Function


' Zコード項目の配下の明細行を合計する
Private Function ZRangeSum(ByVal r1 As Long, ByVal r2 As Long, ByVal c As Long) As String
    ZRangeSum = "=SUMPRODUCT((" & KR(SC_TANKA_O, r1, r2) & "<>"""")*(" & KR(c, r1, r2) & "))"
End Function


Private Function RateKasetsu(ByVal taisho As String) As String
    Dim a As String, b As String, ch As String, sh As String
    a = NumStr(CfgVal(mKCfg, "共通仮設費率A", 485.4))
    b = NumStr(CfgVal(mKCfg, "共通仮設費率B", -0.2231))
    ch = NumStr(CfgVal(mKCfg, "共通仮設費 地域補正", 1.5))
    sh = NumStr(CfgVal(mKCfg, "共通仮設費 週休補正", 1.01))
    RateKasetsu = "=IF(" & taisho & "<=0,0,ROUND(ROUND(ROUND(" & a & "*" & taisho & _
                  "^(" & b & "),2)*" & ch & ",2)*" & sh & ",2)/100)"
End Function


Private Function RateGenba(ByVal taisho As String) As String
    Dim a As String, b As String, ch As String, sh As String
    a = NumStr(CfgVal(mKCfg, "現場管理費率A", 202.3))
    b = NumStr(CfgVal(mKCfg, "現場管理費率B", -0.1034))
    ch = NumStr(CfgVal(mKCfg, "現場管理費 地域補正", 1.2))
    sh = NumStr(CfgVal(mKCfg, "現場管理費 週休補正", 1.02))
    RateGenba = "=IF(" & taisho & "<=0,0,ROUND(ROUND(ROUND(" & a & "*" & taisho & _
                "^(" & b & "),2)*" & ch & ",2)*" & sh & ",2)/100)"
End Function


Private Function RateIppan(ByVal taisho As String) As String
    Dim k As String, t As String
    k = NumStr(CfgVal(mKCfg, "一般管理費率係数", -4.97802))
    t = NumStr(CfgVal(mKCfg, "一般管理費率定数", 56.92101))
    RateIppan = "=IF(" & taisho & "<=0,0,ROUND(" & k & "*LOG10(" & taisho & ")+" & t & ",2)/100)"
End Function


Private Function NumStr(ByVal v As Variant) As String
    NumStr = Format$(CDbl(v), "0.##########")
End Function


'==============================================================
' 明細部の体裁
'==============================================================
Private Sub FinishDetail(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long

    If lastRow < FIRST_ROW Then Exit Sub

    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, SC_LAST)).Borders.LineStyle = xlContinuous

    ws.Range(ws.Cells(FIRST_ROW, SC_TANKA_O), ws.Cells(lastRow, SC_TANKA_N)).NumberFormatLocal = "#,##0"
    ws.Range(ws.Cells(FIRST_ROW, SC_Q_ALL), ws.Cells(lastRow, SC_Q_REST)).NumberFormatLocal = "#,##0.00"
    ws.Range(ws.Cells(FIRST_ROW, SC_A_ALL), ws.Cells(lastRow, SC_N_DONE)).NumberFormatLocal = "#,##0"
    ws.Range(ws.Cells(FIRST_ROW, SC_SHA_O), ws.Cells(lastRow, SC_SN_DONE)).NumberFormatLocal = "#,##0"

    ws.Range(ws.Cells(FIRST_ROW, SC_Q_DONE), ws.Cells(lastRow, SC_Q_DONE)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_MK_SHOBU), ws.Cells(lastRow, SC_MK_SHIN)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_SHA_O), ws.Cells(lastRow, SC_SHA_N)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_MK_SHOBU), ws.Cells(lastRow, SC_ZKUBUN)).HorizontalAlignment = xlCenter

    For r = FIRST_ROW To lastRow
        If ws.Cells(r, SC_HIMOKU).Value <> "" Then
            With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO))
                .Interior.Color = RGB(47, 117, 181)
                .Font.Color = vbWhite
            End With
        ElseIf ws.Cells(r, SC_KUBUN).Value <> "" Then
            ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO)).Interior.Color = RGB(155, 194, 230)
        ElseIf ws.Cells(r, SC_KOSHU).Value <> "" Then
            ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO)).Interior.Color = RGB(189, 215, 238)
        ElseIf ws.Cells(r, SC_SHUBETSU).Value <> "" Then
            ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO)).Interior.Color = RGB(221, 235, 247)
        End If
        ' 階層行には手入力欄を出さない
        If ws.Cells(r, SC_TANKA_O).Value = "" Then
            ws.Range(ws.Cells(r, SC_MK_SHOBU), ws.Cells(r, SC_MK_SHIN)).Interior.ColorIndex = xlColorIndexNone
            ws.Range(ws.Cells(r, SC_Q_DONE), ws.Cells(r, SC_Q_DONE)).Interior.ColorIndex = xlColorIndexNone
            ws.Range(ws.Cells(r, SC_SHA_O), ws.Cells(r, SC_SHA_N)).Interior.ColorIndex = xlColorIndexNone
        End If
    Next r
End Sub


'==============================================================
' 列幅・印刷設定
'==============================================================
Public Sub FinishSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    ws.Columns(SC_CODE).ColumnWidth = 9
    ws.Columns(SC_CODE).Font.Color = RGB(150, 150, 150)
    ws.Columns(SC_HIMOKU).ColumnWidth = 10
    ws.Range(ws.Columns(SC_KUBUN), ws.Columns(SC_SAIBETSU)).ColumnWidth = 4
    ws.Columns(SC_NAME).ColumnWidth = 22
    ws.Columns(SC_KIKAKU).ColumnWidth = 22
    ws.Range(ws.Columns(SC_TANKA_O), ws.Columns(SC_Q_REST)).ColumnWidth = 10
    ws.Columns(SC_TANI).ColumnWidth = 6
    ws.Range(ws.Columns(SC_MK_SHOBU), ws.Columns(SC_ZKUBUN)).ColumnWidth = 6
    ws.Range(ws.Columns(SC_A_ALL), ws.Columns(SC_N_DONE)).ColumnWidth = 13
    ws.Columns(SC_TEKIYO).ColumnWidth = 24
    ws.Range(ws.Columns(SC_SHA_O), ws.Columns(SC_SN_DONE)).ColumnWidth = 11

    ' 新工種を考慮しないときは、使わない列を隠す
    If Not gUseShin Then
        ws.Columns(SC_MK_SHIN).Hidden = True
        ws.Range(ws.Columns(SC_N_ALL), ws.Columns(SC_N_DONE)).Hidden = True
        ws.Range(ws.Columns(SC_SN_ALL), ws.Columns(SC_SN_DONE)).Hidden = True
    End If

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$3:$6"
        .PrintArea = ws.Range(ws.Cells(1, SC_HIMOKU), ws.Cells(lastRow, SC_TEKIYO)).Address
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
    End With

    ws.Activate
    ws.Cells(FIRST_ROW, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub


'==============================================================
' 《M40_Chosho》 スライド調書（様式4-2号）
'==============================================================


Public Sub BuildChosho(ByVal cfg As Object)
    Dim ws As Worksheet, sl As Worksheet
    Dim contract As Double
    Dim fKoujihi As String, fKakaku As String, fDekidaka As String
    Dim fZanNew As String, fZei As String, fShinNuki As String, fShinNukiD As String
    Dim zeiRate As String

    Set sl = ThisWorkbook.Worksheets(SH_SLIDE)
    Set ws = FreshSheet(SH_CHOSHO)
    contract = CDbl(CfgVal(cfg, "請負代金額", 0))
    zeiRate = Format$(CDbl(CfgVal(cfg, "消費税率", 0.1)), "0.##########")

    ' スライド計算表へのリンク（S=前全体 T=前出来形 V=後全体 W=後残工事 X=新抜き全体 Y=新抜き出来形）
    fKoujihi = SlideRef(sl, gRowKoujihi, SC_A_ALL)
    fKakaku = SlideRef(sl, gRowKakaku2, SC_A_ALL)
    fDekidaka = SlideRef(sl, gRowKakaku2, SC_A_DONE)
    fZanNew = SlideRef(sl, gRowKakaku2, SC_B_REST)
    fZei = SlideRef(sl, gRowZei, SC_A_ALL)
    If gUseShin Then
        fShinNuki = SlideRef(sl, gRowKakaku2, SC_N_ALL)
        fShinNukiD = SlideRef(sl, gRowKakaku2, SC_N_DONE)
    End If

    ws.Range("G2").Value = "様式4-2号"
    ws.Range("B3").Value = "スライド調書"
    ws.Range("B3").Font.Size = 14
    ws.Range("B3").Font.Bold = True
    ws.Range("E3").Formula = "=IF(E16>F16,""減額スライド"",IF(E16<F16,""増額スライド"",""""))"
    ws.Range("B4").Value = CfgVal(cfg, "工事名", "")
    ws.Range("G4").Value = "新工種：" & IIf(gUseShin, "考慮する", "考慮しない")
    ws.Range("G4").Font.Bold = True
    ws.Range("G4").Font.Color = IIf(gUseShin, RGB(0, 0, 0), RGB(192, 0, 0))

    ws.Range("C5").Value = "元設計"
    ws.Range("D5").Value = "出来高"
    ws.Range("E5").Value = "残工事"
    ws.Range("G5").Value = "スライド額"
    ws.Range("E6").Value = "変動前"
    ws.Range("F6").Value = "変動後"

    ws.Range("C7").Value = "①"
    ws.Range("B8").Value = "設計額（税込）"
    ws.Range("C8").Formula = "=" & fKoujihi
    ws.Range("H8").Value = "請負率(%)"
    ws.Range("I8").Formula = "=IF(C8=0,0,C14/C8*100)"
    ws.Range("I8").Font.Color = RGB(120, 120, 120)

    ws.Range("C9").Value = "②": ws.Range("D9").Value = "⑤"
    ws.Range("E9").Value = "⑦=②-⑤": ws.Range("F9").Value = "⑧"
    ws.Range("B10").Value = "　工事価格（税抜）"
    ws.Range("C10").Formula = "=" & fKakaku
    ws.Range("D10").Formula = "=" & fDekidaka
    ws.Range("E10").Formula = "=C10-D10"
    ws.Range("F10").Formula = "=" & fZanNew

    ws.Range("B12").Value = "　消費税相当額"
    ws.Range("C12").Formula = "=" & fZei

    ws.Range("C13").Value = "③"
    ws.Range("B14").Value = "請負代金額（税込）"
    ws.Range("C14").Value = contract
    ws.Range("C14").Interior.Color = CLR_INPUT
    ws.Range("H14").Value = "★手入力"
    ws.Range("H14").Font.Color = RGB(192, 0, 0)

    ws.Range("B15").Value = "　請負工事価格"
    ws.Range("C15").Value = "④"
    ws.Range("D15").Value = "⑥=⑤×（③/①）"
    ws.Range("E15").Value = "P1'=④-⑥"
    ws.Range("F15").Value = "P2'=⑧×（③/①）"
    ws.Range("G15").Formula = "=IF(E16>F16,""S'=P2'-P1'+（P1''×1/100）"",""S'=P2'-P1'-（P1''×1/100）"")"

    ws.Range("B16").Value = "　（請負代金額（税抜））"
    ws.Range("C16").Formula = "=ROUNDDOWN(C14/(1+" & zeiRate & "),0)"
    ws.Range("D16").Formula = "=ROUNDDOWN(D10*C14/C8,0)"
    ws.Range("E16").Formula = "=C16-D16"
    ws.Range("F16").Formula = "=ROUNDDOWN(F10*C14/C8,0)"
    ' 増額スライドは1%を控除、減額スライドは1%を戻す
    ws.Range("G16").Formula = "=IF(E16>F16,F16-E16+(E20*1/100),F16-E16-(E20*1/100))"

    ws.Range("B18").Value = "　消費税相当額"
    ws.Range("C18").Formula = "=C14-C16"

    If gUseShin Then
        ws.Range("B19").Value = "　請負工事価格" & vbLf & "（新工種抜き）"
        ws.Range("C19").Value = "⑨" & vbLf & "（請負率考慮）"
        ws.Range("D19").Value = "⑩" & vbLf & "（請負率考慮）"
        ws.Range("E19").Value = "P1''=⑨-⑩"
        ws.Range("B20").Value = "　（請負代金額（税抜））"
        ws.Range("C20").Formula = "=ROUNDDOWN(" & fShinNuki & "*C14/C8,0)"
        ws.Range("D20").Formula = "=ROUNDDOWN(" & fShinNukiD & "*C14/C8,0)"
        ws.Range("E20").Formula = "=C20-D20"
    Else
        ws.Range("B19").Value = "　（新工種を考慮しない設定）"
        ws.Range("E19").Value = "P1''=P1'"
        ws.Range("B20").Value = "　1%の母数"
        ws.Range("E20").Formula = "=E16"
        ws.Range("B19:B20").Font.Color = RGB(120, 120, 120)
    End If

    ws.Range("B22").Value = "※ 黄色いセル（請負代金額（税込））だけ手入力です。" & _
                            "請負率・請負工事価格（税抜）・消費税相当額は計算で出ます"
    If gUseShin Then
        ws.Range("B23").Value = "※ P1''（新工種を抜いた請負ベースの残工事）が受注者負担1%の母数です"
    Else
        ws.Range("B23").Value = "※ 新工種を考慮しない設定です。受注者負担1%の母数は P1'（請負ベースの残工事）です"
    End If
    ws.Range("B24").Value = "※ 金額はすべて「スライド計算表」の経費計算部からリンクしています"
    ws.Range("B22:B24").Font.Color = RGB(120, 120, 120)

    '--- 体裁 ---
    ws.Range("C5:C6").Merge
    ws.Range("D5:D6").Merge
    ws.Range("E5:F5").Merge
    ws.Range("G5:G6").Merge
    With ws.Range("C5:G6")
        .HorizontalAlignment = xlCenter
        .Interior.Color = CLR_HEAD
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    ws.Range("B8:G20").Borders.LineStyle = xlContinuous
    ws.Range("C8:G20").NumberFormatLocal = "#,##0"
    ws.Range("C7:G7").NumberFormatLocal = "@"
    ws.Range("C9:G9").NumberFormatLocal = "@"
    ws.Range("C13:G13").NumberFormatLocal = "@"
    ws.Range("C15:G15").NumberFormatLocal = "@"
    ws.Range("C19:G19").NumberFormatLocal = "@"
    ws.Range("I8").NumberFormatLocal = "0.0000"

    With ws.Range("G16")
        .Font.Bold = True
        .Interior.Color = CLR_TOTAL
    End With
    ws.Range("B19").WrapText = True
    ws.Range("C19:D19").WrapText = True

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Range(ws.Columns("C"), ws.Columns("G")).ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 14
    ws.Rows(19).RowHeight = 34

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
End Sub


' スライド計算表の1セルへの参照文字列を作る
Private Function SlideRef(ByVal sl As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If r = 0 Then
        SlideRef = "0"
    Else
        SlideRef = "'" & SH_SLIDE & "'!" & sl.Cells(r, c).Address(True, True)
    End If
End Function


'==============================================================
' 《M50_Tokushu》 特殊集計区分一覧表CSVから処分費・管材費を自動入力
'==============================================================


'==============================================================
' ① 構造チェック（金額・数量・単価は出しません。そのまま送れます）
'==============================================================
Public Sub 特殊集計区分CSV構造チェック()
    Dim v As Variant
    Dim rows As Collection
    Dim arr() As String
    Dim ws As Worksheet
    Dim lay As Object, cfg As Object
    Dim i As Long, c As Long, r As Long, maxC As Long, kcol As Long, head0 As Long
    Dim charSetName As String
    Dim curKubun As String, kubun As String
    Dim carry As Boolean, isLabel As Boolean

    v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , _
                                    "特殊集計区分一覧表のCSVを選択してください（前・後どちらでも）")
    If VarType(v) = vbBoolean Then Exit Sub

    charSetName = "Shift_JIS"
    If SheetExists(SH_CONFIG) Then
        Set cfg = GetConfig()
        charSetName = CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))
    End If

    Application.ScreenUpdating = False
    Set rows = ParseCsvText(ReadTextFile(CStr(v), charSetName), ",")
    If rows.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "CSVが空でした。", vbExclamation
        Exit Sub
    End If

    Set lay = TkDetect(rows, TkSpecOf(cfg))
    maxC = CLng(lay("列数"))
    kcol = TkKubunCol(lay)
    head0 = CLng(lay("見出し行"))
    carry = CBool(lay("区分は見出し行"))

    Set ws = FreshSheet(SH_TK_CHECK)
    ws.Range("A1").Value = "特殊集計区分CSV 構造チェック（金額・数量・単価は出していません）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "ファイル：" & Mid$(CStr(v), InStrRev(CStr(v), Application.PathSeparator) + 1)
    ws.Range("A3").Value = "読み取った並び：" & Replace(TkLayText(lay), vbCrLf, "　")

    ws.Range("A5").Value = "CSV行"
    ws.Range("B5").Value = "行の扱い"
    ws.Range("C5").Value = "効いている区分"
    ws.Range("D5").Value = "判定"
    For c = 1 To maxC
        ws.Cells(5, c + 4).Value = "第" & c & "列" & TkColRole(lay, c)
    Next c
    With ws.Range(ws.Cells(5, 1), ws.Cells(5, maxC + 4))
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
        .WrapText = True
    End With

    r = 5
    curKubun = ""
    For i = 1 To rows.Count
        arr = rows(i)
        isLabel = False
        kubun = Trim$(ColVal(arr, kcol))

        If i <= head0 Then
            isLabel = True
        ElseIf carry Then
            If kubun <> "" Then
                If TkDataCells(arr, maxC, kcol) < 3 Then isLabel = True
            End If
            If isLabel Then curKubun = kubun
        Else
            curKubun = kubun
        End If

        r = r + 1
        ws.Cells(r, 1).Value = i
        If i <= head0 Then
            ws.Cells(r, 2).Value = "見出し行"
        ElseIf isLabel Then
            ws.Cells(r, 2).Value = "区分の見出し"
        Else
            ws.Cells(r, 2).Value = "明細"
        End If
        ws.Cells(r, 3).Value = IIf(i <= head0, "", curKubun)
        ws.Cells(r, 4).Value = IIf(i <= head0, "", _
            TkKindName(TkKind(curKubun, TkWordsShobun(cfg), TkWordsKanzai(cfg))))

        For c = 1 To maxC
            If TkIsMoneyCol(lay, c) And i > head0 Then
                ws.Cells(r, c + 4).Value = IIf(Trim$(ColVal(arr, c)) = "", "", "（数値）")
            Else
                ws.Cells(r, c + 4).Value = "'" & Trim$(ColVal(arr, c))
            End If
        Next c

        If i <= head0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, maxC + 4)).Interior.Color = RGB(221, 235, 247)
        ElseIf isLabel Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, maxC + 4)).Interior.Color = RGB(255, 242, 204)
        End If
    Next i

    ws.Range(ws.Cells(5, 1), ws.Cells(r, maxC + 4)).Borders.LineStyle = xlContinuous
    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 18
    ws.Columns(4).ColumnWidth = 9
    For c = 1 To maxC
        ws.Columns(c + 4).ColumnWidth = 20
    Next c
    ws.Rows(5).RowHeight = 30

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A5").Select

    MsgBox "「" & SH_TK_CHECK & "」シートに " & rows.Count & " 行を書き出しました。" & vbCrLf & vbCrLf & _
           "読み取った並び" & vbCrLf & TkLayText(lay) & vbCrLf & vbCrLf & _
           "金額・数量・単価は「（数値）」に置き換えてあります。" & vbCrLf & _
           "このシートをそのままコピーして送れば、列の並びを合わせられます。", vbInformation
End Sub


'==============================================================
' ② 取込（前・後の2本を読んで O・P・AA・AB 列に入れる）
'==============================================================
Public Sub 特殊集計区分CSV取込()
    Dim ws As Worksheet, rs As Worksheet
    Dim cfg As Object, acc As Object, rowsByCode As Object, tankaIdx As Object
    Dim missBef As Collection, missAft As Collection
    Dim e As Variant, m As Variant
    Dim pathBef As String, pathAft As String
    Dim layBef As String, layAft As String
    Dim nBefRows As Long, nAftRows As Long
    Dim r As Long, rr As Long, lastRow As Long, i As Long
    Dim wSho As String, wKan As String
    Dim useKin As Boolean, doClear As Boolean
    Dim nSho As Long, nKan As Long, nKinB As Long, nKinA As Long
    Dim nSkip As Long, warnKin As Long, warnRow As Long
    Dim ans As VbMsgBoxResult

    On Error GoTo ErrHandler

    If Not SheetExists(SH_SLIDE) Then
        MsgBox "「" & SH_SLIDE & "」シートがありません。" & vbCrLf & _
               "先に［スライド計算表作成］を実行してください。", vbExclamation
        Exit Sub
    End If
    If Not SheetExists(SH_CONFIG) Then
        MsgBox "「" & SH_CONFIG & "」シートがありません。" & vbCrLf & _
               "先に［設定シート作成］を実行してください。", vbExclamation
        Exit Sub
    End If

    設定シート更新
    Set cfg = GetConfig()
    wSho = TkWordsShobun(cfg)
    wKan = TkWordsKanzai(cfg)
    useKin = (NormText(CStr(CfgVal(cfg, "特殊集計CSVから処分費額も入れる", "する"))) = NormText("する"))

    '--- CSVを2本選ぶ（後は省略できる）---
    pathBef = TkAskPath(CStr(CfgVal(cfg, "特殊集計区分CSVパス（スライド前）", _
                                    CStr(CfgVal(cfg, "特殊集計区分CSVパス", "")))), _
                        "【スライド前】特殊集計区分一覧表のCSVを選択してください")
    If pathBef = "" Then Exit Sub

    pathAft = TkAskPath(CStr(CfgVal(cfg, "特殊集計区分CSVパス（スライド後）", "")), _
                        "【スライド後】特殊集計区分一覧表のCSVを選択してください" & _
                        "（無ければキャンセル。AB列は空のままになり、AA列の額を使います）")

    ans = MsgBox("いま入っている O列（処分費〇）・P列（管材費〇）・" & _
                 "AA列・AB列（処分費額）を" & vbCrLf & _
                 "消してから入れ直しますか？" & vbCrLf & vbCrLf & _
                 "［はい］消してから入れる　／　［いいえ］消さずに足す", _
                 vbQuestion + vbYesNoCancel)
    If ans = vbCancel Then Exit Sub
    doClear = (ans = vbYes)

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SH_SLIDE)

    lastRow = ws.Cells(ws.Rows.Count, SC_CODE).End(xlUp).Row
    If lastRow < FIRST_ROW Then lastRow = FIRST_ROW
    Set rowsByCode = TkRowsByCode(ws, lastRow)
    Set tankaIdx = TkTankaIndex()

    If rowsByCode.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "スライド計算表に明細行が見つかりませんでした。" & vbCrLf & _
               "先に［スライド計算表作成］を実行してください。", vbExclamation
        Exit Sub
    End If

    Set acc = CreateObject("Scripting.Dictionary")
    acc.CompareMode = 1
    Set missBef = New Collection
    Set missAft = New Collection

    TkReadOne ws, cfg, pathBef, True, rowsByCode, tankaIdx, acc, missBef, wSho, wKan, layBef, nBefRows
    If pathAft <> "" Then
        TkReadOne ws, cfg, pathAft, False, rowsByCode, tankaIdx, acc, missAft, wSho, wKan, layAft, nAftRows
    End If

    '--- 結果シート ---
    Set rs = FreshSheet(SH_TK_RESULT)
    rs.Range("A1").Value = "特殊集計区分CSV 取込結果"
    rs.Range("A1").Font.Bold = True
    rs.Range("A2").Value = "スライド前：" & Mid$(pathBef, InStrRev(pathBef, Application.PathSeparator) + 1) & _
                           "　（" & nBefRows & "行）　" & Replace(layBef, vbCrLf, "　")
    rs.Range("A3").Value = "スライド後：" & _
        IIf(pathAft = "", "（指定なし。AB列は空のままにします）", _
            Mid$(pathAft, InStrRev(pathAft, Application.PathSeparator) + 1) & _
            "　（" & nAftRows & "行）　" & Replace(layAft, vbCrLf, "　"))

    rs.Range("A5").Value = "計算表の行"
    rs.Range("B5").Value = "突合コード"
    rs.Range("C5").Value = "当て方"
    rs.Range("D5").Value = "計算表の名称"
    rs.Range("E5").Value = "集計区分名称"
    rs.Range("F5").Value = "判定"
    rs.Range("G5").Value = "処分費等の名称"
    rs.Range("H5").Value = "前 件数"
    rs.Range("I5").Value = "前 金額→AA"
    rs.Range("J5").Value = "後 件数"
    rs.Range("K5").Value = "後 金額→AB"
    rs.Range("L5").Value = "結果・注意"
    With rs.Range("A5:L5")
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
        .WrapText = True
    End With
    rr = 5

    '--- 計算表へ書き込む ---
    If doClear Then
        For r = FIRST_ROW To lastRow
            If TkIsDetail(ws, r) Then
                ws.Cells(r, SC_MK_SHOBU).ClearContents
                ws.Cells(r, SC_MK_KANZA).ClearContents
                ws.Cells(r, SC_SHA_O).ClearContents
                ws.Cells(r, SC_SHA_N).ClearContents
            End If
        Next r
    End If

    For r = FIRST_ROW To lastRow
        If acc.Exists(CStr(r)) Then
            e = acc(CStr(r))

            rr = rr + 1
            rs.Cells(rr, 1).Value = r
            rs.Cells(rr, 2).Value = "'" & CStr(e(A_CODE))
            rs.Cells(rr, 3).Value = e(A_WHICH)
            rs.Cells(rr, 4).Value = ws.Cells(r, SC_NAME).Value
            rs.Cells(rr, 5).Value = e(A_KUBUN)
            rs.Cells(rr, 6).Value = TkKindName(CStr(e(A_KIND)))
            rs.Cells(rr, 7).Value = e(A_NAMES)
            If CLng(e(A_NB)) > 0 Then rs.Cells(rr, 8).Value = CLng(e(A_NB))
            If CBool(e(A_HASB)) Then rs.Cells(rr, 9).Value = CDbl(e(A_BEF))
            If CLng(e(A_NA)) > 0 Then rs.Cells(rr, 10).Value = CLng(e(A_NA))
            If CBool(e(A_HASA)) Then rs.Cells(rr, 11).Value = CDbl(e(A_AFT))
            rs.Range(rs.Cells(rr, 1), rs.Cells(rr, 12)).Borders.LineStyle = xlContinuous

            If CStr(e(A_KIND)) = "処分" Then
                If IsScrapName(CStr(e(A_NAMES))) Or IsScrapName(CStr(ws.Cells(r, SC_NAME).Value)) Then
                    rs.Cells(rr, 12).Value = "スクラップなので処分費にしませんでした"
                    rs.Cells(rr, 12).Font.Color = RGB(192, 0, 0)
                    nSkip = nSkip + 1
                Else
                    ws.Cells(r, SC_MK_SHOBU).Value = "〇"
                    nSho = nSho + 1

                    If useKin And CBool(e(A_HASB)) Then
                        ws.Cells(r, SC_SHA_O).Value = CDbl(e(A_BEF))
                        nKinB = nKinB + 1
                        rs.Cells(rr, 12).Value = "O列に〇／AA列に金額"
                    Else
                        rs.Cells(rr, 12).Value = "O列に〇。★CSVに金額が無いので、AA列に処分費の額を手で入れてください"
                        rs.Cells(rr, 12).Font.Color = RGB(192, 0, 0)
                        rs.Cells(rr, 9).Interior.Color = CLR_INPUT
                        warnKin = warnKin + 1
                    End If

                    If useKin And CBool(e(A_HASA)) Then
                        ws.Cells(r, SC_SHA_N).Value = CDbl(e(A_AFT))
                        nKinA = nKinA + 1
                        rs.Cells(rr, 12).Value = CStr(rs.Cells(rr, 12).Value) & "／AB列に金額"
                    End If
                End If
            ElseIf CStr(e(A_KIND)) = "管材" Then
                ws.Cells(r, SC_MK_KANZA).Value = "〇"
                nKan = nKan + 1
                rs.Cells(rr, 12).Value = "P列に〇"
            End If

            If CStr(e(A_NOTE)) <> "" Then
                rs.Cells(rr, 12).Value = CStr(rs.Cells(rr, 12).Value) & "　" & CStr(e(A_NOTE))
                rs.Cells(rr, 3).Interior.Color = CLR_INPUT
                warnRow = warnRow + 1
            End If
        End If
    Next r

    '--- 突合できなかったCSV行 ---
    If missBef.Count + missAft.Count > 0 Then
        rr = rr + 2
        rs.Cells(rr, 1).Value = "▼ 計算表に突合先が見つからなかったCSV行（前 " & missBef.Count & _
            "行／後 " & missAft.Count & "行）　…　下のコードがスライド計算表のA列に無いということです"
        rs.Cells(rr, 1).Font.Bold = True
        rs.Cells(rr, 1).Font.Color = RGB(192, 0, 0)
        rr = rr + 1
        rs.Cells(rr, 1).Value = "前／後"
        rs.Cells(rr, 2).Value = "CSV行"
        rs.Cells(rr, 3).Value = "さがしたコード"
        rs.Cells(rr, 5).Value = "集計区分名称"
        rs.Cells(rr, 6).Value = "判定"
        rs.Cells(rr, 7).Value = "処分費等の名称"
        rs.Cells(rr, 8).Value = "親施工数量"
        rs.Cells(rr, 9).Value = "金額"
        rs.Range(rs.Cells(rr, 1), rs.Cells(rr, 12)).Font.Bold = True
        rs.Range(rs.Cells(rr, 1), rs.Cells(rr, 12)).Interior.Color = CLR_HEAD

        For i = 1 To missBef.Count
            m = missBef(i)
            rr = rr + 1
            TkMissRow rs, rr, m
        Next i
        For i = 1 To missAft.Count
            m = missAft(i)
            rr = rr + 1
            TkMissRow rs, rr, m
        Next i
    End If

    rs.Columns("A").ColumnWidth = 10
    rs.Columns("B").ColumnWidth = 24
    rs.Columns("C").ColumnWidth = 12
    rs.Columns("D").ColumnWidth = 24
    rs.Columns("E").ColumnWidth = 18
    rs.Columns("F").ColumnWidth = 8
    rs.Columns("G").ColumnWidth = 28
    rs.Columns("H").ColumnWidth = 7
    rs.Columns("I").ColumnWidth = 13
    rs.Columns("J").ColumnWidth = 7
    rs.Columns("K").ColumnWidth = 13
    rs.Columns("L").ColumnWidth = 52
    rs.Columns("I").NumberFormat = "#,##0"
    rs.Columns("K").NumberFormat = "#,##0"
    rs.Rows(5).RowHeight = 30

    Application.ScreenUpdating = True

    If acc.Count = 0 Then
        rs.Activate
        MsgBox "1行も取り込めませんでした。" & vbCrLf & vbCrLf & _
               "スライド前：" & layBef & vbCrLf & vbCrLf & _
               IIf(pathAft = "", "", "スライド後：" & layAft & vbCrLf & vbCrLf) & _
               "★が付いていれば、その列が読み取れていません。" & vbCrLf & _
               "「" & SH_CONFIG & "」シートの「特殊集計CSV列指定」に" & vbCrLf & _
               "　区分名称=5,コード=1,名称=2,金額=3,代価表コード=6,親コード=7,親数量=8" & vbCrLf & _
               "のように書いてから、もう一度実行してください。" & vbCrLf & vbCrLf & _
               "★が無ければ、コードが計算表のA列に無いということです。" & vbCrLf & _
               "「" & SH_TK_RESULT & "」シートの「さがしたコード」を見てください。", vbExclamation
        Exit Sub
    End If

    ws.Activate

    MsgBox "特殊集計区分CSVを取り込みました。" & vbCrLf & vbCrLf & _
           "処分費〇（O列）　：" & nSho & " 行" & vbCrLf & _
           "　AA列に金額（前）：" & nKinB & " 行" & vbCrLf & _
           "　AB列に金額（後）：" & nKinA & " 行" & vbCrLf & _
           IIf(warnKin > 0, "　★金額が無かった　：" & warnKin & " 行（AA列を手入力してください）" & vbCrLf, "") & _
           "管材費〇（P列）　：" & nKan & " 行" & vbCrLf & _
           "スクラップで除外　：" & nSkip & " 行" & vbCrLf & _
           "突合できなかった　：前 " & missBef.Count & " 行／後 " & missAft.Count & " 行" & vbCrLf & _
           IIf(warnRow > 0, "当て方に注意がある行：" & warnRow & " 行" & vbCrLf, "") & vbCrLf & _
           "内容は「" & SH_TK_RESULT & "」シートで確認できます。" & vbCrLf & _
           IIf(missBef.Count + missAft.Count > 0 Or warnKin > 0 Or warnRow > 0, _
               "赤字・黄色の行は手で直してください。", "すべて取り込めました。"), vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "内容：" & Err.Description & "（" & Err.Number & "）", vbCritical
End Sub


'==============================================================
' 内部処理
'==============================================================

Private Sub TkMissRow(ByVal rs As Worksheet, ByVal rr As Long, ByVal m As Variant)
    rs.Cells(rr, 1).Value = m(0)
    rs.Cells(rr, 2).Value = m(1)
    rs.Cells(rr, 3).Value = "'" & CStr(m(2))
    rs.Cells(rr, 5).Value = m(3)
    rs.Cells(rr, 6).Value = TkKindName(CStr(m(4)))
    rs.Cells(rr, 7).Value = m(5)
    If CDbl(m(6)) <> 0 Then rs.Cells(rr, 8).Value = CDbl(m(6))
    If CDbl(m(7)) <> 0 Then rs.Cells(rr, 9).Value = CDbl(m(7))
    rs.Cells(rr, 12).Value = "突合先が見つかりませんでした"
    rs.Cells(rr, 12).Font.Color = RGB(192, 0, 0)
    rs.Range(rs.Cells(rr, 1), rs.Cells(rr, 12)).Borders.LineStyle = xlContinuous
End Sub


Private Function TkAskPath(ByVal presetPath As String, ByVal caption As String) As String
    Dim v As Variant
    If presetPath <> "" Then
        If Dir(presetPath) <> "" Then TkAskPath = presetPath: Exit Function
    End If
    v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , caption)
    If VarType(v) = vbBoolean Then Exit Function
    TkAskPath = CStr(v)
End Function


' スライド計算表の明細行を「コード → 行番号の一覧」にする
Private Function TkRowsByCode(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim d As Object, c As Collection
    Dim r As Long
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set TkRowsByCode = d

    For r = FIRST_ROW To lastRow
        If TkIsDetail(ws, r) Then
            k = NormText(CStr(ws.Cells(r, SC_CODE).Value))
            If k <> "" Then
                If Not d.Exists(k) Then d.Add k, New Collection
                Set c = d(k)
                c.Add r
            End If
        End If
    Next r
End Function


' 一覧表CSVを1本読んで、計算表の行ごとに集計する
Private Sub TkReadOne(ByVal ws As Worksheet, ByVal cfg As Object, ByVal filePath As String, _
                      ByVal isBefore As Boolean, ByVal rowsByCode As Object, _
                      ByVal tankaIdx As Object, _
                      ByVal acc As Object, ByVal miss As Collection, _
                      ByVal wSho As String, ByVal wKan As String, _
                      ByRef layText As String, ByRef nDataRows As Long)
    Dim rows As Collection
    Dim lay As Object
    Dim arr() As String, e As Variant
    Dim i As Long, maxC As Long, kcol As Long, head0 As Long
    Dim carry As Boolean, isLabel As Boolean, hasQty As Boolean
    Dim curKubun As String, kubun As String, kind As String
    Dim daika As String, oya As String, baseCode As String, nm As String
    Dim tgtCode As String, which As String, cands As String, note As String, n1 As String
    Dim qty As Double, amt As Double
    Dim tgtRow As Long
    Dim key As String
    Dim whichLabel As String

    Set rows = ParseCsvText(ReadTextFile(filePath, CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))), _
                            CStr(CfgVal(cfg, "区切り文字", ",")))
    If rows.Count = 0 Then
        layText = "（CSVが空でした）"
        Exit Sub
    End If

    Set lay = TkDetect(rows, TkSpecOf(cfg))
    layText = TkLayText(lay)

    If TkCol(lay, "代価表コード") = 0 And TkCol(lay, "親コード") = 0 _
       And TkCol(lay, "コード") = 0 Then
        layText = "★コードの列が読み取れませんでした　" & layText
        Exit Sub
    End If
    If TkKubunCol(lay) = 0 Then
        layText = "★特殊集計区分の列が読み取れませんでした　" & layText
        Exit Sub
    End If
    If TkCol(lay, "名称") = 0 Then
        layText = "★名称の列が読み取れませんでした　" & layText
        Exit Sub
    End If

    maxC = CLng(lay("列数"))
    kcol = TkKubunCol(lay)
    head0 = CLng(lay("見出し行"))
    carry = CBool(lay("区分は見出し行"))
    whichLabel = IIf(isBefore, "前", "後")
    curKubun = ""

    For i = head0 + 1 To rows.Count
        arr = rows(i)
        isLabel = False
        kubun = Trim$(ColVal(arr, kcol))

        If carry Then
            If kubun <> "" Then
                If TkDataCells(arr, maxC, kcol) < 3 Then isLabel = True
            End If
            If isLabel Then curKubun = kubun
        Else
            curKubun = kubun
        End If

        If Not isLabel Then
            kind = TkKind(curKubun, wSho, wKan)

            If kind <> "" Then
                nDataRows = nDataRows + 1
                nm = Trim$(ColVal(arr, TkCol(lay, "名称")))
                daika = Trim$(ColVal(arr, TkCol(lay, "代価表コード")))
                oya = Trim$(ColVal(arr, TkCol(lay, "親コード")))
                baseCode = Trim$(ColVal(arr, TkCol(lay, "コード")))
                qty = ToNum(ColVal(arr, TkCol(lay, "親数量")))
                If TkCol(lay, "親数量") = 0 Then qty = ToNum(ColVal(arr, TkCol(lay, "数量")))
                amt = ToNum(ColVal(arr, TkCol(lay, "金額")))

                '--- 内訳の行をコードでたどる（F列 → G列 → 基礎単価コード）---
                tgtRow = 0
                tgtCode = ""
                which = ""
                cands = ""
                note = ""

                ' 1) F列 代価表等コード番号（親施工数量は代価表の中の数量なので使わない）
                If TkIsCodeish(daika) Then
                    cands = "F列:" & daika
                    tgtRow = TkResolve(ws, rowsByCode, tankaIdx, daika, 0, False, "", 0, n1, tgtCode)
                    If tgtRow > 0 Then which = "F列 代価表": note = n1
                End If

                ' 2) G列 親施工単価コード番号
                If tgtRow = 0 And TkIsCodeish(oya) Then
                    If cands <> "" Then cands = cands & " → "
                    cands = cands & "G列:" & oya
                    ' 基礎単価がそのまま内訳に出ている場合に備えて名称も手がかりにする
                    tgtRow = TkResolve(ws, rowsByCode, tankaIdx, oya, qty, True, _
                                       IIf(NormText(oya) = NormText(baseCode), nm, ""), 0, n1, tgtCode)
                    If tgtRow > 0 Then
                        which = "G列 親施工単価"
                        note = n1
                        If TkIsCodeish(daika) Then
                            note = note & "※F列(" & daika & ")では当たらないのでG列で当てました"
                        End If
                    End If
                End If

                ' 3) 基礎単価コード（名称も手がかりに使う）
                If tgtRow = 0 And TkIsCodeish(baseCode) Then
                    If cands <> "" Then cands = cands & " → "
                    cands = cands & "基礎単価:" & baseCode
                    tgtRow = TkResolve(ws, rowsByCode, tankaIdx, baseCode, qty, True, nm, 0, n1, tgtCode)
                    If tgtRow > 0 Then
                        which = "基礎単価"
                        note = n1 & "※基礎単価コードで当てました"
                    End If
                End If

                If tgtRow > 0 Then
                    key = CStr(tgtRow)
                    If acc.Exists(key) Then
                        e = acc(key)
                        If CStr(e(A_KIND)) <> "処分" Then e(A_KIND) = kind
                        If nm <> "" Then
                            If InStr(1, CStr(e(A_NAMES)), nm) = 0 Then _
                                e(A_NAMES) = CStr(e(A_NAMES)) & "／" & nm
                        End If
                        If note <> "" Then
                            If InStr(1, CStr(e(A_NOTE)), note) = 0 Then _
                                e(A_NOTE) = Trim$(CStr(e(A_NOTE)) & " " & note)
                        End If
                        If isBefore Then
                            e(A_BEF) = CDbl(e(A_BEF)) + amt
                            e(A_NB) = CLng(e(A_NB)) + 1
                            e(A_CSVB) = TkJoin(CStr(e(A_CSVB)), CStr(i))
                            If amt <> 0 Then e(A_HASB) = True
                        Else
                            e(A_AFT) = CDbl(e(A_AFT)) + amt
                            e(A_NA) = CLng(e(A_NA)) + 1
                            e(A_CSVA) = TkJoin(CStr(e(A_CSVA)), CStr(i))
                            If amt <> 0 Then e(A_HASA) = True
                        End If
                        acc(key) = e
                    Else
                        ReDim e(0 To A_COUNT - 1)
                        e(A_KIND) = kind
                        e(A_BEF) = 0
                        e(A_AFT) = 0
                        e(A_CSVB) = ""
                        e(A_CSVA) = ""
                        e(A_NAMES) = nm
                        e(A_NB) = 0
                        e(A_NA) = 0
                        e(A_NOTE) = note
                        e(A_KUBUN) = curKubun
                        e(A_CODE) = tgtCode
                        e(A_WHICH) = which
                        e(A_HASB) = False
                        e(A_HASA) = False
                        If isBefore Then
                            e(A_BEF) = amt
                            e(A_NB) = 1
                            e(A_CSVB) = CStr(i)
                            e(A_HASB) = (amt <> 0)
                        Else
                            e(A_AFT) = amt
                            e(A_NA) = 1
                            e(A_CSVA) = CStr(i)
                            e(A_HASA) = (amt <> 0)
                        End If
                        acc.Add key, e
                    End If
                Else
                    miss.Add Array(whichLabel, i, cands, curKubun, kind, nm, qty, amt)
                End If
            End If
        End If
    Next i
End Sub


Private Function TkJoin(ByVal a As String, ByVal b As String) As String
    If a = "" Then TkJoin = b Else TkJoin = a & "," & b
End Function


' 「_単価表」シート（代価表の中身の索引）を読む
'   コード → Array(所属代価表, 数量) の一覧
Private Function TkTankaIndex() As Object
    Dim d As Object, c As Collection
    Dim ws As Worksheet
    Dim r As Long, lastR As Long
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set TkTankaIndex = d

    If Not SheetExists(SH_TANKA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_TANKA)

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastR
        k = NormText(CStr(ws.Cells(r, 1).Value))
        If k <> "" Then
            If Not d.Exists(k) Then d.Add k, New Collection
            Set c = d(k)
            c.Add Array(NormText(CStr(ws.Cells(r, 2).Value)), ToNum(CStr(ws.Cells(r, 5).Value)))
        End If
    Next r
End Function


' コードから内訳の行を決める
'   a) 内訳に数量の合う行があればそれ
'   b) 単価表部に数量の合う行があれば、その所属代価表までさかのぼって同じことをする
'   c) 内訳にそのコードの行があればそれ（名称→数量→先頭の順で絞る）
'   d) 単価表部にその コードが1件だけあれば、その所属代価表までさかのぼる
Private Function TkResolve(ByVal ws As Worksheet, ByVal rowsByCode As Object, _
                           ByVal tankaIdx As Object, ByVal code As String, _
                           ByVal qty As Double, ByVal useQty As Boolean, _
                           ByVal nmHint As String, ByVal depth As Long, _
                           ByRef note As String, ByRef hitCode As String) As Long
    Dim r As Long
    Dim owner As String
    Dim n2 As String

    note = ""
    If depth > 5 Then Exit Function
    If Not TkIsCodeish(code) Then Exit Function

    ' a) 内訳に数量の合う行
    If useQty And qty <> 0 Then
        r = TkRowByQty(ws, rowsByCode, code, qty)
        If r > 0 Then hitCode = code: TkResolve = r: Exit Function
    End If

    ' b) 単価表部に数量の合う行 → その所属代価表へ
    If useQty And qty <> 0 Then
        owner = TkOwnerOf(tankaIdx, code, qty, True)
        If owner <> "" Then
            r = TkResolve(ws, rowsByCode, tankaIdx, owner, 0, False, "", depth + 1, n2, hitCode)
            If r > 0 Then
                note = "※" & code & "（数量" & qty & "）は代価表" & owner & _
                       "の中にあるので、その行に入れました" & n2
                TkResolve = r
                Exit Function
            End If
        End If
    End If

    ' c) 内訳にそのコードの行
    r = TkPickRow(ws, rowsByCode, code, qty, useQty, nmHint, n2)
    If r > 0 Then
        hitCode = code
        note = n2
        TkResolve = r
        Exit Function
    End If

    ' d) 単価表部にそのコードが1件だけ → その所属代価表へ
    owner = TkOwnerOf(tankaIdx, code, 0, False)
    If owner <> "" Then
        r = TkResolve(ws, rowsByCode, tankaIdx, owner, 0, False, "", depth + 1, n2, hitCode)
        If r > 0 Then
            note = "※" & code & " は代価表" & owner & "の中にあるので、その行に入れました" & n2
            TkResolve = r
        End If
    End If
End Function


' 単価表部で、そのコードが入っている代価表のコードを返す
'   exact=True のときは数量も一致するものだけ。候補が複数あって決められなければ空
Private Function TkOwnerOf(ByVal tankaIdx As Object, ByVal code As String, _
                           ByVal qty As Double, ByVal exact As Boolean) As String
    Dim c As Collection
    Dim e As Variant
    Dim i As Long
    Dim k As String, found As String

    If tankaIdx Is Nothing Then Exit Function
    k = NormText(code)
    If k = "" Then Exit Function
    If Not tankaIdx.Exists(k) Then Exit Function

    Set c = tankaIdx(k)

    If exact Then
        For i = 1 To c.Count
            e = c(i)
            If Abs(CDbl(e(1)) - qty) < 0.001 Then
                If found = "" Then
                    found = CStr(e(0))
                ElseIf found <> CStr(e(0)) Then
                    Exit Function          ' 決められない
                End If
            End If
        Next i
        TkOwnerOf = found
        Exit Function
    End If

    For i = 1 To c.Count
        e = c(i)
        If found = "" Then
            found = CStr(e(0))
        ElseIf found <> CStr(e(0)) Then
            Exit Function                  ' 所属代価表が複数あって決められない
        End If
    Next i
    TkOwnerOf = found
End Function


' 内訳で、そのコードかつ数量が一致する行をさがす
Private Function TkRowByQty(ByVal ws As Worksheet, ByVal rowsByCode As Object, _
                            ByVal code As String, ByVal qty As Double) As Long
    Dim c As Collection
    Dim i As Long, r As Long
    Dim k As String

    k = NormText(code)
    If k = "" Then Exit Function
    If Not rowsByCode.Exists(k) Then Exit Function

    Set c = rowsByCode(k)
    For i = 1 To c.Count
        r = CLng(c(i))
        If IsNumeric(ws.Cells(r, SC_Q_ALL).Value) Then
            If Abs(CDbl(ws.Cells(r, SC_Q_ALL).Value) - qty) < 0.001 Then
                TkRowByQty = r
                Exit Function
            End If
        End If
    Next i
End Function


' コードから計算表の明細行を決める
'   同じコードが1行だけならそれ。複数あるときは 名称 → 数量 → 先頭 の順で絞る
Private Function TkPickRow(ByVal ws As Worksheet, ByVal rowsByCode As Object, _
                           ByVal code As String, ByVal qty As Double, ByVal useQty As Boolean, _
                           ByVal nmHint As String, ByRef note As String) As Long
    Dim c As Collection
    Dim i As Long, r As Long
    Dim k As String, hint As String

    note = ""
    k = NormText(code)
    If k = "" Then Exit Function
    If Not rowsByCode.Exists(k) Then Exit Function

    Set c = rowsByCode(k)
    If c.Count = 0 Then Exit Function
    If c.Count = 1 Then TkPickRow = CLng(c(1)): Exit Function

    ' 名称で絞る
    hint = NormText(nmHint)
    If hint <> "" Then
        For i = 1 To c.Count
            r = CLng(c(i))
            If NormText(CStr(ws.Cells(r, SC_NAME).Value)) = hint Then TkPickRow = r: Exit Function
        Next i
    End If

    ' 数量で絞る
    If useQty And qty <> 0 Then
        For i = 1 To c.Count
            r = CLng(c(i))
            If IsNumeric(ws.Cells(r, SC_Q_ALL).Value) Then
                If Abs(CDbl(ws.Cells(r, SC_Q_ALL).Value) - qty) < 0.001 Then
                    TkPickRow = r
                    Exit Function
                End If
            End If
        Next i
    End If

    note = "※同じコードの明細が" & c.Count & "行あります。先頭の行に入れました（要確認）"
    TkPickRow = CLng(c(1))
End Function


' 明細行かどうか（AC列の処分費額の式は明細行だけが持っている）
Private Function TkIsDetail(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    If Not ws.Cells(r, SC_SA_ALL).HasFormula Then Exit Function
    If NormText(CStr(ws.Cells(r, SC_NAME).Value)) = "" Then Exit Function
    TkIsDetail = True
End Function


' 区分の列以外に、中身の入ったセルがいくつあるか（明細行かどうかの目印）
Private Function TkDataCells(ByRef arr() As String, ByVal maxC As Long, _
                             ByVal skipCol As Long) As Long
    Dim c As Long, n As Long
    For c = 1 To maxC
        If c <> skipCol Then
            If Trim$(ColVal(arr, c)) <> "" Then n = n + 1
        End If
    Next c
    TkDataCells = n
End Function


' 区分名から「処分」「管材」を判定する
Private Function TkKind(ByVal s As String, ByVal wSho As String, ByVal wKan As String) As String
    If NormText(s) = "" Then Exit Function
    If InStr(1, NormText(s), NormText("区分外")) > 0 Then Exit Function
    If TkHasWord(s, wSho) Then TkKind = "処分": Exit Function
    If TkHasWord(s, wKan) Then TkKind = "管材"
End Function


Private Function TkKindName(ByVal kind As String) As String
    Select Case kind
        Case "処分": TkKindName = "処分費"
        Case "管材": TkKindName = "管材費"
        Case Else:   TkKindName = "―"
    End Select
End Function


Private Function TkHasWord(ByVal s As String, ByVal words As String) As Boolean
    Dim a As Variant, w As Variant
    Dim t As String
    t = NormText(s)
    If t = "" Then Exit Function
    a = Split(Replace(words, "、", ","), ",")
    For Each w In a
        If NormText(CStr(w)) <> "" Then
            If InStr(1, t, NormText(CStr(w))) > 0 Then TkHasWord = True: Exit Function
        End If
    Next w
End Function


Private Function TkWordsShobun(ByVal cfg As Object) As String
    If cfg Is Nothing Then TkWordsShobun = "処分": Exit Function
    TkWordsShobun = CStr(CfgVal(cfg, "処分費の区分名", "処分"))
End Function


Private Function TkWordsKanzai(ByVal cfg As Object) As String
    If cfg Is Nothing Then TkWordsKanzai = "水道,管材,管財": Exit Function
    TkWordsKanzai = CStr(CfgVal(cfg, "管材費の区分名", "水道,管材,管財"))
End Function


Private Function TkSpecOf(ByVal cfg As Object) As String
    If cfg Is Nothing Then Exit Function
    TkSpecOf = CStr(CfgVal(cfg, "特殊集計CSV列指定", "自動"))
End Function


' 判定に使う区分の列（「集計区分名称」があればそちら。無ければ「集計区分」）
Private Function TkKubunCol(ByVal lay As Object) As Long
    If TkCol(lay, "区分名称") > 0 Then
        TkKubunCol = TkCol(lay, "区分名称")
    Else
        TkKubunCol = TkCol(lay, "区分")
    End If
End Function


'--------------------------------------------------------------
' 列の並びを読み取る
'   戻り値は Dictionary。"区分" "区分名称" "コード" "名称" "単位" "規格1"
'   "規格2" "数量" "単価" "金額" "摘要" "代価表コード" "親コード" "親数量"
'   に列番号（0＝無し）、ほかに "見出し行"／"列数"／"区分は見出し行"。
'--------------------------------------------------------------
Private Function TkDetect(ByVal rows As Collection, ByVal spec As String) As Object
    Dim lay As Object, ov As Object
    Dim arr() As String
    Dim ks As Variant, kk As Variant
    Dim i As Long, c As Long, maxC As Long, head0 As Long
    Dim hit As Long, bestHit As Long, bestRow As Long
    Dim t As String, v As String
    Dim cntK() As Long, best As Long, bestCol As Long
    Dim nOnly As Long, nWith As Long, nData As Long
    Dim sCode() As Long, sText() As Long, sTani() As Long, sNum() As Long
    Dim used() As Boolean
    Dim cc As Long, tc As Long, nc As Long, k1c As Long, k2c As Long
    Dim numCols() As Long, nNum As Long

    Set lay = CreateObject("Scripting.Dictionary")
    lay.CompareMode = 1
    ks = TkSlots()
    For Each kk In ks
        lay(CStr(kk)) = 0
    Next kk
    lay("見出し行") = 0
    lay("区分は見出し行") = False

    '--- 列数 ---
    maxC = 0
    For i = 1 To rows.Count
        arr = rows(i)
        If UBound(arr) + 1 > maxC Then maxC = UBound(arr) + 1
    Next i
    If maxC < 1 Then maxC = 1
    lay("列数") = maxC

    '--- 1) 見出し行をさがす（先頭20行まで）---
    bestHit = 0
    bestRow = 0
    For i = 1 To rows.Count
        If i > 20 Then Exit For
        arr = rows(i)
        hit = 0
        For c = 1 To maxC
            If TkHeadName(ColVal(arr, c)) <> "" Then hit = hit + 1
        Next c
        If hit > bestHit Then
            bestHit = hit
            bestRow = i
        End If
    Next i
    If bestHit >= 3 Then
        arr = rows(bestRow)
        For c = 1 To maxC
            t = TkHeadName(ColVal(arr, c))
            If t <> "" Then
                If CLng(lay(t)) = 0 Then lay(t) = c
            End If
        Next c
        lay("見出し行") = bestRow
    End If
    head0 = CLng(lay("見出し行"))

    '--- 2) 区分の列を中身から当てる ---
    If TkKubunCol(lay) = 0 Then
        ReDim cntK(1 To maxC)
        For i = head0 + 1 To rows.Count
            arr = rows(i)
            For c = 1 To maxC
                If TkLooksKubun(ColVal(arr, c)) Then cntK(c) = cntK(c) + 1
            Next c
        Next i
        best = 0
        bestCol = 0
        For c = 1 To maxC
            If cntK(c) > best Then best = cntK(c): bestCol = c
        Next c
        If best > 0 Then lay("区分") = bestCol
    End If

    '--- 3) 区分が見出し行として立っているか（下の行へ引き継ぐか）---
    If TkKubunCol(lay) > 0 Then
        For i = head0 + 1 To rows.Count
            arr = rows(i)
            If TkLooksKubun(ColVal(arr, TkKubunCol(lay))) Then
                If TkDataCells(arr, maxC, TkKubunCol(lay)) >= 3 Then
                    nWith = nWith + 1
                Else
                    nOnly = nOnly + 1
                End If
            End If
        Next i
        lay("区分は見出し行") = (nOnly > nWith)
    End If

    '--- 4) 見出しが無いときは、残りの列を中身から当てる ---
    If head0 = 0 Then
        ReDim sCode(1 To maxC)
        ReDim sText(1 To maxC)
        ReDim sTani(1 To maxC)
        ReDim sNum(1 To maxC)

        For i = 1 To rows.Count
            arr = rows(i)
            If TkDataCells(arr, maxC, TkKubunCol(lay)) >= 3 Then
                nData = nData + 1
                For c = 1 To maxC
                    v = Trim$(ColVal(arr, c))
                    If v <> "" Then
                        If TkIsNumish(v) Then
                            sNum(c) = sNum(c) + 1
                        Else
                            sText(c) = sText(c) + 1
                        End If
                        If TkIsCodeish(v) Then sCode(c) = sCode(c) + 1
                        If TkIsTaniish(v) Then sTani(c) = sTani(c) + 1
                    End If
                Next c
            End If
        Next i

        If nData > 0 Then
            ReDim used(1 To maxC)
            If Not CBool(lay("区分は見出し行")) Then
                If TkKubunCol(lay) > 0 Then used(TkKubunCol(lay)) = True
            End If

            ' コード：コードらしい値がいちばん多い列
            cc = 0: best = 0
            For c = 1 To maxC
                If Not used(c) Then
                    If sCode(c) > best Then best = sCode(c): cc = c
                End If
            Next c
            If cc > 0 Then
                If best * 2 >= nData Then lay("コード") = cc: used(cc) = True
            End If

            ' 単位：単位らしい値がいちばん多い列
            tc = 0: best = 0
            For c = 1 To maxC
                If Not used(c) Then
                    If sTani(c) > best Then best = sTani(c): tc = c
                End If
            Next c
            If tc > 0 Then
                If best * 10 >= nData * 3 Then lay("単位") = tc: used(tc) = True
            End If

            ' 名称：文字の入った行がいちばん多い列（同数なら左）
            nc = 0: best = 0
            For c = 1 To maxC
                If Not used(c) Then
                    If sText(c) > best Then best = sText(c): nc = c
                End If
            Next c
            If nc > 0 And best > 0 Then
                lay("名称") = nc
                used(nc) = True
                k1c = 0: k2c = 0
                For c = nc + 1 To maxC
                    If Not used(c) And sText(c) > 0 Then
                        If k1c = 0 Then
                            k1c = c
                        ElseIf k2c = 0 Then
                            k2c = c
                        End If
                    End If
                Next c
                If k1c > 0 Then lay("規格1") = k1c: used(k1c) = True
                If k2c > 0 Then lay("規格2") = k2c: used(k2c) = True
            End If

            ' 数量・単価・金額：数値の列を左から
            ReDim numCols(1 To maxC)
            For c = 1 To maxC
                If Not used(c) Then
                    If sNum(c) * 2 >= nData Then nNum = nNum + 1: numCols(nNum) = c
                End If
            Next c
            If nNum >= 3 Then
                lay("数量") = numCols(1)
                lay("単価") = numCols(2)
                lay("金額") = numCols(nNum)
            ElseIf nNum = 2 Then
                lay("数量") = numCols(1)
                lay("金額") = numCols(2)
            ElseIf nNum = 1 Then
                lay("金額") = numCols(1)
            End If
        End If
    End If

    '--- 5) 設定シートの列指定を最後にかぶせる（手で書いた指定が優先）---
    Set ov = TkParseSpec(spec)
    For Each kk In ov.Keys
        lay(CStr(kk)) = CLng(ov(CStr(kk)))
    Next kk

    Set TkDetect = lay
End Function


Private Function TkSlots() As Variant
    TkSlots = Array("区分", "区分名称", "コード", "名称", "単位", "規格1", "規格2", _
                    "数量", "単価", "金額", "摘要", "代価表コード", "親コード", "親数量")
End Function


Private Function TkCol(ByVal lay As Object, ByVal name As String) As Long
    If lay Is Nothing Then Exit Function
    If Not lay.Exists(name) Then Exit Function
    TkCol = CLng(lay(name))
End Function


' 見出しの文字から項目名を決める
'   NormText は StrConv(vbNarrow) で全角カタカナを半角にするので、
'   「コード」のようなカタカナを含む語は、必ず語の側も NormText して比べること。
'   （全角リテラルのまま比べると一致せず、親施工単価コード番号を読み落とす）
Private Function TkHeadName(ByVal s As String) As String
    Dim t As String, raw As String

    raw = Replace(Replace(Trim$(s), " ", ""), "　", "")
    If raw = "" Then Exit Function

    ' まず実物の見出し名そのままで照合する（文字コードの変換に左右されないように）
    Select Case raw
        Case "基礎単価コード":       TkHeadName = "コード": Exit Function
        Case "代価表等コード番号":   TkHeadName = "代価表コード": Exit Function
        Case "親施工単価コード番号": TkHeadName = "親コード": Exit Function
        Case "親施工数量":           TkHeadName = "親数量": Exit Function
        Case "集計区分名称":         TkHeadName = "区分名称": Exit Function
        Case "集計区分":             TkHeadName = "区分": Exit Function
        Case "特殊集計区分名称":     TkHeadName = "区分名称": Exit Function
        Case "特殊集計区分":         TkHeadName = "区分": Exit Function
        Case "名称":                 TkHeadName = "名称": Exit Function
        Case "金額":                 TkHeadName = "金額": Exit Function
        Case "単位":                 TkHeadName = "単位": Exit Function
        Case "数量":                 TkHeadName = "数量": Exit Function
    End Select

    t = NormText(s)
    If t = "" Then Exit Function

    If TkHas(t, "特殊集計") Or TkHas(t, "集計区分") Then
        ' 「集計区分」は番号、「集計区分名称」が名前。判定には名前を使う
        If TkHas(t, "名称") Then TkHeadName = "区分名称" Else TkHeadName = "区分"
        Exit Function
    End If
    If t = NormText("区分") Then TkHeadName = "区分": Exit Function
    If TkHas(t, "代価表") Then TkHeadName = "代価表コード": Exit Function
    If TkHas(t, "親") Then
        If TkHas(t, "数量") Then
            TkHeadName = "親数量"
        ElseIf TkHas(t, "コード") Or TkHas(t, "単価") Then
            TkHeadName = "親コード"
        End If
        Exit Function
    End If
    If TkHas(t, "コード") Then TkHeadName = "コード": Exit Function
    If TkHas(t, "名称") Or TkHas(t, "品名") Then TkHeadName = "名称": Exit Function
    If TkHas(t, "規格") Or TkHas(t, "仕様") Then
        If TkHas(t, "2") Then TkHeadName = "規格2" Else TkHeadName = "規格1"
        Exit Function
    End If
    If TkHas(t, "単位") Then TkHeadName = "単位": Exit Function
    If TkHas(t, "数量") Then TkHeadName = "数量": Exit Function
    If TkHas(t, "金額") Then TkHeadName = "金額": Exit Function
    If TkHas(t, "単価") Then TkHeadName = "単価": Exit Function
    If TkHas(t, "摘要") Then TkHeadName = "摘要": Exit Function
End Function


' NormText した文字列 t の中に語が入っているか（語も NormText して比べる）
Private Function TkHas(ByVal t As String, ByVal word As String) As Boolean
    Dim w As String
    w = NormText(word)
    If w = "" Then Exit Function
    TkHas = (InStr(1, t, w) > 0)
End Function


' 特殊集計区分の名前らしいか
Private Function TkLooksKubun(ByVal s As String) As Boolean
    Dim t As String
    t = NormText(s)
    If t = "" Then Exit Function
    If Len(t) > 24 Then Exit Function
    If InStr(1, t, "処分") > 0 Then TkLooksKubun = True: Exit Function
    If InStr(1, t, "水道") > 0 Then TkLooksKubun = True: Exit Function
    If InStr(1, t, "管材") > 0 Then TkLooksKubun = True: Exit Function
    If InStr(1, t, "管財") > 0 Then TkLooksKubun = True: Exit Function
    If InStr(1, t, "区分外") > 0 Then TkLooksKubun = True
End Function


' 数値らしいか
Private Function TkIsNumish(ByVal s As String) As Boolean
    Dim t As String
    t = Replace(NormText(s), ",", "")
    If t = "" Or t = "-" Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    TkIsNumish = True
End Function


' 施工単価コードらしいか（英字で始まり、数字を含む半角英数）
Private Function TkIsCodeish(ByVal s As String) As Boolean
    Dim t As String, i As Long, ch As String
    Dim hasDigit As Boolean

    t = UCase$(NormText(s))
    If Len(t) < 3 Or Len(t) > 14 Then Exit Function
    If Left$(t, 1) < "A" Or Left$(t, 1) > "Z" Then Exit Function

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            hasDigit = True
        ElseIf ch >= "A" And ch <= "Z" Then
            ' ok
        ElseIf ch = "-" Or ch = "/" Then
            ' ok
        Else
            Exit Function
        End If
    Next i

    TkIsCodeish = hasDigit
End Function


' 単位らしいか
Private Function TkIsTaniish(ByVal s As String) As Boolean
    Dim t As String
    Dim a As Variant, w As Variant

    t = NormText(s)
    If t = "" Or Len(t) > 4 Then Exit Function

    a = Split(TK_TANI_LIST, "|")
    For Each w In a
        If t = NormText(CStr(w)) Then TkIsTaniish = True: Exit Function
    Next w
End Function


' 金額・単価・数量の列か（構造チェックで数値を隠すのに使う）
Private Function TkIsMoneyCol(ByVal lay As Object, ByVal c As Long) As Boolean
    If c = 0 Then Exit Function
    If c = TkCol(lay, "金額") Then TkIsMoneyCol = True: Exit Function
    If c = TkCol(lay, "単価") Then TkIsMoneyCol = True: Exit Function
    If c = TkCol(lay, "数量") Then TkIsMoneyCol = True: Exit Function
    If c = TkCol(lay, "親数量") Then TkIsMoneyCol = True
End Function


Private Function TkColRole(ByVal lay As Object, ByVal c As Long) As String
    Dim ks As Variant, kk As Variant
    ks = TkSlots()
    For Each kk In ks
        If TkCol(lay, CStr(kk)) = c Then TkColRole = vbLf & "＝" & CStr(kk): Exit Function
    Next kk
End Function


Private Function TkLayText(ByVal lay As Object) As String
    Dim ks As Variant, kk As Variant
    Dim s As String
    ks = TkSlots()
    For Each kk In ks
        If TkCol(lay, CStr(kk)) > 0 Then
            If s <> "" Then s = s & "／"
            s = s & CStr(kk) & "=" & CStr(TkCol(lay, CStr(kk))) & "列"
        End If
    Next kk
    s = s & vbCrLf & "見出し行=" & IIf(CLng(lay("見出し行")) = 0, "なし", CStr(lay("見出し行"))) & _
        "／区分の引き継ぎ=" & IIf(CBool(lay("区分は見出し行")), "する", "しない")
    TkLayText = s
End Function


' 「区分名称=5,コード=1,…」を読む。列はA,B,…の記号でも数字でもよい
Private Function TkParseSpec(ByVal spec As String) As Object
    Dim d As Object
    Dim a As Variant, p As Variant
    Dim kv() As String
    Dim nm As String, cno As Long

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set TkParseSpec = d

    If NormText(spec) = "" Then Exit Function
    If NormText(spec) = NormText("自動") Then Exit Function

    a = Split(Replace(Replace(spec, "、", ","), "；", ","), ",")
    For Each p In a
        If InStr(1, CStr(p), "=") > 0 Then
            kv = Split(CStr(p), "=")
            nm = TkSpecName(kv(0))
            cno = TkToColNo(kv(1))
            If nm <> "" And cno > 0 Then d(nm) = cno
        End If
    Next p
End Function


Private Function TkSpecName(ByVal s As String) As String
    Dim t As String
    t = NormText(s)
    If t = "" Then Exit Function

    If t = NormText("区分") Or t = NormText("集計区分") Or t = NormText("特殊集計区分") Then
        TkSpecName = "区分": Exit Function
    End If
    If t = NormText("区分名称") Or t = NormText("集計区分名称") Then
        TkSpecName = "区分名称": Exit Function
    End If
    If t = NormText("コード") Or t = NormText("基礎単価コード") Then
        TkSpecName = "コード": Exit Function
    End If
    If t = NormText("代価表コード") Or t = NormText("代価表") Or _
       t = NormText("代価表等コード番号") Then
        TkSpecName = "代価表コード": Exit Function
    End If
    If t = NormText("親コード") Or t = NormText("親施工単価コード番号") Then
        TkSpecName = "親コード": Exit Function
    End If
    If t = NormText("親数量") Or t = NormText("親施工数量") Then
        TkSpecName = "親数量": Exit Function
    End If
    If t = NormText("名称") Then TkSpecName = "名称": Exit Function
    If t = NormText("単位") Then TkSpecName = "単位": Exit Function
    If t = NormText("規格") Or t = NormText("規格1") Then TkSpecName = "規格1": Exit Function
    If t = NormText("規格2") Then TkSpecName = "規格2": Exit Function
    If t = NormText("数量") Then TkSpecName = "数量": Exit Function
    If t = NormText("金額") Then TkSpecName = "金額": Exit Function
    If t = NormText("単価") Then TkSpecName = "単価": Exit Function
    If t = NormText("摘要") Then TkSpecName = "摘要": Exit Function
End Function


Private Function TkToColNo(ByVal s As String) As Long
    Dim t As String, i As Long, ch As String, n As Long
    t = UCase$(NormText(s))
    If t = "" Then Exit Function
    If IsNumeric(t) Then TkToColNo = CLng(t): Exit Function
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
        n = n * 26 + (Asc(ch) - 64)
    Next i
    TkToColNo = n
End Function


'==============================================================
' 《M90_SampleData》 動作確認用サンプルCSV生成
'==============================================================


Public Sub テスト用CSV作成()
    Dim baseDir As String, pOld As String, pNew As String

    baseDir = ThisWorkbook.Path
    If baseDir = "" Then
        MsgBox "先にこのブックを保存してください。", vbExclamation
        Exit Sub
    End If

    pOld = baseDir & Application.PathSeparator & "sample_旧単価.csv"
    pNew = baseDir & Application.PathSeparator & "sample_新単価.csv"

    WriteTextFile pOld, SampleCsv(False), "Shift_JIS"
    WriteTextFile pNew, SampleCsv(True), "Shift_JIS"

    If SheetExists(SH_CONFIG) Then
        SetConfigValue "旧単価CSVパス", pOld
        SetConfigValue "新単価CSVパス", pNew
        SetConfigValue "工事名", "○○地内　配水管布設工事（テスト）"
        SetConfigValue "工事場所", "神戸市○○区○○町地内"
        SetConfigValue "工期（自）", "令和7年4月1日"
        SetConfigValue "工期（至）", "令和8年3月20日"
        SetConfigValue "請負代金額", 120000000
        SetConfigValue "基準日", "令和7年10月1日"
        SetConfigValue "発注者", "神戸市"
        SetConfigValue "受注者", "株式会社○○建設"
        SetConfigValue "契約保証費", 77899
    End If

    MsgBox "サンプルCSVを作成しました。" & vbCrLf & vbCrLf & _
           pOld & vbCrLf & pNew & vbCrLf & vbCrLf & _
           "続けて［スライド計算表作成］を実行してください。", vbInformation
End Sub


Private Sub SetConfigValue(ByVal label As String, ByVal v As Variant)
    Dim ws As Worksheet
    Dim r As Long, lastR As Long

    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 1 To lastR
        If NormText(CStr(ws.Cells(r, 1).Value)) = NormText(label) Then
            ws.Cells(r, 2).Value = v
            Exit Sub
        End If
    Next r
End Sub


'--------------------------------------------------------------
' サンプル本体
'   isNew = False … 当初の単価適用日／True … 基準日の単価適用日
'--------------------------------------------------------------
Private Function SampleCsv(ByVal isNew As Boolean) As String
    mSb = ""

    ' 1行目はヘッダ行（積算システムが付ける制御行）
    mSb = "A1,8.69E+12,1,71001,80901,0,0,0,0,0,0,0,0" & vbCrLf

    '--- 単価表部（Gコードの代価表。内訳には入らない）---
    Lv "G0001", "／再掘削工", "式", "殻運搬処理　処分費含む", ""
    Det "舗装版破砕工", "㎡", "", "", 93, 93, IIf(isNew, 1420, 1350)
    Det "床掘", "", "", "", 150, 150, IIf(isNew, 336, 320)
    Lv "G0003", "／交通誘導警備員", "式", "", ""
    Det "交通誘導警備員B", "人日", "", "", 1152, 1152, IIf(isNew, 12080, 11870)
    Lv "G1005", "／仮設資材運搬", "式", "積込み取卸し含む", "往復分"
    Det "仮設材等の運搬", "t", "(鋼矢板､H形鋼等)", "往路", 2, 2, IIf(isNew, 5820, 5480)
    Lv "G1002", "／通水試験費", "式", "器具損料・諸雑費含む", ""
    Det "通水試験工", "日", "", "", 1, 1, IIf(isNew, 86400, 82300)
    Lv "G0007", "／スクラップ", "式", "撤去管・残管等", "；故銑A"
    Det "スクラップ", "ｔ", "故銑A", "", -7.9, -7.9, IIf(isNew, 23500, 9500)

    '--- 内訳部 ---
    Lv "X1000", "本工事費", "", "", ""

    Lv "Y2301", "管路(開削)", "", "", ""
    Lv "Y230101", "材料", "式", "", ""
    Lv "Y23010101", "管材料", "", "", ""
    Lv "Y2301010102", "GX形管", "", "", ""
    Det "ＧＸ形直管", "本", "継手材含む", "", 114, 114, IIf(isNew, 35690, 35400)
    Det "ＧＸ形直管", "本", "継手材含む", "", 30, 0, IIf(isNew, 35690, 35400)
    Lv "Y23010103", "ポリエチレン管", "", "", ""
    Lv "Y2301010301", "配水用ﾎﾟﾘｴﾁﾚﾝ管", "", "", ""
    Det "ＰＥ形 ＥＦ受口付直管", "本", "配水用", "", 500, 0, IIf(isNew, 8308, 8101)

    Lv "Y230102", "管布設工", "式", "", ""
    Lv "Y23010201", "管路土工", "", "", ""
    Lv "Y2301020101", "舗装版切断", "", "", ""
    Det "舗装版切断", "m", "", "", 5260, 5260, IIf(isNew, 1613, 1583)
    Lv "Y2301020110", "管路掘削", "", "", ""
    Det "床掘", "", "試掘", "", 1000, 80, IIf(isNew, 336, 320)
    Det "土砂等運搬", "m3", "現場～仮置場L=4.2km", "", 2000, 3010, IIf(isNew, 1180, 1150)
    Lv "Y2301020112", "再掘削", "", "", ""
    DetG "G0001", "／再掘削工", "式", "殻運搬処理　処分費含む", 1, 1, IIf(isNew, 452000, 410000)

    Lv "Y230104", "管工", "式", "", ""
    Lv "Y23010401", "管据付・撤去工", "", "", ""
    Lv "Y2301040101", "管据付", "", "", ""
    Det "鋳鉄管据付", "ｍ", "", "", 638.1, 638.1, IIf(isNew, 5820, 5430)

    Lv "Y230107", "付帯工", "式", "", ""
    Lv "Y23010701", "舗装撤去工", "", "", ""
    Lv "Y2301070101", "舗装版切断", "", "", ""
    Det "舗装版切断", "m", "", "", 650, 650, IIf(isNew, 1613, 1583)
    Lv "Y23010702", "道路復旧工", "", "", ""
    Lv "Y2301070201", "舗装復旧", "", "", ""
    Det "舗装復旧工", "㎡", "", "", 3120, 3120, IIf(isNew, 2240, 2080)

    Lv "Y230108", "仮設工", "式", "", ""
    Lv "Y23010803", "交通管理工", "", "", ""
    DetG "G0003", "／交通誘導警備員", "式", "", 1, 1, IIf(isNew, 13916160, 13674240)

    '--- 諸経費部 ---
    Lv "Z0001", "運搬費", "式", "", ""
    Lv "YZ000000003", "仮設材運搬費", "式", "", ""
    DetG "G1005", "／仮設資材運搬", "式", "積込み取卸し含む", 1, 1, IIf(isNew, 23280, 21920)
    Lv "Z0002", "準備費", "式", "", ""
    Lv "Z0006", "技術管理費", "式", "", ""
    DetG "G1002", "／通水試験費", "式", "器具損料・諸雑費含む", 1, 1, IIf(isNew, 86400, 82300)
    Lv "Z0013", "支給品費", "式", "", ""
    Lv "Z0040", "工期延長等に伴う増加費用", "式", "", ""
    Lv "Z0045", "設計委託費", "式", "", ""
    Lv "Z0047", "スクラップ", "式", "", ""
    DetG "G0007", "／スクラップ", "式", "撤去管・残管等", 1, 1, IIf(isNew, -185650, -75050)

    SampleCsv = mSb
End Function


' 階層行・見出し行（数量なし）
Private Sub Lv(ByVal code As String, ByVal nm As String, ByVal tani As String, _
               ByVal k1 As String, ByVal k2 As String)
    Row13 code, nm, tani, k1, k2, "", 0, 0, 0
End Sub


' 明細行（コードなし）
Private Sub Det(ByVal nm As String, ByVal tani As String, ByVal k1 As String, _
                ByVal k2 As String, ByVal qNew As Double, ByVal qOrg As Double, _
                ByVal tanka As Double)
    Row13 "", nm, tani, k1, k2, "", qNew, qOrg, tanka
End Sub


' 明細行（Gコードの代価を引く行）
Private Sub DetG(ByVal code As String, ByVal nm As String, ByVal tani As String, _
                 ByVal k1 As String, ByVal qNew As Double, ByVal qOrg As Double, _
                 ByVal tanka As Double)
    Row13 code, nm, tani, k1, "", "", qNew, qOrg, tanka
End Sub


' 13列。①＝変更設計、②＝当初設計
Private Sub Row13(ByVal code As String, ByVal nm As String, ByVal tani As String, _
                  ByVal k1 As String, ByVal k2 As String, ByVal tekiyo As String, _
                  ByVal qNew As Double, ByVal qOrg As Double, ByVal tanka As Double)
    mSb = mSb & Q(code) & "," & Q(nm) & "," & Q(tani) & "," & Q(k1) & "," & Q(k2) & "," & Q(tekiyo) & _
          "," & IIf(tanka = 0, "", Format$(tanka, "0")) & "," & IIf(tanka = 0, "", Format$(tanka, "0")) & _
          "," & IIf(qNew = 0 And qOrg = 0, "", Format$(qNew, "0.##")) & _
          "," & IIf(qNew = 0 And qOrg = 0, "", Format$(qOrg, "0.##")) & _
          "," & IIf(tanka = 0, "", Format$(qNew * tanka, "0")) & _
          "," & IIf(tanka = 0, "", Format$(qOrg * tanka, "0")) & "," & vbCrLf
End Sub


Private Function Q(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Then
        Q = """" & Replace(s, """", """""") & """"
    Else
        Q = s
    End If
End Function


'==============================================================
' 《M92_CsvCheck》 CSVの構造チェック（診断用）
'==============================================================


Public Sub CSV構造チェック()
    Dim v As Variant
    Dim rows As Collection
    Dim arr() As String
    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim cfg As Object
    Dim code As String, lvl As String, kubun As String
    Dim inZ As Boolean, inNaiyaku As Boolean
    Dim charSetName As String

    v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , _
                                    "構造を確認するCSVを選択してください")
    If VarType(v) = vbBoolean Then Exit Sub

    charSetName = "Shift_JIS"
    If SheetExists(SH_CONFIG) Then
        Set cfg = GetConfig()
        charSetName = CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))
    End If

    Application.ScreenUpdating = False
    Set rows = ParseCsvText(ReadTextFile(CStr(v), charSetName), ",")

    Set ws = FreshSheet("_CSV構造")
    ws.Range("A1").Value = "CSV構造チェック（金額・数量は出していません）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "ファイル：" & Mid$(CStr(v), InStrRev(CStr(v), Application.PathSeparator) + 1)

    ws.Range("A4").Value = "CSV行"
    ws.Range("B4").Value = "コード"
    ws.Range("C4").Value = "レベル判定"
    ws.Range("D4").Value = "行の区分"
    ws.Range("E4").Value = "施工単価名称"
    ws.Range("F4").Value = "単位"
    ws.Range("G4").Value = "規格1"
    ws.Range("H4").Value = "規格2"
    ws.Range("I4").Value = "摘要"
    With ws.Range("A4:I4")
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With

    r = 4
    inZ = False
    inNaiyaku = False

    For i = 1 To rows.Count
        arr = rows(i)
        code = Trim$(ColVal(arr, CSV_CODE))
        lvl = LevelOf(code)

        If lvl = "費目" Then inNaiyaku = True
        If lvl = "Z" Then inZ = True

        If Not inNaiyaku Then
            ' X1000より前はGコードの単価表。内訳には入れない
            kubun = "単価表部（内訳に入れません）"
            GoTo WriteLine
        End If

        Select Case lvl
            Case "G":    kubun = "内訳の明細（Gコードの代価）"
            Case "費目": kubun = "費目（本工事費）"
            Case "L1":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・工事区分"
            Case "L2":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・工種"
            Case "L3":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・種別"
            Case "L4":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・細別"
            Case "YZ":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・YZ（工種あつかい）"
            Case "Z":    kubun = "★Zコード項目　区分＝" & _
                                 IIf(ZKubunPub(code, Trim$(ColVal(arr, CSV_NAME))) = "", _
                                     "（使わない）", ZKubunPub(code, Trim$(ColVal(arr, CSV_NAME))))
            Case Else:   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・明細"
        End Select

WriteLine:

        r = r + 1
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = "'" & code
        ws.Cells(r, 3).Value = IIf(lvl = "", "（明細）", lvl)
        ws.Cells(r, 4).Value = kubun
        ws.Cells(r, 5).Value = Trim$(ColVal(arr, CSV_NAME))
        ws.Cells(r, 6).Value = Trim$(ColVal(arr, CSV_TANI))
        ws.Cells(r, 7).Value = Trim$(ColVal(arr, CSV_KIKAKU1))
        ws.Cells(r, 8).Value = Trim$(ColVal(arr, CSV_KIKAKU2))
        ws.Cells(r, 9).Value = Trim$(ColVal(arr, CSV_TEKIYO))

        If lvl = "Z" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(255, 242, 204)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font.Bold = True
        ElseIf lvl = "費目" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(221, 235, 247)
        End If
    Next i

    ws.Range(ws.Cells(4, 1), ws.Cells(r, 9)).Borders.LineStyle = xlContinuous
    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 10
    ws.Columns(4).ColumnWidth = 30
    ws.Columns(5).ColumnWidth = 28
    ws.Columns(6).ColumnWidth = 7
    ws.Columns(7).ColumnWidth = 20
    ws.Columns(8).ColumnWidth = 20
    ws.Columns(9).ColumnWidth = 20

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A4").Select

    MsgBox "「_CSV構造」シートに " & (r - 4) & " 行を書き出しました。" & vbCrLf & vbCrLf & _
           "A列からI列を選択してコピーすれば、そのまま送れます。" & vbCrLf & _
           "単価・数量・金額は含まれていません。", vbInformation
End Sub


' M20_Slide の ZKubun は Private なので、ここから使えるように同じ判定を置く
Private Function ZKubunPub(ByVal code As String, ByVal nm As String) As String
    Dim t As String, i As Long, ch As String, digits As String
    Dim n As Long

    If IsScrapName(nm) Then ZKubunPub = "ス（スクラップ）": Exit Function
    If InStr(1, NormText(nm), NormText("支給品費")) > 0 Then Exit Function

    t = NormText(code)
    If Len(t) = 0 Then Exit Function
    If UCase$(Left$(t, 1)) <> "Z" Then Exit Function

    For i = 2 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit For
    Next i
    If digits = "" Then Exit Function

    n = CLng(digits)
    If n < 40 Then
        ZKubunPub = "積（共通仮設費の積上分）"
    ElseIf n < 45 Then
        ZKubunPub = "原（工事原価に加算）"
    Else
        ZKubunPub = "価（工事価格に加算）"
    End If
End Function
