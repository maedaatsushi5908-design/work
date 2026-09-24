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
Public Const R_COUNT  As Long = 9

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
Public Const SC_MK_SCRAP As Long = 18   ' R  スクラップ〇  ☆自動判定（修正可）
Public Const SC_A_ALL    As Long = 19   ' S  スライド前・全体      =K*I
Public Const SC_A_DONE   As Long = 20   ' T  スライド前・出来形    =L*I
Public Const SC_A_REST   As Long = 21   ' U  スライド前・残工事    =M*I
Public Const SC_B_ALL    As Long = 22   ' V  スライド後・全体      =K*J
Public Const SC_B_REST   As Long = 23   ' W  スライド後・残工事    =M*J
Public Const SC_N_ALL    As Long = 24   ' X  新工種抜き・全体      =IF(Q="〇",0,S)
Public Const SC_N_DONE   As Long = 25   ' Y  新工種抜き・出来形    =IF(Q="〇",0,T)
Public Const SC_TEKIYO   As Long = 26   ' Z  摘要
Public Const SC_SP_O     As Long = 27   ' AA 処分費単価 スライド前 ★手入力
Public Const SC_SP_N     As Long = 28   ' AB 処分費単価 スライド後 ★手入力
Public Const SC_SA_ALL   As Long = 29   ' AC 処分費額 前・全体
Public Const SC_SA_DONE  As Long = 30   ' AD 処分費額 前・出来形
Public Const SC_SA_REST  As Long = 31   ' AE 処分費額 前・残工事
Public Const SC_SB_ALL   As Long = 32   ' AF 処分費額 後・全体
Public Const SC_SB_REST  As Long = 33   ' AG 処分費額 後・残工事
Public Const SC_SN_ALL   As Long = 34   ' AH 処分費額 新抜き・全体
Public Const SC_SN_DONE  As Long = 35   ' AI 処分費額 新抜き・出来形
Public Const SC_LAST     As Long = 35

Public Const SH_SLIDE    As String = "スライド計算表"
Public Const FIRST_ROW   As Long = 8

Public Const CLR_INPUT   As Long = 65535          ' 黄色（手入力）
Public Const CLR_HEAD    As Long = 15849925       ' 薄い青
Public Const CLR_TOTAL   As Long = 49407          ' オレンジ

' ---- 他モジュールが参照する行 ----
Public gDirectFirst As Long     ' 直接工事費部の先頭行
Public gDirectLast  As Long     ' 直接工事費部の最終行
Public gZFirst      As Long     ' 共通仮設費 積上げ部の先頭行（0＝なし）
Public gZLast       As Long     ' 　　　　　　　　　　最終行
Public gZItems      As Object   ' Dictionary 名称 → "先頭行|最終行"
Public gDetailLast  As Long     ' 明細の最終行
Public gRowKakaku2  As Long     ' ㉑'工事価格（スクラップ込み）の行
Public gRowZei      As Long     ' ㉓消費税相当額の行
Public gRowKoujihi  As Long     ' ㉔工事費の行
Public gLastRow     As Long

' スクラップ行の判定用
Private mScrapLevel As Long

' 経費計算部の作業用
Private mKWs As Worksheet
Private mKCfg As Object

'--- M40_Chosho の宣言 ---
Public Const SH_CHOSHO As String = "スライド調書（様式4-2号）"

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
           "　・O列 処分費〇／P列 管材費〇／Q列 新工種〇" & vbCrLf & _
           "　・AA列/AB列 処分費単価（合算単価から処分費分だけ）" & vbCrLf & _
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
'   usedLabel に採用した系列（"当初" / "変更"）を返す
'==============================================================
Private Function LoadEstima(ByVal filePath As String, ByVal cfg As Object, _
                            ByRef usedLabel As String) As Collection
    Dim rows As Collection, recs As Collection
    Dim arr() As String
    Dim i As Long, rx1000 As Long
    Dim firstIsToshu As Boolean
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

    '--- X1000（本工事費）の位置を探す ---
    rx1000 = 0
    For i = 1 To rows.Count
        arr = rows(i)
        If InStr(1, ColVal(arr, CSV_CODE), "X1000", vbTextCompare) > 0 Then
            rx1000 = i
            Exit For
        End If
    Next i

    '--- ①が当初か変更かを判定（X1000の次の行の設計金額②が0なら ①＝当初）---
    firstIsToshu = True
    If rx1000 > 0 And rx1000 < rows.Count Then
        arr = rows(rx1000 + 1)
        If ToNum(ColVal(arr, CSV_KIN2)) <> 0 Then firstIsToshu = False
    End If

    '--- 使用する系列を決める ---
    seriesName = CStr(CfgVal(cfg, "使用系列", "自動"))
    Select Case seriesName
        Case "当初": useFirst = firstIsToshu
        Case "変更": useFirst = Not firstIsToshu
        Case Else:   useFirst = True        ' 自動＝①（常に最新側）
    End Select

    If useFirst Then
        usedLabel = IIf(firstIsToshu, "当初", "変更")
    Else
        usedLabel = IIf(firstIsToshu, "変更", "当初")
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
        Else
            rec(R_TANKA) = ToNum(ColVal(arr, CSV_TANKA2))
            rec(R_SURYO) = ToNum(ColVal(arr, CSV_SURYO2))
            rec(R_KIN) = ToNum(ColVal(arr, CSV_KIN2))
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
    PutItem ws, r, "使用系列", "自動", "自動／当初／変更。通常は自動（①側＝最新を採用）": r = r + 2

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
    Dim cnt As Long, firstZ As Long, skippedG As Long
    Dim zCnt As Long, zRowArr() As Long, zLvlArr() As Long
    Dim curItem As String, curStart As Long

    n = recOld.Count
    ReDim newTanka(1 To n)
    warnText = MatchNewPrices(recOld, recNew, newTanka)

    Set ws = FreshSheet(SH_SLIDE)
    Set gZItems = CreateObject("Scripting.Dictionary")
    WriteHeader ws, cfg

    ReDim rowArr(1 To n)
    ReDim lvlArr(1 To n)
    cnt = 0
    r = FIRST_ROW - 1
    firstZ = 0
    gDirectFirst = FIRST_ROW
    mScrapLevel = -1

    '--- 直接工事費部（X1000配下）---
    For i = 1 To n
        rec = recOld(i)
        lvl = CStr(rec(R_LEVEL))

        If lvl = "G" Then
            skippedG = skippedG + 1
        ElseIf lvl = "Z" Or lvl = "YZ" Then
            firstZ = i
            Exit For
        Else
            r = r + 1
            lvlNum = LevelNum(lvl)
            WriteRow ws, r, rec, newTanka(i), lvlNum, lvl
            cnt = cnt + 1
            rowArr(cnt) = r
            lvlArr(cnt) = lvlNum
        End If
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
        curItem = "": curStart = 0
        mScrapLevel = -1

        For i = firstZ To n
            rec = recOld(i)
            lvl = CStr(rec(R_LEVEL))
            If lvl = "G" Then GoTo NextZ

            r = r + 1
            lvlNum = LevelNum(lvl)
            WriteRow ws, r, rec, newTanka(i), lvlNum, lvl

            If lvl = "Z" Then
                If curItem <> "" Then gZItems(curItem) = curStart & "|" & (r - 1)
                curItem = CStr(rec(R_NAME))
                curStart = r
            End If

            zCnt = zCnt + 1
            zRowArr(zCnt) = r
            zLvlArr(zCnt) = lvlNum
NextZ:
        Next i

        If curItem <> "" Then gZItems(curItem) = curStart & "|" & r
        WriteSumFormulas ws, zRowArr, zLvlArr, zCnt
        gZLast = r
    End If

    gDetailLast = r
    FinishDetail ws, r

    '--- 経費計算部 ---
    r = BuildKeihiSection(ws, cfg, r + 2)
    gLastRow = r
    FinishSheet ws, r

    If skippedG > 0 Then
        warnText = warnText & IIf(warnText = "", "", vbCrLf) & _
                   "Gコード行 " & skippedG & " 行は計算表に出していません。"
    End If

    BuildSlideSheet = warnText
End Function


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


Public Function LevelNum(ByVal lvl As String) As Long
    Select Case lvl
        Case "費目": LevelNum = 0
        Case "L1":   LevelNum = 1
        Case "L2":   LevelNum = 2
        Case "L3":   LevelNum = 3
        Case "L4":   LevelNum = 4
        Case "Z":    LevelNum = 1
        Case "YZ":   LevelNum = 2
        Case Else:   LevelNum = 9
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
    ws.Cells(5, SC_MK_SHIN).Value = "新工種"
    ws.Cells(5, SC_MK_SCRAP).Value = "スクラップ"
    ws.Cells(3, SC_A_ALL).Value = "スライド前（旧単価）"
    ws.Cells(5, SC_A_ALL).Value = "全体"
    ws.Cells(5, SC_A_DONE).Value = "出来形"
    ws.Cells(5, SC_A_REST).Value = "残工事"
    ws.Cells(3, SC_B_ALL).Value = "スライド後（新単価）"
    ws.Cells(5, SC_B_ALL).Value = "全体"
    ws.Cells(5, SC_B_REST).Value = "残工事"
    ws.Cells(3, SC_N_ALL).Value = "新工種抜き"
    ws.Cells(5, SC_N_ALL).Value = "全体"
    ws.Cells(5, SC_N_DONE).Value = "出来形"
    ws.Cells(3, SC_TEKIYO).Value = "摘要"
    ws.Cells(3, SC_SP_O).Value = "処分費単価（手入力）"
    ws.Cells(5, SC_SP_O).Value = "スライド前"
    ws.Cells(5, SC_SP_N).Value = "スライド後"
    ws.Cells(3, SC_SA_ALL).Value = "処分費額"
    ws.Cells(5, SC_SA_ALL).Value = "前・全体"
    ws.Cells(5, SC_SA_DONE).Value = "前・出来形"
    ws.Cells(5, SC_SA_REST).Value = "前・残工事"
    ws.Cells(5, SC_SB_ALL).Value = "後・全体"
    ws.Cells(5, SC_SB_REST).Value = "後・残工事"
    ws.Cells(5, SC_SN_ALL).Value = "新抜き・全体"
    ws.Cells(5, SC_SN_DONE).Value = "新抜き・出来形"

    With ws.Range(ws.Cells(3, 1), ws.Cells(6, SC_LAST))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(3, SC_MK_SHOBU), ws.Cells(6, SC_MK_SCRAP)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(3, SC_SP_O), ws.Cells(6, SC_SP_N)).Interior.Color = CLR_INPUT
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
        ' スクラップは㉒として工事価格の後に足すので、①直接工事費からは外す
        If mScrapLevel >= 0 Or IsScrapName(CStr(rec(R_NAME))) Then
            ws.Cells(r, SC_MK_SCRAP).Value = "〇"
        End If
    Else
        ws.Cells(r, SC_TANI).Value = "式"
        If mScrapLevel >= 0 And lvlNum <= mScrapLevel Then mScrapLevel = -1
        If IsScrapName(CStr(rec(R_NAME))) Then mScrapLevel = lvlNum
    End If
End Sub


Public Sub WriteDetailFormulas(ByVal ws As Worksheet, ByVal r As Long)
    Dim aI As String, aJ As String, aK As String, aL As String, aM As String
    Dim aO As String, aQ As String, aX As String, aY As String
    Dim spO As String, spN As String

    aI = ws.Cells(r, SC_TANKA_O).Address(False, False)
    aJ = ws.Cells(r, SC_TANKA_N).Address(False, False)
    aK = ws.Cells(r, SC_Q_ALL).Address(False, False)
    aL = ws.Cells(r, SC_Q_DONE).Address(False, False)
    aM = ws.Cells(r, SC_Q_REST).Address(False, False)
    aO = ws.Cells(r, SC_MK_SHOBU).Address(False, False)
    aQ = ws.Cells(r, SC_MK_SHIN).Address(False, False)
    aX = ws.Cells(r, SC_SP_O).Address(False, False)
    aY = ws.Cells(r, SC_SP_N).Address(False, False)

    ' 処分費単価：未入力ならその行の設計単価を使う
    spO = "IF(" & aX & "=""""," & aI & "," & aX & ")"
    spN = "IF(" & aY & "="""",IF(" & aX & "=""""," & aJ & "," & aX & ")," & aY & ")"

    ws.Cells(r, SC_Q_REST).Formula = "=" & aK & "-" & aL
    ws.Cells(r, SC_A_ALL).Formula = "=" & aK & "*" & aI
    ws.Cells(r, SC_A_DONE).Formula = "=" & aL & "*" & aI
    ws.Cells(r, SC_A_REST).Formula = "=" & aM & "*" & aI
    ws.Cells(r, SC_B_ALL).Formula = "=" & aK & "*" & aJ
    ws.Cells(r, SC_B_REST).Formula = "=" & aM & "*" & aJ
    ' 新工種抜き：新工種〇の行を0にする
    ws.Cells(r, SC_N_ALL).Formula = "=IF(" & aQ & "=""〇"",0," & _
        ws.Cells(r, SC_A_ALL).Address(False, False) & ")"
    ws.Cells(r, SC_N_DONE).Formula = "=IF(" & aQ & "=""〇"",0," & _
        ws.Cells(r, SC_A_DONE).Address(False, False) & ")"

    ws.Cells(r, SC_SA_ALL).Formula = "=IF(" & aO & "=""〇""," & aK & "*" & spO & ",0)"
    ws.Cells(r, SC_SA_DONE).Formula = "=IF(" & aO & "=""〇""," & aL & "*" & spO & ",0)"
    ws.Cells(r, SC_SA_REST).Formula = "=IF(" & aO & "=""〇""," & aM & "*" & spO & ",0)"
    ws.Cells(r, SC_SB_ALL).Formula = "=IF(" & aO & "=""〇""," & aK & "*" & spN & ",0)"
    ws.Cells(r, SC_SB_REST).Formula = "=IF(" & aO & "=""〇""," & aM & "*" & spN & ",0)"
    ws.Cells(r, SC_SN_ALL).Formula = "=IF(" & aQ & "=""〇"",0," & _
        ws.Cells(r, SC_SA_ALL).Address(False, False) & ")"
    ws.Cells(r, SC_SN_DONE).Formula = "=IF(" & aQ & "=""〇"",0," & _
        ws.Cells(r, SC_SA_DONE).Address(False, False) & ")"
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

    cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST, SC_N_ALL, SC_N_DONE)

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
    Dim i As Long, c As Long, r As Long
    Dim kojo As String, zei As String
    Dim unpan As String, gijutsu As String, scrap As String
    Dim rDCHOKU As Long, rKANZAI As Long, rSHOBUN As Long, rTAIGAI As Long
    Dim rUNPAN As Long, rTSUMI2 As Long, rGIJUTSU As Long
    Dim rKTAISHO As Long, rKRITSU As Long, rKBUN As Long, rKKEI As Long, rJUN As Long
    Dim rGTAISHO As Long, rGRITSU As Long, rGKEI As Long, rGENKA As Long
    Dim rITAISHO As Long, rIRITSU As Long, rIBUN As Long, rHOSHO As Long
    Dim rIKEI0 As Long, rKAKAKU0 As Long, rHASU As Long, rIKEI As Long, rKAKAKU As Long
    Dim rSCRAP As Long

    Set mKWs = ws
    Set mKCfg = cfg

    cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST, SC_N_ALL, SC_N_DONE)
    shoCols = Array(SC_SA_ALL, SC_SA_DONE, SC_SA_REST, SC_SB_ALL, SC_SB_REST, SC_SN_ALL, SC_SN_DONE)
    ' 0＝自前の率（手入力可）／-1＝率なし（素材のみ）／1以上＝その系列番号の率を参照
    rateSrc = Array(0, 1, -1, 0, 4, 0, 6)

    kojo = NumStr(CfgVal(cfg, "処分費控除率", 0.03))
    zei = NumStr(CfgVal(cfg, "消費税率", 0.1))
    unpan = FindZItem("運搬費")
    gijutsu = FindZItem("技術管理費")
    scrap = FindZItem("スクラップ")

    '--- 行番号を先に確定させる ---
    r = startRow
    ws.Cells(r, SC_HIMOKU).Value = "【経費計算】　※黄色セルは手入力。経費率は積算システムの値があれば上書きしてください"
    With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_MK_SCRAP))
        .Merge
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
    End With
    r = r + 1

    rDCHOKU = r:  KLabel ws, r, "①直接工事費":                                          r = r + 1
    rKANZAI = r:  KLabel ws, r, "①'水道工事における管材費":                             r = r + 1
    rSHOBUN = r:  KLabel ws, r, "②直接工事費内の処分費等":                              r = r + 1
    rTAIGAI = r:  KLabel ws, r, "③②のうち率計算の対象外費　※②－(①－①'/2)×3%":       r = r + 1
    rUNPAN = r:   KLabel ws, r, "④共通仮設費積上分－運搬費":                            r = r + 1
    rTSUMI2 = r:  KLabel ws, r, "④'共通仮設費積上分－その他":                           r = r + 1
    rGIJUTSU = r: KLabel ws, r, "⑤共通仮設費積上分－技術管理費":                        r = r + 1
    rKTAISHO = r: KLabel ws, r, "⑥共通仮設費対象額　※①－③－①'/2":                    r = r + 1
    rKRITSU = r:  KLabel ws, r, "⑦共通仮設費率":                                        r = r + 1
    rKBUN = r:    KLabel ws, r, "⑦'共通仮設費率分　※⑥×⑦":                            r = r + 1
    rKKEI = r:    KLabel ws, r, "⑧共通仮設費【合計】　※④＋④'＋⑤＋⑦'":                r = r + 1
    rJUN = r:     KLabel ws, r, "⑨純工事費　※①＋⑧":                                   r = r + 1
    rGTAISHO = r: KLabel ws, r, "⑩現場管理費対象額　※⑨－③－①'/2":                    r = r + 1
    rGRITSU = r:  KLabel ws, r, "⑪現場管理費率":                                        r = r + 1
    rGKEI = r:    KLabel ws, r, "⑫現場管理費【合計】　※⑩×⑪":                          r = r + 1
    rGENKA = r:   KLabel ws, r, "⑬工事原価　※⑨＋⑫":                                   r = r + 1
    rITAISHO = r: KLabel ws, r, "⑭一般管理費対象額　※⑬－③":                           r = r + 1
    rIRITSU = r:  KLabel ws, r, "⑮一般管理費率":                                        r = r + 1
    rIBUN = r:    KLabel ws, r, "⑮'一般管理費率分　※⑭×⑮":                            r = r + 1
    rHOSHO = r:   KLabel ws, r, "⑯契約保証費　※当初設計時額で固定":                     r = r + 1
    rIKEI0 = r:   KLabel ws, r, "⑰一般管理費【合計】（端数整理前）　※⑮'＋⑯":          r = r + 1
    rKAKAKU0 = r: KLabel ws, r, "⑱工事価格（端数整理前）　※⑬＋⑰":                      r = r + 1
    rHASU = r:    KLabel ws, r, "⑲端数整理":                                            r = r + 1
    rIKEI = r:    KLabel ws, r, "⑳一般管理費【合計】（端数整理後）　※⑰－⑲":           r = r + 1
    rKAKAKU = r:  KLabel ws, r, "㉑工事価格　※⑬＋⑳":                                   r = r + 1
    rSCRAP = r:   KLabel ws, r, "㉒スクラップ　※スクラップ〇の行から集計":               r = r + 1
    gRowKakaku2 = r: KLabel ws, r, "㉑'工事価格（スクラップ込み）　※㉑＋㉒":             r = r + 1
    gRowZei = r:     KLabel ws, r, "㉓消費税相当額　※(㉑＋㉒)×消費税率":                r = r + 1
    gRowKoujihi = r: KLabel ws, r, "㉔工事費　※㉑＋㉒＋㉓":                              r = r + 1

    '--- 系列ごとに数式を入れる ---
    For i = 0 To 6
        c = CLng(cols(i))

        ws.Cells(rDCHOKU, c).Formula = "=SUMPRODUCT((" & KR(SC_TANKA_O, gDirectFirst, gDirectLast) & _
            "<>"""")*(" & KR(SC_MK_SCRAP, gDirectFirst, gDirectLast) & "<>""〇"")*(" & _
            KR(c, gDirectFirst, gDirectLast) & "))"

        ws.Cells(rKANZAI, c).Formula = "=SUMPRODUCT((" & KR(SC_MK_KANZA, gDirectFirst, gDirectLast) & _
            "=""〇"")*(" & KR(SC_MK_SCRAP, gDirectFirst, gDirectLast) & "<>""〇"")*(" & _
            KR(c, gDirectFirst, gDirectLast) & "))"

        ws.Cells(rSHOBUN, c).Formula = "=SUMPRODUCT((" & KR(SC_TANKA_O, gDirectFirst, gDirectLast) & _
            "<>"""")*(" & KR(SC_MK_SCRAP, gDirectFirst, gDirectLast) & "<>""〇"")*(" & _
            KR(CLng(shoCols(i)), gDirectFirst, gDirectLast) & "))"

        ws.Cells(rTAIGAI, c).Formula = "=MAX(0," & KC(rSHOBUN, c) & "-ROUNDDOWN((" & _
            KC(rDCHOKU, c) & "-" & KC(rKANZAI, c) & "/2)*" & kojo & ",0))"

        ws.Cells(rUNPAN, c).Formula = ZItemSum(unpan, c)
        ws.Cells(rGIJUTSU, c).Formula = ZItemSum(gijutsu, c)

        If gZFirst > 0 Then
            ws.Cells(rTSUMI2, c).Formula = "=SUMPRODUCT((" & KR(SC_TANKA_O, gZFirst, gZLast) & _
                "<>"""")*(" & KR(SC_MK_SCRAP, gZFirst, gZLast) & "<>""〇"")*(" & _
                KR(c, gZFirst, gZLast) & "))-" & KC(rUNPAN, c) & "-" & KC(rGIJUTSU, c)
        Else
            ws.Cells(rTSUMI2, c).Value = 0
        End If

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
        ws.Cells(rKKEI, c).Formula = "=" & KC(rUNPAN, c) & "+" & KC(rTSUMI2, c) & _
                                     "+" & KC(rGIJUTSU, c) & "+" & KC(rKBUN, c)
        ws.Cells(rJUN, c).Formula = "=" & KC(rDCHOKU, c) & "+" & KC(rKKEI, c)

        ws.Cells(rGTAISHO, c).Formula = "=" & KC(rJUN, c) & "-" & KC(rTAIGAI, c) & _
                                        "-" & KC(rKANZAI, c) & "/2"
        KRate ws, rGRITSU, c, CLng(rateSrc(i)), cols, RateGenba(KC(rGTAISHO, c))
        ws.Cells(rGKEI, c).Formula = "=ROUNDDOWN(" & KC(rGTAISHO, c) & "*" & KC(rGRITSU, c) & ",-3)"
        ws.Cells(rGENKA, c).Formula = "=" & KC(rJUN, c) & "+" & KC(rGKEI, c)

        ws.Cells(rITAISHO, c).Formula = "=" & KC(rGENKA, c) & "-" & KC(rTAIGAI, c)
        KRate ws, rIRITSU, c, CLng(rateSrc(i)), cols, RateIppan(KC(rITAISHO, c))
        ws.Cells(rIBUN, c).Formula = "=" & KC(rITAISHO, c) & "*" & KC(rIRITSU, c)
        ws.Cells(rHOSHO, c).Value = CDbl(CfgVal(cfg, "契約保証費", 0))
        ws.Cells(rHOSHO, c).Interior.Color = CLR_INPUT
        ws.Cells(rIKEI0, c).Formula = "=" & KC(rIBUN, c) & "+" & KC(rHOSHO, c)
        ws.Cells(rKAKAKU0, c).Formula = "=" & KC(rGENKA, c) & "+" & KC(rIKEI0, c)
        ws.Cells(rHASU, c).Formula = "=" & KC(rKAKAKU0, c) & "-ROUNDDOWN(" & KC(rKAKAKU0, c) & ",-3)"
        ws.Cells(rIKEI, c).Formula = "=" & KC(rIKEI0, c) & "-" & KC(rHASU, c)
        ws.Cells(rKAKAKU, c).Formula = "=" & KC(rGENKA, c) & "+" & KC(rIKEI, c)

        ws.Cells(rSCRAP, c).Formula = "=SUMPRODUCT((" & KR(SC_TANKA_O, gDirectFirst, gDetailLast) & _
            "<>"""")*(" & KR(SC_MK_SCRAP, gDirectFirst, gDetailLast) & "=""〇"")*(" & _
            KR(c, gDirectFirst, gDetailLast) & "))"

        ws.Cells(gRowKakaku2, c).Formula = "=" & KC(rKAKAKU, c) & "+" & KC(rSCRAP, c)
        ws.Cells(gRowZei, c).Formula = "=(" & KC(rKAKAKU, c) & "+" & KC(rSCRAP, c) & ")*" & zei
        ws.Cells(gRowKoujihi, c).Formula = "=" & KC(rKAKAKU, c) & "+" & KC(rSCRAP, c) & _
                                           "+" & KC(gRowZei, c)
NextSeries:
    Next i

    '--- 体裁 ---
    With ws.Range(ws.Cells(startRow + 1, SC_A_ALL), ws.Cells(gRowKoujihi, SC_N_DONE))
        .NumberFormatLocal = "#,##0"
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(rKRITSU, SC_A_ALL), ws.Cells(rKRITSU, SC_N_DONE)).NumberFormatLocal = "0.0000"
    ws.Range(ws.Cells(rGRITSU, SC_A_ALL), ws.Cells(rGRITSU, SC_N_DONE)).NumberFormatLocal = "0.0000"
    ws.Range(ws.Cells(rIRITSU, SC_A_ALL), ws.Cells(rIRITSU, SC_N_DONE)).NumberFormatLocal = "0.0000"
    KHiLite ws, rJUN
    KHiLite ws, rGENKA
    KHiLite ws, rKAKAKU
    KHiLite ws, gRowKakaku2
    KHiLite ws, gRowKoujihi

    If unpan = "" Then KWarn ws, rUNPAN, "運搬費の行が見つかりませんでした"
    If gijutsu = "" Then KWarn ws, rGIJUTSU, "技術管理費の行が見つかりませんでした"

    BuildKeihiSection = gRowKoujihi
End Function


'--------------------------------------------------------------
' 経費計算部の補助
'--------------------------------------------------------------
Private Sub KLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal s As String)
    ws.Cells(r, SC_HIMOKU).Value = s
    With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_MK_SCRAP))
        .Merge
        .HorizontalAlignment = xlLeft
        .Borders.LineStyle = xlContinuous
    End With
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
    With ws.Range(ws.Cells(r, SC_A_ALL), ws.Cells(r, SC_N_DONE))
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


' gZItems から名称を含む項目の行範囲を探す（"r1|r2" を返す）
Private Function FindZItem(ByVal keyword As String) As String
    Dim k As Variant
    If gZItems Is Nothing Then Exit Function
    For Each k In gZItems.Keys
        If InStr(1, NormText(CStr(k)), NormText(keyword)) > 0 Then
            FindZItem = gZItems(k)
            Exit Function
        End If
    Next k
End Function


Private Function ZItemSum(ByVal rangeSpec As String, ByVal c As Long) As String
    Dim p() As String, r1 As Long, r2 As Long

    If rangeSpec = "" Then
        ZItemSum = "=0"
        Exit Function
    End If

    p = Split(rangeSpec, "|")
    r1 = CLng(p(0)): r2 = CLng(p(1))

    ZItemSum = "=SUMPRODUCT((" & KR(SC_TANKA_O, r1, r2) & "<>"""")*(" & _
               KR(SC_MK_SCRAP, r1, r2) & "<>""〇"")*(" & KR(c, r1, r2) & "))"
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
    ws.Range(ws.Cells(FIRST_ROW, SC_SP_O), ws.Cells(lastRow, SC_SN_DONE)).NumberFormatLocal = "#,##0"

    ws.Range(ws.Cells(FIRST_ROW, SC_Q_DONE), ws.Cells(lastRow, SC_Q_DONE)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_MK_SHOBU), ws.Cells(lastRow, SC_MK_SCRAP)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_SP_O), ws.Cells(lastRow, SC_SP_N)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_MK_SHOBU), ws.Cells(lastRow, SC_MK_SCRAP)).HorizontalAlignment = xlCenter

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
            ws.Range(ws.Cells(r, SC_MK_SHOBU), ws.Cells(r, SC_MK_SCRAP)).Interior.ColorIndex = xlColorIndexNone
            ws.Range(ws.Cells(r, SC_Q_DONE), ws.Cells(r, SC_Q_DONE)).Interior.ColorIndex = xlColorIndexNone
            ws.Range(ws.Cells(r, SC_SP_O), ws.Cells(r, SC_SP_N)).Interior.ColorIndex = xlColorIndexNone
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
    ws.Range(ws.Columns(SC_MK_SHOBU), ws.Columns(SC_MK_SCRAP)).ColumnWidth = 6
    ws.Range(ws.Columns(SC_A_ALL), ws.Columns(SC_N_DONE)).ColumnWidth = 13
    ws.Columns(SC_TEKIYO).ColumnWidth = 24
    ws.Range(ws.Columns(SC_SP_O), ws.Columns(SC_SN_DONE)).ColumnWidth = 11

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

    Set sl = ThisWorkbook.Worksheets(SH_SLIDE)
    Set ws = FreshSheet(SH_CHOSHO)
    contract = CDbl(CfgVal(cfg, "請負代金額", 0))

    ' スライド計算表へのリンク（S=前全体 T=前出来形 V=後全体 W=後残工事 X=新抜き全体 Y=新抜き出来形）
    fKoujihi = SlideRef(sl, gRowKoujihi, SC_A_ALL)
    fKakaku = SlideRef(sl, gRowKakaku2, SC_A_ALL)
    fDekidaka = SlideRef(sl, gRowKakaku2, SC_A_DONE)
    fZanNew = SlideRef(sl, gRowKakaku2, SC_B_REST)
    fZei = SlideRef(sl, gRowZei, SC_A_ALL)
    fShinNuki = SlideRef(sl, gRowKakaku2, SC_N_ALL)
    fShinNukiD = SlideRef(sl, gRowKakaku2, SC_N_DONE)

    ws.Range("G2").Value = "様式4-2号"
    ws.Range("B3").Value = "スライド調書"
    ws.Range("B3").Font.Size = 14
    ws.Range("B3").Font.Bold = True
    ws.Range("E3").Formula = "=IF(E16>F16,""減額スライド"",IF(E16<F16,""増額スライド"",""""))"
    ws.Range("B4").Value = CfgVal(cfg, "工事名", "")

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
    If contract > 0 Then
        ws.Range("I8").Formula = "=" & Format$(contract, "0.##########") & "/C8*100"
    Else
        ws.Range("I8").Value = 100
    End If
    ws.Range("I8").Interior.Color = CLR_INPUT

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
    ws.Range("C14").Formula = "=ROUNDDOWN(C8*I8/100,0)"

    ws.Range("B15").Value = "　請負工事価格"
    ws.Range("C15").Value = "④"
    ws.Range("D15").Value = "⑥=⑤×（③/①）"
    ws.Range("E15").Value = "P1'=④-⑥"
    ws.Range("F15").Value = "P2'=⑧×（③/①）"
    ws.Range("G15").Formula = "=IF(E16>F16,""S'=P2'-P1'+（P1''×1/100）"",""S'=P2'-P1'-（P1''×1/100）"")"

    ws.Range("B16").Value = "　（請負代金額（税抜））"
    ws.Range("C16").Formula = "=ROUNDDOWN(C14*10/11,0)"
    ws.Range("D16").Formula = "=ROUNDDOWN(D10*C14/C8,0)"
    ws.Range("E16").Formula = "=C16-D16"
    ws.Range("F16").Formula = "=ROUNDDOWN(F10*C14/C8,0)"
    ' 増額スライドは1%を控除、減額スライドは1%を戻す
    ws.Range("G16").Formula = "=IF(E16>F16,F16-E16+(E20*1/100),F16-E16-(E20*1/100))"

    ws.Range("B18").Value = "　消費税相当額"
    ws.Range("C18").Formula = "=C14-C16"

    ws.Range("B19").Value = "　請負工事価格" & vbLf & "（新工種抜き）"
    ws.Range("C19").Value = "⑨" & vbLf & "（請負率考慮）"
    ws.Range("D19").Value = "⑩" & vbLf & "（請負率考慮）"
    ws.Range("E19").Value = "P1''=⑨-⑩"
    ws.Range("B20").Value = "　（請負代金額（税抜））"
    ws.Range("C20").Formula = "=ROUNDDOWN(" & fShinNuki & "*C14/C8,0)"
    ws.Range("D20").Formula = "=ROUNDDOWN(" & fShinNukiD & "*C14/C8,0)"
    ws.Range("E20").Formula = "=C20-D20"

    ws.Range("B22").Value = "※ P1''（新工種を抜いた請負ベースの残工事）が受注者負担1%の母数です"
    ws.Range("B23").Value = "※ 金額はすべて「スライド計算表」の経費計算部からリンクしています"
    ws.Range("B22:B23").Font.Color = RGB(120, 120, 120)

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
        SetConfigValue "工事名", "○○地内　道路改良工事（テスト）"
        SetConfigValue "工事場所", "神戸市○○区○○町地内"
        SetConfigValue "工期（自）", "令和7年4月1日"
        SetConfigValue "工期（至）", "令和8年3月20日"
        SetConfigValue "請負代金額", 120000000
        SetConfigValue "基準日", "令和7年10月1日"
        SetConfigValue "発注者", "神戸市"
        SetConfigValue "受注者", "株式会社○○建設"
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
' サンプル（エスティマ「スライド用csv出力」形式・12列・見出しなし）
'   isNew = False … 当初の単価適用日
'   isNew = True  … 基準日の単価適用日（労務比率の高い工種ほど上昇）
'--------------------------------------------------------------
Private Function SampleCsv(ByVal isNew As Boolean) As String
    mSb = ""

    Head "G0100", "○○地内　道路改良工事"
    Head "X1000", "本工事費"

    Head "Y10001", "道路改良"
    Head "Y20001", "土工"
    Head "Y30001", "掘削工"
    Head "Y40001", "機械掘削"
    Det "床掘り", "m3", "土砂", "オープンカット", "", 1250, IIf(isNew, 336, 320)
    Head "Y40002", "埋戻工"
    Det "埋戻し", "m3", "流用土", "人力併用", "", 880, IIf(isNew, 452, 410)
    Det "土材料", "m3", "再生クラッシャラン0~40", "", "裏込め部", 40, IIf(isNew, 1800, 1800)
    Head "Y40003", "残土処理工"
    Det "土砂等運搬", "m3", "DT10t", "L=5km", "片道9.5km", 620, IIf(isNew, 1613, 1583)
    Det "残土等処分", "m3", "布施畑環境センター", "", "", 620, IIf(isNew, 3272, 3272)

    Head "Y30002", "擁壁工"
    Head "Y40011", "場所打擁壁工"
    Det "コンクリート", "m3", "24-12-25(20)(高炉)", "", "躯体", 145, IIf(isNew, 35690, 35400)
    Det "型枠", "m2", "一般型枠", "", "", 980, IIf(isNew, 8308, 8101)
    Det "鉄筋工", "t", "SD345", "D16～D25", "", 12.4, IIf(isNew, 135500, 117500)

    Head "Y20002", "舗装工"
    Head "Y30011", "路盤工"
    Head "Y40021", "下層路盤"
    Det "下層路盤", "m2", "クラッシャラン", "t=200", "", 3100, IIf(isNew, 1420, 1350)
    Head "Y40022", "上層路盤"
    Det "上層路盤", "m2", "粒度調整砕石", "t=150", "", 3100, IIf(isNew, 1680, 1590)
    Head "Y30012", "表層工"
    Head "Y40031", "表層"
    Det "表層", "m2", "密粒度AS", "t=50", "", 3100, IIf(isNew, 2240, 2080)

    Head "Y10002", "交通管理"
    Head "Y20011", "交通管理工"
    Head "Y30021", "交通誘導警備員"
    Det "交通誘導警備員B", "人日", "", "", "", 352, IIf(isNew, 12080, 11870)

    ' 諸経費（共通仮設費の積上げ分）
    Head "Z0001", "運搬費"
    Head "YZ0001", "重機分解組立輸送"
    Det "重機分解組立輸送", "回", "分解組立+輸送(往復)", "", "", 1, IIf(isNew, 831400, 828600)
    Head "Z0010", "準備費"
    Head "YZ0011", "木根等処分費"
    Det "高木伐採・根株撤去", "本", "幹周90cm～110cm", "", "集材含む", 2, IIf(isNew, 77750, 75485)
    Det "生木処分費", "t", "枝葉", "", "", 5.09, IIf(isNew, 16000, 16000)
    Head "Z0020", "技術管理費"
    Head "YZ0021", "土質試験費"
    Det "土の一軸圧縮試験", "試料", "", "", "2供試体/試料", 3, IIf(isNew, 10400, 10400)

    ' スクラップ（数量がマイナス。㉒として工事価格の後に足す）
    Head "Z0030", "スクラップ"
    Head "YZ0031", "ｽｸﾗｯﾌﾟ"
    Det "ｽｸﾗｯﾌﾟ", "t", "", "", "", -1.93, IIf(isNew, 23500, 9500)

    SampleCsv = mSb
End Function


' 階層行（コードあり・単価/数量なし）
Private Sub Head(ByVal code As String, ByVal nm As String)
    Row12 code, nm, "式", "", "", "", 0, 0
End Sub


' 明細行（コードなし）
Private Sub Det(ByVal nm As String, ByVal tani As String, ByVal k1 As String, _
                ByVal k2 As String, ByVal tekiyo As String, _
                ByVal suryo As Double, ByVal tanka As Double)
    Row12 "", nm, tani, k1, k2, tekiyo, suryo, tanka
End Sub


Private Sub Row12(ByVal code As String, ByVal nm As String, ByVal tani As String, _
                  ByVal k1 As String, ByVal k2 As String, ByVal tekiyo As String, _
                  ByVal suryo As Double, ByVal tanka As Double)
    mSb = mSb & Q(code) & "," & Q(nm) & "," & Q(tani) & "," & Q(k1) & "," & Q(k2) & "," & Q(tekiyo) & _
          "," & Format$(tanka, "0") & ",0," & _
          Format$(suryo, "0.00") & ",0," & _
          Format$(suryo * tanka, "0") & ",0" & vbCrLf
End Sub


Private Function Q(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Then
        Q = """" & Replace(s, """", """""") & """"
    Else
        Q = s
    End If
End Function
