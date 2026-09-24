'==============================================================
' M00_Main  ―  インフレスライド計算表 作成マクロ（本体）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M00_Main" にしてください。
'
' 使い方
'   1) M00_Main / M10_CsvIO / M20_Slide / M30_Shokeihi を貼り付ける
'   2) ［設定シート作成］→「設定」シートに工事情報を入力
'   3) ［スライド計算表作成］→ 旧単価CSV・新単価CSVを順に選ぶ
'   4) できたシートの黄色セル（出来形数量・処分費〇・処分費単価）を手入力する
'
' 前提
'   ・エスティマの 機能→CSV連携→抽出指示ファイル「スライド用csv出力」で出力したCSV
'   ・単価適用日を変えて2回出力する（旧単価で1本、新単価で1本）
'==============================================================
Option Explicit

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
