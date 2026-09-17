'==============================================================
' M00_Main  ―  インフレスライド設計書 作成マクロ（本体）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M00_Main" にしてください。
'
' 使い方
'   1) M00_Main / M10_CsvIO / M20_Report の3モジュールを貼り付ける
'   2) 「設定シート作成」を実行 → 「設定」シートに工事情報とCSVの列番号を入力
'   3) 「インフレスライド設計書作成」を実行
'
' 前提
'   ・積算システムから「変動前（旧単価）」「変動後（新単価）」の2本のCSVを出力
'   ・両CSVの明細を 工種｜種別｜細別｜規格｜単位 で突合し、差額を算出する
'==============================================================
Option Explicit

Public Const SH_CONFIG As String = "設定"


'==============================================================
' メイン処理
'==============================================================
Public Sub インフレスライド設計書作成()
    Dim cfg As Object
    Dim pathOld As String, pathNew As String
    Dim dOld As Object, dNew As Object

    On Error GoTo ErrHandler

    If Not SheetExists(SH_CONFIG) Then
        MsgBox "「" & SH_CONFIG & "」シートがありません。" & vbCrLf & _
               "先に［設定シート作成］を実行してください。", vbExclamation
        Exit Sub
    End If

    Set cfg = GetConfig()

    pathOld = AskCsvPath(CStr(CfgVal(cfg, "変動前CSVパス", "")), "変動前（旧単価）のCSVを選択してください")
    If pathOld = "" Then Exit Sub
    pathNew = AskCsvPath(CStr(CfgVal(cfg, "変動後CSVパス", "")), "変動後（新単価）のCSVを選択してください")
    If pathNew = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "CSVを読み込んでいます…"

    Set dOld = LoadMeisai(pathOld, cfg)
    Set dNew = LoadMeisai(pathNew, cfg)

    If dOld.Count = 0 And dNew.Count = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "明細を1行も読み取れませんでした。" & vbCrLf & _
               "「設定」シートの［見出し行数］と［列マッピング］を確認してください。", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "設計書を作成しています…"
    BuildReport cfg, dOld, dNew

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "作成しました。" & vbCrLf & vbCrLf & _
           "変動前 明細：" & dOld.Count & " 行" & vbCrLf & _
           "変動後 明細：" & dNew.Count & " 行", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "内容：" & Err.Description & vbCrLf & _
           "番号：" & Err.Number, vbCritical
End Sub


'==============================================================
' CSVパスの決定（設定が空ならダイアログ）
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
' CSV → 明細Dictionary
'   key   : 工種|種別|細別|規格|単位（正規化済み）
'   value : Variant配列（F_xxx 定数でアクセス）
'   同一キーが複数行あるときは数量・金額を合算する
'==============================================================
Private Function LoadMeisai(ByVal filePath As String, ByVal cfg As Object) As Object
    Dim rows As Collection, d As Object, ex As Object
    Dim arr() As String
    Dim i As Long, headRows As Long
    Dim koshu As String, shubetsu As String, saibetsu As String
    Dim kikaku As String, tani As String, tekiyo As String
    Dim suryo As Double, tanka As Double, kingaku As Double
    Dim k As String
    Dim rec As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set ex = ExcludeWords(cfg)

    Set rows = ParseCsvText( _
                    ReadTextFile(filePath, CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))), _
                    CStr(CfgVal(cfg, "区切り文字", ",")))

    headRows = CLng(CfgVal(cfg, "見出し行数", 1))

    For i = headRows + 1 To rows.Count
        arr = rows(i)

        koshu    = Trim$(ColVal(arr, CLng(CfgVal(cfg, "工種列", 1))))
        shubetsu = Trim$(ColVal(arr, CLng(CfgVal(cfg, "種別列", 2))))
        saibetsu = Trim$(ColVal(arr, CLng(CfgVal(cfg, "細別列", 3))))
        kikaku   = Trim$(ColVal(arr, CLng(CfgVal(cfg, "規格列", 4))))
        tani     = Trim$(ColVal(arr, CLng(CfgVal(cfg, "単位列", 5))))
        tekiyo   = Trim$(ColVal(arr, CLng(CfgVal(cfg, "摘要列", 9))))

        suryo   = ToNum(ColVal(arr, CLng(CfgVal(cfg, "数量列", 6))))
        tanka   = ToNum(ColVal(arr, CLng(CfgVal(cfg, "単価列", 7))))
        kingaku = ToNum(ColVal(arr, CLng(CfgVal(cfg, "金額列", 8))))

        ' 名称が全て空の行は読み飛ばす
        If NormText(koshu & shubetsu & saibetsu & kikaku) = "" Then GoTo ContinueLoop

        ' 合計・小計などの集計行は二重計上になるため除外
        If ex.Exists(NormText(koshu)) Then GoTo ContinueLoop
        If ex.Exists(NormText(shubetsu)) Then GoTo ContinueLoop
        If ex.Exists(NormText(saibetsu)) Then GoTo ContinueLoop

        ' 金額列が無い／空の場合は 数量×単価 で補う
        If kingaku = 0 And suryo <> 0 And tanka <> 0 Then kingaku = suryo * tanka

        ' 数量・単価・金額がすべて0の行（見出し行）は除外
        If suryo = 0 And tanka = 0 And kingaku = 0 Then GoTo ContinueLoop

        k = NormText(koshu) & "|" & NormText(shubetsu) & "|" & NormText(saibetsu) & _
            "|" & NormText(kikaku) & "|" & NormText(tani)

        If d.Exists(k) Then
            rec = d(k)
            rec(F_SURYO) = CDbl(rec(F_SURYO)) + suryo
            rec(F_KINGAKU) = CDbl(rec(F_KINGAKU)) + kingaku
            If CDbl(rec(F_SURYO)) <> 0 Then
                rec(F_TANKA) = CDbl(rec(F_KINGAKU)) / CDbl(rec(F_SURYO))
            End If
            d(k) = rec
        Else
            ReDim rec(0 To F_COUNT - 1)
            rec(F_KOSHU) = koshu
            rec(F_SHUBETSU) = shubetsu
            rec(F_SAIBETSU) = saibetsu
            rec(F_KIKAKU) = kikaku
            rec(F_TANI) = tani
            rec(F_SURYO) = suryo
            rec(F_TANKA) = tanka
            rec(F_KINGAKU) = kingaku
            rec(F_TEKIYO) = tekiyo
            d.Add k, rec
        End If

ContinueLoop:
    Next i

    Set LoadMeisai = d
End Function


'==============================================================
' 設定シートの読み取り
'==============================================================
Public Function GetConfig() As Object
    Dim ws As Worksheet, d As Object
    Dim r As Long, lastR As Long
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                   ' vbTextCompare
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)

    lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
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


'--------------------------------------------------------------
' 除外語リスト（設定シートの「除外語」行以下、B列を読む）
'--------------------------------------------------------------
Private Function ExcludeWords(ByVal cfg As Object) As Object
    Dim ws As Worksheet, d As Object
    Dim r As Long, lastR As Long, startR As Long
    Dim w As String

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)

    lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    startR = 0

    For r = 1 To lastR
        If NormText(CStr(ws.Cells(r, 1).Value)) = "除外語" Then
            startR = r
            Exit For
        End If
    Next r

    If startR = 0 Then
        ' 既定の除外語
        d.Add "合計", True: d.Add "小計", True: d.Add "総計", True: d.Add "計", True
        Set ExcludeWords = d
        Exit Function
    End If

    For r = startR To ws.Cells(ws.rows.Count, 2).End(xlUp).Row
        w = NormText(CStr(ws.Cells(r, 2).Value))
        If w <> "" Then
            If Not d.Exists(w) Then d.Add w, True
        End If
    Next r

    Set ExcludeWords = d
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

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Set FreshSheet = ws
End Function
