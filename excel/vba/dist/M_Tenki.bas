Attribute VB_Name = "M_Tenki"
'==================================================================
' 01 総括表 用
'
' このファイル1つだけを標準モジュールに貼り付ければ動く。
' マクロは「定義シートを作成」「照合実行」「転記実行」。
'
' 元は excel/vba/src/ の M_Util.bas / M_Config.bas / M_Engine.bas / M_Main.bas を
' つなげたもの。直すときは src 側を直して build_vba.py を実行する。
'==================================================================
Option Explicit

'--- M_Util.bas の宣言 ----------------------------------------
'==================================================================
' M_Util - 文字列正規化・セル操作の共通ヘルパー
'
' 数量計算書は作成者や年度によって全角/半角、スペースの入り方が
' まちまちなので、比較する前に必ず Norm を通して表記ゆれを吸収する。
'==================================================================

' 照合方式
Public Enum MatchMode
    mmPartial = 0   ' 部分一致（既定）
    mmExact = 1     ' 完全一致
    mmDia = 2       ' 口径一致（数字だけ取り出して数値比較）
End Enum

'--- M_Config.bas の宣言 --------------------------------------
'==================================================================
' M_Config - 転記定義シートの作成と読み込み
'
' 「どこから拾って、どこへ入れるか」を VBA に直接書かず、
' ブック内の「転記定義」シートに持たせる。別の工事の数量計算書に
' 適用するときは、このシートを書き換えるだけでコード修正は不要。
'==================================================================

Public Const DEF_SHEET As String = "転記定義"
Public Const DEST_SHEET As String = "総括表（管工事）"
Public Const REPORT_SHEET As String = "照合結果"
Public Const FIRST_DATA_ROW As Long = 2

' 定義1行分
Public Type TDef
    No          As Long
    Enabled     As Boolean
    Kind        As String      ' "自動転記" / "手入力"
    ItemName    As String      ' 名称（総括表の表記）
    Spec        As String      ' 摘要（口径など）
    DestCell    As String      ' 転記先セル（例 D10）
    SrcFile     As String      ' 転記元ファイル名
    SrcSheet    As String      ' 転記元シート名
    Key1Col     As String
    Key1        As String
    Key1Mode    As MatchMode
    Key2Col     As String
    Key2        As String
    Key2Mode    As MatchMode
    RowOffset   As Long        ' 見つけた行からの相対行
    ValueCol    As String      ' 値のある列（直接指定する場合）
    HeaderRow   As Long        ' 見出し行（列を見出しから探す場合）
    HeaderKey   As String      ' 見出しキー
    Decimals    As Long        ' 丸め桁数（-1 で丸めない）
    Note        As String
End Type

Private Const HDR As String = "No|有効|区分|名称|摘要|転記先セル|元ファイル|元シート|" & _
    "キー1列|キー1|キー1方式|キー2列|キー2|キー2方式|行オフセット|値列|見出し行|見出しキー|丸め|備考"

'--- M_Engine.bas の宣言 --------------------------------------
'==================================================================
' M_Engine - 転記元セルの解決エンジン
'
' 定義シートの「キー」からセル番地を毎回探し直す。番地を直接
' 持たないので、行が増減しても、ファイル名が元の日本語名でも動く。
'==================================================================

Private mOpened As Collection      ' このマクロが開いたブック（最後に閉じる）
Private mFolder As String          ' 転記元フォルダ（一度選んだら記憶）

'--- M_Main.bas の宣言 ----------------------------------------
'==================================================================
' M_Main - 実行の入口
'
'   照合実行()       … 書き換えずに、総括表の値と計算値を突き合わせる
'   転記実行()       … 差異を確認したうえで総括表に書き込む
'   定義シートを作成() … M_Config にある（最初に一度だけ実行）
'
' 使い方は excel/vba/README.md を参照。
'==================================================================

Private Type TResult
    No       As Long
    Kind     As String
    ItemName As String
    Spec     As String
    DestCell As String
    OldVal   As Variant
    NewVal   As Variant
    Judge    As String     ' 一致 / 差異 / 転記 / エラー / 手入力 / 対象外
    Source   As String
    Message  As String
End Type

'==================================================================
' ここから M_Util.bas
'==================================================================
'------------------------------------------------------------------
' 表記ゆれを吸収した比較用文字列を返す
'   ・全角英数記号 → 半角
'   ・空白（半角/全角/タブ/改行）を除去
'   ・φ Φ ｆ などの径記号を除去
'   ・英字は大文字に統一
'------------------------------------------------------------------
Public Function Norm(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Or IsEmpty(s) Then Norm = "": Exit Function
    t = CStr(s)
    If Len(t) = 0 Then Norm = "": Exit Function

    t = StrConv(t, vbNarrow)                 ' 全角→半角
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, " ", "")                  ' 半角スペース
    t = Replace(t, ChrW(&H3000), "")         ' 全角スペース
    t = Replace(t, ChrW(&H3C6), "")          ' φ
    t = Replace(t, ChrW(&H3A6), "")          ' Φ
    t = Replace(t, "f", "")
    t = Replace(t, "F", "")
    Norm = UCase$(t)
End Function

'------------------------------------------------------------------
' 文字列から最初の数字の並びを取り出して数値で返す（見つからなければ -1）
'   "SP400A" → 400   "φ400 1種管" → 400   "300A" → 300
'------------------------------------------------------------------
Public Function ExtractNum(ByVal s As Variant) As Double
    Dim t As String, i As Long, ch As String, buf As String
    Dim started As Boolean

    ExtractNum = -1
    If IsError(s) Or IsEmpty(s) Then Exit Function
    t = StrConv(CStr(s), vbNarrow)

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            buf = buf & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i

    If Len(buf) > 0 Then ExtractNum = CDbl(buf)
End Function

'------------------------------------------------------------------
' 指定の照合方式でキーと対象を比較する
'------------------------------------------------------------------
Public Function KeyMatches(ByVal target As Variant, ByVal key As String, _
                           ByVal mode As MatchMode) As Boolean
    Dim nt As String, nk As String

    If Len(key) = 0 Then KeyMatches = False: Exit Function

    Select Case mode
        Case mmDia
            Dim a As Double, b As Double
            a = ExtractNum(target)
            b = ExtractNum(key)
            KeyMatches = (a >= 0 And b >= 0 And a = b)
        Case mmExact
            KeyMatches = (Norm(target) = Norm(key))
        Case Else
            nt = Norm(target)
            nk = Norm(key)
            KeyMatches = (Len(nt) > 0 And InStr(1, nt, nk, vbTextCompare) > 0)
    End Select
End Function

'------------------------------------------------------------------
' 列文字("A" / "AO") → 列番号。数字が渡された場合はそのまま数値化。
' 不正な値は 0 を返す。
'------------------------------------------------------------------
Public Function ColToNum(ByVal col As Variant) As Long
    Dim s As String
    s = Trim$(StrConv(CStr(col), vbNarrow))
    If Len(s) = 0 Then ColToNum = 0: Exit Function

    If IsNumeric(s) Then
        ColToNum = CLng(s)
    Else
        On Error Resume Next
        ColToNum = Range(s & "1").Column
        If Err.Number <> 0 Then Err.Clear: ColToNum = 0
        On Error GoTo 0
    End If
End Function

'------------------------------------------------------------------
' シート名をゆるく探す（全角/半角スペースの違いを無視）
' 見つからなければ Nothing
'------------------------------------------------------------------
Public Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet, target As String
    target = Norm(sheetName)

    For Each ws In wb.Worksheets
        If Norm(ws.Name) = target Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Set FindSheet = Nothing
End Function

'------------------------------------------------------------------
' セルが数値ならその値、それ以外は Empty を返す
'------------------------------------------------------------------
Public Function NumOrEmpty(ByVal c As Range) As Variant
    If IsError(c.Value) Then
        NumOrEmpty = Empty
    ElseIf IsNumeric(c.Value) And Not IsEmpty(c.Value) And CStr(c.Value) <> "" Then
        NumOrEmpty = CDbl(c.Value)
    Else
        NumOrEmpty = Empty
    End If
End Function

'==================================================================
' ここから M_Config.bas
'==================================================================
'------------------------------------------------------------------
' 定義シートを作成し、初期値（本工事の解析結果）を書き込む
' 既にある場合は確認のうえ作り直す
'------------------------------------------------------------------
Public Sub 定義シートを作成()
    Dim ws As Worksheet, hdrArr As Variant, i As Long
    Dim rows_ As Variant, r As Long

    Set ws = FindSheet(ThisWorkbook, DEF_SHEET)
    If Not ws Is Nothing Then
        If MsgBox("「" & DEF_SHEET & "」シートは既にあります。" & vbCrLf & _
                  "作り直すと、加えた変更は失われます。続けますか？", _
                  vbYesNo + vbExclamation, "確認") <> vbYes Then Exit Sub
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = DEF_SHEET

    hdrArr = Split(HDR, "|")
    For i = 0 To UBound(hdrArr)
        ws.Cells(1, i + 1).Value = hdrArr(i)
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(hdrArr) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
    End With

    rows_ = DefaultDefs()
    For r = 0 To UBound(rows_)
        For i = 0 To UBound(rows_(r))
            ws.Cells(r + FIRST_DATA_ROW, i + 1).Value = rows_(r)(i)
        Next i
    Next r

    ws.Rows(1).AutoFilter
    ws.Columns.AutoFit
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ws.Activate
    ActiveWindow.FreezePanes = True

    MsgBox "「" & DEF_SHEET & "」シートを作成しました。" & vbCrLf & vbCrLf & _
           "自動転記: " & CountKind(rows_, "自動転記") & " 件" & vbCrLf & _
           "手入力  : " & CountKind(rows_, "手入力") & " 件（照合のみ）", _
           vbInformation, "完了"
End Sub

Private Function CountKind(ByVal rows_ As Variant, ByVal kind As String) As Long
    Dim r As Long, n As Long
    For r = 0 To UBound(rows_)
        If rows_(r)(2) = kind Then n = n + 1
    Next r
    CountKind = n
End Function

'------------------------------------------------------------------
' 定義シートを読み込んで TDef の配列を返す
'------------------------------------------------------------------
Public Function LoadDefs(ByRef defs() As TDef) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long, n As Long

    Set ws = FindSheet(ThisWorkbook, DEF_SHEET)
    If ws Is Nothing Then
        MsgBox "「" & DEF_SHEET & "」シートがありません。" & vbCrLf & _
               "先に「定義シートを作成」を実行してください。", vbExclamation
        LoadDefs = 0
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row   ' 転記先セル列で判定
    If lastRow < FIRST_DATA_ROW Then LoadDefs = 0: Exit Function

    ReDim defs(0 To lastRow - FIRST_DATA_ROW)
    For r = FIRST_DATA_ROW To lastRow
        With defs(n)
            .No = Val(ws.Cells(r, 1).Value)
            .Enabled = (Norm(ws.Cells(r, 2).Value) <> "" And Norm(ws.Cells(r, 2).Value) <> "X")
            .Kind = Trim$(CStr(ws.Cells(r, 3).Value))
            .ItemName = Trim$(CStr(ws.Cells(r, 4).Value))
            .Spec = Trim$(CStr(ws.Cells(r, 5).Value))
            .DestCell = Trim$(CStr(ws.Cells(r, 6).Value))
            .SrcFile = Trim$(CStr(ws.Cells(r, 7).Value))
            .SrcSheet = Trim$(CStr(ws.Cells(r, 8).Value))
            .Key1Col = Trim$(CStr(ws.Cells(r, 9).Value))
            .Key1 = Trim$(CStr(ws.Cells(r, 10).Value))
            .Key1Mode = ParseMode(ws.Cells(r, 11).Value)
            .Key2Col = Trim$(CStr(ws.Cells(r, 12).Value))
            .Key2 = Trim$(CStr(ws.Cells(r, 13).Value))
            .Key2Mode = ParseMode(ws.Cells(r, 14).Value)
            .RowOffset = Val(ws.Cells(r, 15).Value)
            .ValueCol = Trim$(CStr(ws.Cells(r, 16).Value))
            .HeaderRow = Val(ws.Cells(r, 17).Value)
            .HeaderKey = Trim$(CStr(ws.Cells(r, 18).Value))
            If Len(Trim$(CStr(ws.Cells(r, 19).Value))) = 0 Then
                .Decimals = -1
            Else
                .Decimals = Val(ws.Cells(r, 19).Value)
            End If
            .Note = Trim$(CStr(ws.Cells(r, 20).Value))
        End With
        n = n + 1
    Next r

    ReDim Preserve defs(0 To n - 1)
    LoadDefs = n
End Function

Private Function ParseMode(ByVal v As Variant) As MatchMode
    Select Case Norm(v)
        Case "完全", "EXACT":  ParseMode = mmExact
        Case "口径", "DIA":    ParseMode = mmDia
        Case Else:             ParseMode = mmPartial
    End Select
End Function

'------------------------------------------------------------------
' 本工事（東白川特2高層配水池揚水管取替工事）の解析から得た初期定義
'   列順は HDR と同じ
'------------------------------------------------------------------
Private Function DefaultDefs() As Variant
    Dim d As Collection: Set d = New Collection
    Dim n As Long

    ' ---- 02 鋳鉄管製造 数量計算書 / 　印刷　 ----------------------
    ' 「据付延長 設計書 入力数値」行を口径列で引く（丸め済みの値が並んでいる）
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋳鉄管据付", "200", "D10", _
        "02_chutetsukan.xls", "　印刷　", "B", "据付延長 設計書 入力数値", "部分", _
        "", "", "", 0, "", 5, "200", 1, "口径見出しは行5(G=75〜O=450)")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋳鉄管据付", "300", "D11", _
        "02_chutetsukan.xls", "　印刷　", "B", "据付延長 設計書 入力数値", "部分", _
        "", "", "", 0, "", 5, "300", 1, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋳鉄管据付", "400", "D12", _
        "02_chutetsukan.xls", "　印刷　", "B", "据付延長 設計書 入力数値", "部分", _
        "", "", "", 0, "", 5, "400", 1, "")

    ' 継手類は品名(C列)×形質(D列)で行を決め、数量(F列)を読む
    n = n + 1: d.Add Array(n, "○", "自動転記", "GX継手（直部）", "200", "J32", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 直 管", "部分", "D", "200", "口径", 0, "F", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "GX継手（直部）", "400", "J34", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 直 管", "部分", "D", "400", "口径", 0, "F", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "GX継手（異形部）", "200", "J38", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 異形管接合材", "部分", "D", "200", "口径", 0, "F", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "GX継手（異形部）", "300", "J39", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 異形管接合材", "部分", "D", "300", "口径", 0, "F", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "GX継手（異形部）", "400", "J40", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 異形管接合材", "部分", "D", "400", "口径", 0, "F", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "GX継手（特殊押輪部）", "400", "J45", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 特殊押輪", "部分", "D", "400", "口径", 0, "F", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "挿口加工費（GX形）", "200", "D139", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 挿口リング", "部分", "D", "200", "口径", 0, "F", 0, "", 0, _
        "03切管表 W列とも一致するはず")
    n = n + 1: d.Add Array(n, "○", "自動転記", "挿口加工費（GX形）", "400", "D140", _
        "02_chutetsukan.xls", "　印刷　", "C", "GX形 挿口リング", "部分", "D", "400", "口径", 0, "F", 0, "", 0, _
        "03切管表 W列とも一致するはず")

    ' ---- 05-1 根拠 / 延長集計表  採用(伏越削除) --------------------
    ' G列=管種記号、K列=据付、L列=撤去（見出しは行138）
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管据付", "80", "D18", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "G", "SP80A", "完全", _
        "", "", "", 0, "K", 0, "", 1, "K列=据付 L列=撤去（見出し行138）")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管据付", "300", "D22", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "G", "SP300A", "完全", _
        "", "", "", 0, "K", 0, "", 1, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管据付", "400", "D23", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "G", "SP400A", "完全", _
        "", "", "", 0, "K", 0, "", 1, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管撤去", "250", "D44", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "G", "SP250A", "完全", _
        "", "", "", 0, "L", 0, "", 1, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管撤去", "300", "D45", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "G", "SP300A", "完全", _
        "", "", "", 0, "L", 0, "", 1, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管撤去", "400", "D46", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "G", "SP400A", "完全", _
        "", "", "", 0, "L", 0, "", 1, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "電送管据付", "80（FEP)", "J62", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "O", "布設延長合計", "部分", _
        "", "", "", 0, "P", 0, "", 1, "1079.4m×2条")
    n = n + 1: d.Add Array(n, "○", "自動転記", "電送管撤去", "82（VE)", "J64", _
        "05-1_dokou_konkyo.xlsx", "延長集計表  採用(伏越削除)", "O", "撤去延長合計", "部分", _
        "", "", "", 0, "P", 0, "", 1, "")

    ' ---- 03 切管表 / 各シートの「合 計」行の次行 -------------------
    ' U列=切断本数、W列=挿口リング本数
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋳鉄管切断 （新　管）", "200", "J86", _
        "03_kirikan.xlsx", "GX　200", "B", "合 計", "部分", "", "", "", 1, "U", 0, "", 0, _
        "「合 計」の次の行。U列=切断本数")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋳鉄管切断 （新　管）", "400", "J88", _
        "03_kirikan.xlsx", "GX　400 (4)", "B", "合 計", "部分", "", "", "", 1, "U", 0, "", 0, _
        "φ400は最終シート(4)の合計を使う。シートが増えたら要変更")

    ' ---- 08 鋼管工事数量 / 工事数量表 ------------------------------
    ' A列=工種（結合のため下方向に空白）、L列=規格、AO列=数量
    n = n + 1: d.Add Array(n, "○", "自動転記", "ステンレス鋼管 現場溶接", "80", "D60", _
        "08_koukan.xls", "工事数量表", "A", "電気溶接", "部分", "L", "80A", "部分", 0, "AO", 0, "", 0, _
        "A列は上方向に補完して判定")
    n = n + 1: d.Add Array(n, "○", "自動転記", "ステンレス鋼管 現場溶接", "300", "D61", _
        "08_koukan.xls", "工事数量表", "A", "電気溶接", "部分", "L", "300A", "部分", 0, "AO", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "ステンレス鋼管 現場溶接", "400", "D62", _
        "08_koukan.xls", "工事数量表", "A", "電気溶接", "部分", "L", "400A", "部分", 0, "AO", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管現場溶接", "300", "D64", _
        "08_koukan.xls", "工事数量表", "A", "閉塞蓋設置", "部分", "L", "300", "口径", 0, "AO", 0, "", 0, "閉塞蓋部")
    n = n + 1: d.Add Array(n, "○", "自動転記", "鋼管現場溶接", "400", "D65", _
        "08_koukan.xls", "工事数量表", "A", "閉塞蓋設置", "部分", "L", "400", "口径", 0, "AO", 0, "", 0, "閉塞蓋部")

    ' ---- 手入力項目（転記元が特定できず。照合時に現在値のみ表示）----
    n = n + 1: d.Add Array(n, "○", "手入力", "GX継手 取外し （異形部）", "400", "D109", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "図面から計数")
    n = n + 1: d.Add Array(n, "○", "手入力", "フランジ継手", "80", "J51", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "16K(絶縁)")
    n = n + 1: d.Add Array(n, "○", "手入力", "フランジ継手", "250", "J55", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "20K")
    n = n + 1: d.Add Array(n, "○", "手入力", "フランジ継手 取外し", "75", "J72", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "フランジ継手 取外し", "250", "J76", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "空気弁据付", "25 16k", "J145", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "補修弁2有 F短管有")
    n = n + 1: d.Add Array(n, "○", "手入力", "空気弁据付", "75　10k", "J149", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "補修弁2有 F短管有")
    n = n + 1: d.Add Array(n, "○", "手入力", "空気弁据付", "75　16k", "J151", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "補修弁2有 F短管有")
    n = n + 1: d.Add Array(n, "○", "手入力", "空気弁撤去", "25", "J165", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "空気弁撤去", "75", "J167", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "仕切弁据付 （車道部）", "200", "D170", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "仕切弁据付 （車道部）", "400", "D172", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "仕切弁据付 （歩道部）", "200", "D176", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")
    n = n + 1: d.Add Array(n, "○", "手入力", "仕切弁据付 （歩道部）", "400", "D178", "", "", "", "", "", "", "", "", 0, "", 0, "", 0, "")

    Dim arr() As Variant, i As Long
    ReDim arr(0 To d.Count - 1)
    For i = 1 To d.Count
        arr(i - 1) = d(i)
    Next i
    DefaultDefs = arr
End Function

'==================================================================
' ここから M_Engine.bas
'==================================================================
'==================================================================
' ブックの管理
'==================================================================
Private Sub InitSources()
    Set mOpened = New Collection
    mFolder = ""
End Sub

Private Sub CloseSources()
    Dim i As Long, wb As Workbook
    If mOpened Is Nothing Then Exit Sub
    For i = mOpened.Count To 1 Step -1
        On Error Resume Next
        Set wb = mOpened(i)
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        On Error GoTo 0
    Next i
    Set mOpened = Nothing
End Sub

'------------------------------------------------------------------
' 転記元ブックを取得する
'   1. 既に開いていればそれを使う
'   2. ThisWorkbook と同じフォルダを探す
'   3. 見つからなければ、先頭の番号（02 / 05-1 など）が一致する
'      ファイルを探す ＝ 元の日本語ファイル名のままでも拾える
'   4. それでも駄目ならフォルダを尋ねる（一度だけ）
'------------------------------------------------------------------
Public Function GetSourceBook(ByVal fileName As String, ByRef errMsg As String) As Workbook
    Dim wb As Workbook, path_ As String, actual As String

    errMsg = ""
    If Len(fileName) = 0 Then errMsg = "元ファイルが指定されていません": Exit Function

    ' 1. 開いているブックから
    For Each wb In Application.Workbooks
        If Norm(wb.Name) = Norm(fileName) Then
            Set GetSourceBook = wb
            Exit Function
        End If
    Next wb

    ' 2〜3. フォルダから探す
    actual = LocateFile(ThisWorkbook.path, fileName)
    If Len(actual) = 0 And Len(mFolder) > 0 Then
        actual = LocateFile(mFolder, fileName)
    End If

    ' 4. フォルダを尋ねる
    If Len(actual) = 0 Then
        If AskFolder(fileName) Then
            actual = LocateFile(mFolder, fileName)
        End If
    End If

    If Len(actual) = 0 Then
        errMsg = "ファイルが見つかりません: " & fileName
        Exit Function
    End If

    ' 既に開いていないか、実ファイル名で再確認
    For Each wb In Application.Workbooks
        If Norm(wb.Name) = Norm(Dir$(actual)) Then
            Set GetSourceBook = wb
            Exit Function
        End If
    Next wb

    On Error GoTo OpenFailed
    Set wb = Application.Workbooks.Open(fileName:=actual, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0

    mOpened.Add wb
    Set GetSourceBook = wb
    Exit Function

OpenFailed:
    errMsg = "ファイルを開けません: " & actual & " (" & Err.Description & ")"
End Function

'------------------------------------------------------------------
' フォルダ内から目的のファイルを探して、フルパスを返す
'------------------------------------------------------------------
Private Function LocateFile(ByVal folder_ As String, ByVal wantName As String) As String
    Dim p As String, token As String, f As String, ext As String

    LocateFile = ""
    If Len(folder_) = 0 Then Exit Function
    If Right$(folder_, 1) <> "\" Then folder_ = folder_ & "\"

    ' 完全一致
    p = folder_ & wantName
    If Len(Dir$(p)) > 0 Then LocateFile = p: Exit Function

    ' 先頭の番号が一致するファイル（例 "02_..." → "02　東白川　鋳鉄管製造..."）
    token = LeadingToken(wantName)
    If Len(token) = 0 Then Exit Function

    f = Dir$(folder_ & "*.xls*")
    Do While Len(f) > 0
        If LeadingToken(f) = token Then
            LocateFile = folder_ & f
            Exit Function
        End If
        f = Dir$
    Loop
End Function

'------------------------------------------------------------------
' ファイル名の先頭にある番号を取り出す（"05-1_x.xlsx" → "05-1"）
'------------------------------------------------------------------
Private Function LeadingToken(ByVal s As String) As String
    Dim i As Long, ch As String, buf As String
    s = StrConv(s, vbNarrow)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Then
            buf = buf & ch
        Else
            Exit For
        End If
    Next i
    Do While Right$(buf, 1) = "-"
        buf = Left$(buf, Len(buf) - 1)
    Loop
    LeadingToken = buf
End Function

Private Function AskFolder(ByVal fileName As String) As Boolean
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "「" & fileName & "」がある フォルダを選んでください"
    If fd.Show = -1 Then
        mFolder = fd.SelectedItems(1)
        AskFolder = True
    End If
End Function

'==================================================================
' 値の解決
'==================================================================
Public Function ResolveValue(ByRef d As TDef, ByRef outVal As Variant, _
                             ByRef errMsg As String) As Boolean
    Dim wb As Workbook, ws As Worksheet
    Dim r As Long, c As Long, v As Variant

    outVal = Empty
    errMsg = ""

    Set wb = GetSourceBook(d.SrcFile, errMsg)
    If wb Is Nothing Then Exit Function

    Set ws = FindSheet(wb, d.SrcSheet)
    If ws Is Nothing Then
        errMsg = "シートがありません: [" & wb.Name & "] " & d.SrcSheet
        Exit Function
    End If

    r = FindTargetRow(ws, d)
    If r <= 0 Then
        errMsg = "行が見つかりません: キー1=" & d.Key1 & _
                 IIf(Len(d.Key2) > 0, " / キー2=" & d.Key2, "")
        Exit Function
    End If

    c = FindValueCol(ws, d)
    If c <= 0 Then
        errMsg = "列が特定できません: 値列=" & d.ValueCol & " 見出しキー=" & d.HeaderKey
        Exit Function
    End If

    v = NumOrEmpty(ws.Cells(r, c))
    If IsEmpty(v) Then
        errMsg = "数値ではありません: " & ws.Name & "!" & ws.Cells(r, c).Address(False, False)
        Exit Function
    End If

    If d.Decimals >= 0 Then v = Application.WorksheetFunction.Round(v, d.Decimals)

    outVal = v
    ResolveValue = True
End Function

'------------------------------------------------------------------
' キーに合う行を探す
'   キー列が結合や省略で空になっている場合に備え、上方向に値を補完する
'------------------------------------------------------------------
Private Function FindTargetRow(ByVal ws As Worksheet, ByRef d As TDef) As Long
    Dim k1c As Long, k2c As Long, lastRow As Long, r As Long
    Dim cur1 As String, v As Variant
    Dim ok1 As Boolean, ok2 As Boolean

    k1c = ColToNum(d.Key1Col)
    k2c = ColToNum(d.Key2Col)
    lastRow = LastUsedRow(ws)

    If k1c = 0 And k2c = 0 Then FindTargetRow = 0: Exit Function

    For r = 1 To lastRow
        ' キー1（上方向に補完）
        If k1c > 0 Then
            v = ws.Cells(r, k1c).Value
            If Len(Norm(v)) > 0 Then cur1 = CStr(v)
        End If

        If k1c = 0 Or Len(d.Key1) = 0 Then
            ok1 = True
        Else
            ok1 = KeyMatches(cur1, d.Key1, d.Key1Mode)
        End If

        If ok1 Then
            If k2c = 0 Or Len(d.Key2) = 0 Then
                ok2 = True
            Else
                ok2 = KeyMatches(ws.Cells(r, k2c).Value, d.Key2, d.Key2Mode)
            End If

            If ok2 Then
                FindTargetRow = r + d.RowOffset
                Exit Function
            End If
        End If
    Next r

    FindTargetRow = 0
End Function

'------------------------------------------------------------------
' 値のある列を決める
'   値列が指定されていればそれ。無ければ見出しから探す。
'   指定の見出し行に無い場合は、上から40行を走査して探し直す。
'------------------------------------------------------------------
Private Function FindValueCol(ByVal ws As Worksheet, ByRef d As TDef) As Long
    Dim c As Long, r As Long, lastCol As Long, mode As MatchMode

    If Len(d.ValueCol) > 0 Then
        FindValueCol = ColToNum(d.ValueCol)
        Exit Function
    End If

    If Len(d.HeaderKey) = 0 Then FindValueCol = 0: Exit Function

    ' 見出しが数字（口径）なら数値比較、そうでなければ部分一致
    If IsNumeric(StrConv(d.HeaderKey, vbNarrow)) Then
        mode = mmDia
    Else
        mode = mmPartial
    End If

    lastCol = LastUsedCol(ws)

    If d.HeaderRow > 0 Then
        For c = 1 To lastCol
            If KeyMatches(ws.Cells(d.HeaderRow, c).Value, d.HeaderKey, mode) Then
                FindValueCol = c
                Exit Function
            End If
        Next c
    End If

    ' 見つからないので行がずれたとみなして探し直す
    For r = 1 To Application.Min(40, LastUsedRow(ws))
        For c = 1 To lastCol
            If KeyMatches(ws.Cells(r, c).Value, d.HeaderKey, mode) Then
                FindValueCol = c
                Exit Function
            End If
        Next c
    Next r

    FindValueCol = 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If Err.Number <> 0 Or LastUsedRow < 1 Then LastUsedRow = 1
    On Error GoTo 0
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If Err.Number <> 0 Or LastUsedCol < 1 Then LastUsedCol = 1
    On Error GoTo 0
End Function

'==================================================================
' ここから M_Main.bas
'==================================================================
'==================================================================
' 照合（書き換えない）
'==================================================================
Public Sub 照合実行()
    Dim results() As TResult, n As Long
    n = RunAll(results, False)
    If n < 0 Then Exit Sub
    WriteReport results, n, "照合"
    MsgBox Summary(results, n, False), vbInformation, "照合が終わりました"
End Sub

'==================================================================
' 転記（確認のうえ書き換える）
'==================================================================
Public Sub 転記実行()
    Dim results() As TResult, n As Long, diffs As Long, i As Long
    Dim ws As Worksheet, ans As VbMsgBoxResult

    ' まず書き換えずに差分を出す
    n = RunAll(results, False)
    If n < 0 Then Exit Sub

    For i = 0 To n - 1
        If results(i).Judge = "差異" Then diffs = diffs + 1
    Next i

    If diffs = 0 Then
        WriteReport results, n, "照合"
        MsgBox "総括表はすべて計算値と一致しています。" & vbCrLf & _
               "書き換える必要はありませんでした。", vbInformation, "転記は不要です"
        Exit Sub
    End If

    ans = MsgBox(diffs & " 件の差異が見つかりました。" & vbCrLf & vbCrLf & _
                 "総括表に書き込みますか？" & vbCrLf & _
                 "（書き込む前に、総括表のバックアップシートを作ります）", _
                 vbYesNo + vbQuestion, "転記の確認")
    If ans <> vbYes Then
        WriteReport results, n, "照合"
        MsgBox "書き込みは行いませんでした。差分は「" & REPORT_SHEET & "」を確認してください。", _
               vbInformation
        Exit Sub
    End If

    MakeBackup

    ' 実際に書き込む
    n = RunAll(results, True)
    If n < 0 Then Exit Sub
    WriteReport results, n, "転記"
    MsgBox Summary(results, n, True), vbInformation, "転記が終わりました"
End Sub

'==================================================================
' 本体：定義に沿って値を解決し、必要なら書き込む
'   戻り値 … 件数。異常時は -1
'==================================================================
Private Function RunAll(ByRef results() As TResult, ByVal doWrite As Boolean) As Long
    Dim defs() As TDef, cnt As Long, i As Long
    Dim ws As Worksheet, cell_ As Range
    Dim v As Variant, msg As String
    Dim nmActual As String, spActual As String
    Dim scr As Boolean, calc As XlCalculation

    RunAll = -1

    cnt = LoadDefs(defs)
    If cnt = 0 Then Exit Function

    Set ws = FindSheet(ThisWorkbook, DEST_SHEET)
    If ws Is Nothing Then
        MsgBox "転記先シート「" & DEST_SHEET & "」が見つかりません。" & vbCrLf & _
               "M_Config の DEST_SHEET を実際のシート名に合わせてください。", vbExclamation
        Exit Function
    End If

    scr = Application.ScreenUpdating
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    InitSources
    ReDim results(0 To cnt - 1)

    On Error GoTo Cleanup
    For i = 0 To cnt - 1
        With results(i)
            .No = defs(i).No
            .Kind = defs(i).Kind
            .ItemName = defs(i).ItemName
            .Spec = defs(i).Spec
            .DestCell = defs(i).DestCell
            .Source = defs(i).SrcFile & IIf(Len(defs(i).SrcSheet) > 0, " / " & defs(i).SrcSheet, "")
            .OldVal = Empty
            .NewVal = Empty
            .Message = ""
        End With

        ' 転記先セルの妥当性
        Set cell_ = Nothing
        On Error Resume Next
        Set cell_ = ws.Range(results(i).DestCell)
        On Error GoTo Cleanup
        If cell_ Is Nothing Then
            results(i).Judge = "エラー"
            results(i).Message = "転記先セルが不正です: " & results(i).DestCell
            GoTo NextItem
        End If

        results(i).OldVal = NumOrEmpty(cell_)

        ' 総括表側の名称・摘要が定義とずれていないか確認
        DestLabels ws, cell_, nmActual, spActual
        If Len(defs(i).ItemName) > 0 Then
            If Not KeyMatches(nmActual, defs(i).ItemName, mmPartial) Then
                results(i).Message = "※総括表の名称が定義と違います（実際: " & nmActual & "）"
            End If
        End If

        If Not defs(i).Enabled Then
            results(i).Judge = "対象外"
            GoTo NextItem
        End If

        If defs(i).Kind <> "自動転記" Then
            results(i).Judge = "手入力"
            results(i).Message = Trim$(results(i).Message & " " & defs(i).Note)
            GoTo NextItem
        End If

        ' 転記元から値を取る
        If Not ResolveValue(defs(i), v, msg) Then
            results(i).Judge = "エラー"
            results(i).Message = Trim$(results(i).Message & " " & msg)
            GoTo NextItem
        End If

        results(i).NewVal = v

        If Not IsEmpty(results(i).OldVal) Then
            If Abs(CDbl(results(i).OldVal) - CDbl(v)) < 0.00000001 Then
                results(i).Judge = "一致"
                GoTo NextItem
            End If
        End If

        If doWrite Then
            cell_.Value = v
            results(i).Judge = "転記"
        Else
            results(i).Judge = "差異"
        End If

NextItem:
    Next i

    RunAll = cnt

Cleanup:
    If Err.Number <> 0 Then
        MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
               Err.Number & ": " & Err.Description, vbCritical
        RunAll = -1
    End If
    CloseSources
    Application.DisplayAlerts = True
    Application.Calculation = calc
    Application.ScreenUpdating = scr
    Application.CalculateFull
End Function

'------------------------------------------------------------------
' 転記先セルの行から、総括表上の名称・摘要を読む
'   D列 → 名称B 摘要C ／ J列 → 名称F 摘要G
'   名称は結合や省略で空のことがあるので上方向に補完する
'------------------------------------------------------------------
Private Sub DestLabels(ByVal ws As Worksheet, ByVal cell_ As Range, _
                       ByRef outName As String, ByRef outSpec As String)
    Dim nameCol As Long, specCol As Long, r As Long, i As Long

    outName = "": outSpec = ""
    r = cell_.Row

    Select Case cell_.Column
        Case 4:  nameCol = 2: specCol = 3     ' D列
        Case 10: nameCol = 6: specCol = 7     ' J列
        Case Else: Exit Sub
    End Select

    outSpec = Trim$(CStr(ws.Cells(r, specCol).Value))

    For i = r To Application.Max(1, r - 40) Step -1
        If Len(Norm(ws.Cells(i, nameCol).Value)) > 0 Then
            outName = " " & Trim$(CStr(ws.Cells(i, nameCol).Value))
            outName = Trim$(Replace(Replace(outName, vbLf, " "), vbCr, " "))
            Exit For
        End If
    Next i
End Sub

'------------------------------------------------------------------
' 総括表のバックアップシートを作る
'------------------------------------------------------------------
Private Sub MakeBackup()
    Dim ws As Worksheet, bk As Worksheet, nm As String

    Set ws = FindSheet(ThisWorkbook, DEST_SHEET)
    If ws Is Nothing Then Exit Sub

    nm = "BK_" & Format$(Now, "mmdd_hhnn")
    If Len(nm) > 31 Then nm = Left$(nm, 31)

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set bk = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    bk.Name = nm
    bk.Visible = xlSheetVisible
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------
' 結果シートを書き出す
'------------------------------------------------------------------
Private Sub WriteReport(ByRef results() As TResult, ByVal n As Long, ByVal mode As String)
    Dim ws As Worksheet, i As Long, r As Long
    Dim hdr As Variant, c As Long

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(REPORT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = REPORT_SHEET

    ws.Range("A1").Value = mode & "結果　" & Format$(Now, "yyyy/mm/dd hh:nn")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12

    hdr = Array("No", "区分", "名称", "摘要", "転記先", "総括表の値", "計算値", "差", "判定", "転記元", "メッセージ")
    For c = 0 To UBound(hdr)
        ws.Cells(3, c + 1).Value = hdr(c)
    Next c
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(hdr) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
    End With

    r = 4
    For i = 0 To n - 1
        With results(i)
            ws.Cells(r, 1).Value = .No
            ws.Cells(r, 2).Value = .Kind
            ws.Cells(r, 3).Value = .ItemName
            ws.Cells(r, 4).Value = .Spec
            ws.Cells(r, 5).Value = .DestCell
            If Not IsEmpty(.OldVal) Then ws.Cells(r, 6).Value = .OldVal
            If Not IsEmpty(.NewVal) Then ws.Cells(r, 7).Value = .NewVal
            If Not IsEmpty(.OldVal) And Not IsEmpty(.NewVal) Then
                ws.Cells(r, 8).Value = CDbl(.NewVal) - CDbl(.OldVal)
            End If
            ws.Cells(r, 9).Value = .Judge
            ws.Cells(r, 10).Value = .Source
            ws.Cells(r, 11).Value = .Message

            Select Case .Judge
                Case "差異":   ws.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
                Case "転記":   ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
                Case "エラー": ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
                Case "手入力": ws.Cells(r, 9).Interior.Color = RGB(226, 226, 226)
            End Select
        End With
        r = r + 1
    Next i

    ws.Rows(3).AutoFilter
    ws.Columns.AutoFit
    If ws.Columns(11).ColumnWidth > 60 Then ws.Columns(11).ColumnWidth = 60
    ws.Activate
    ws.Range("A4").Select
End Sub

Private Function Summary(ByRef results() As TResult, ByVal n As Long, _
                         ByVal wrote As Boolean) As String
    Dim i As Long
    Dim nOK As Long, nDiff As Long, nWr As Long, nErr As Long, nMan As Long

    For i = 0 To n - 1
        Select Case results(i).Judge
            Case "一致":   nOK = nOK + 1
            Case "差異":   nDiff = nDiff + 1
            Case "転記":   nWr = nWr + 1
            Case "エラー": nErr = nErr + 1
            Case "手入力": nMan = nMan + 1
        End Select
    Next i

    Summary = "一致  : " & nOK & " 件" & vbCrLf
    If wrote Then
        Summary = Summary & "転記  : " & nWr & " 件" & vbCrLf
    Else
        Summary = Summary & "差異  : " & nDiff & " 件" & vbCrLf
    End If
    Summary = Summary & "手入力: " & nMan & " 件（照合対象外）" & vbCrLf & _
                        "エラー: " & nErr & " 件" & vbCrLf & vbCrLf & _
                        "詳しくは「" & REPORT_SHEET & "」シートを見てください。"
    If nErr > 0 Then
        Summary = Summary & vbCrLf & vbCrLf & _
                  "※エラーの多くは、転記元ファイルが見つからないか" & vbCrLf & _
                  "　シート名・キーが変わったことが原因です。"
    End If
End Function
