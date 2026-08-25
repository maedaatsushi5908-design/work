Attribute VB_Name = "M_Config"
'==================================================================
' M_Config - 転記定義シートの作成と読み込み
'
' 「どこから拾って、どこへ入れるか」を VBA に直接書かず、
' ブック内の「転記定義」シートに持たせる。別の工事の数量計算書に
' 適用するときは、このシートを書き換えるだけでコード修正は不要。
'==================================================================
Option Explicit

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
        "", "", "", 0, "", 5, "200", 1, "口径見出しは行5(G=75～O=450)")
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
