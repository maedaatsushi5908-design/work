'==============================================================
' M50_Tokushu  ―  特殊集計区分一覧表CSVから処分費・管材費を自動入力
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M50_Tokushu" に。
'
' 積算システムの「特殊集計区分一覧表」CSVを読み、スライド計算表の
'   O 列（処分費〇）  ← 特殊集計区分（処分）の行
'   P 列（管材費〇）  ← 特殊集計区分（水道）の行
'   AA列（処分費額）  ← その行の金額（全数量分）
' を自動で入れます。手入力の〇と金額が要らなくなります。
'
' マクロは2つです。
'   ・特殊集計区分CSV構造チェック … 金額を出さずに列の並びだけ書き出す（送付用）
'   ・特殊集計区分CSV取込         … 実際にスライド計算表へ入れる
'
' 列の並びは中身から自動で読み取りますが、うまくいかないときは
' 設定シートの「特殊集計CSV列指定」に
'   区分=14,コード=1,名称=2,単位=3,規格1=4,規格2=5,金額=11
' のように書けば、その指定が優先されます（列はA,B,…の記号でも可）。
'
' 突合は コード｜名称｜規格｜単位 で行います。同じキーが複数あるときは
' 上から順に1行ずつ対応させます（CSVと計算表の並びが同じ前提）。
'==============================================================
Option Explicit

' 単位らしさの判定に使う語彙（縦棒区切り）
Private Const TK_TANI_LIST As String = "式|個|本|箇所|か所|ヶ所|カ所|枚|組|台|人|日|月|回|基|面|条|袋|缶|丁|冊|" & _
    "m|m2|m3|㎡|㎥|ｍ|ｍ2|ｍ3|km|cm|mm|t|kg|g|L|l|ℓ|ha|a|人日|台日|日当|空m3|区間|社|戸|世帯|箇|ケ|個所"

Public Const SH_TK_RESULT As String = "_特殊集計取込結果"
Public Const SH_TK_CHECK  As String = "_特殊集計CSV構造"


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
                                    "特殊集計区分一覧表のCSVを選択してください")
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

    Set lay = TkDetect(rows, TkSpecOf(cfg), 0)
    maxC = CLng(lay("列数"))
    kcol = TkCol(lay, "区分")
    head0 = CLng(lay("見出し行"))
    carry = CBool(lay("区分は見出し行"))

    Set ws = FreshSheet(SH_TK_CHECK)
    ws.Range("A1").Value = "特殊集計区分CSV 構造チェック（金額・数量・単価は出していません）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "ファイル：" & Mid$(CStr(v), InStrRev(CStr(v), Application.PathSeparator) + 1)
    ws.Range("A3").Value = "読み取った並び：" & TkLayText(lay)

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
' ② 取込（スライド計算表のO列・P列・AA列に入れる）
'==============================================================
Public Sub 特殊集計区分CSV取込()
    Dim ws As Worksheet, rs As Worksheet
    Dim cfg As Object, lay As Object, map As Object
    Dim rows As Collection, q As Collection, rest As Collection
    Dim arr() As String, e As Variant
    Dim v As Variant, kAny As Variant
    Dim i As Long, r As Long, rr As Long, lastRow As Long
    Dim maxC As Long, kcol As Long, head0 As Long
    Dim charSetName As String
    Dim wSho As String, wKan As String
    Dim useKin As Boolean, doClear As Boolean
    Dim carry As Boolean, isLabel As Boolean
    Dim curKubun As String, kubun As String
    Dim kind As String, k As String, nm As String
    Dim nSho As Long, nKan As Long, nKin As Long, nSkip As Long, warnQ As Long
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
    charSetName = CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))
    wSho = TkWordsShobun(cfg)
    wKan = TkWordsKanzai(cfg)
    useKin = (NormText(CStr(CfgVal(cfg, "特殊集計CSVから処分費額も入れる", "する"))) = NormText("する"))

    '--- CSVを選ぶ ---
    v = CStr(CfgVal(cfg, "特殊集計区分CSVパス", ""))
    If CStr(v) = "" Then
        v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , _
                                        "特殊集計区分一覧表のCSVを選択してください")
        If VarType(v) = vbBoolean Then Exit Sub
    ElseIf Dir(CStr(v)) = "" Then
        v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , _
                                        "特殊集計区分一覧表のCSVを選択してください")
        If VarType(v) = vbBoolean Then Exit Sub
    End If

    ans = MsgBox("いま入っている O列（処分費〇）・P列（管材費〇）・AA列（処分費額）を" & vbCrLf & _
                 "消してから入れ直しますか？" & vbCrLf & vbCrLf & _
                 "［はい］消してから入れる　／　［いいえ］消さずに足す", _
                 vbQuestion + vbYesNoCancel)
    If ans = vbCancel Then Exit Sub
    doClear = (ans = vbYes)

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(SH_SLIDE)
    Set rows = ParseCsvText(ReadTextFile(CStr(v), charSetName), _
                            CStr(CfgVal(cfg, "区切り文字", ",")))
    If rows.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "CSVが空でした。", vbExclamation
        Exit Sub
    End If

    Set lay = TkDetect(rows, TkSpecOf(cfg), TkKinDefault(cfg))
    maxC = CLng(lay("列数"))
    kcol = TkCol(lay, "区分")
    head0 = CLng(lay("見出し行"))
    carry = CBool(lay("区分は見出し行"))

    If kcol = 0 Or TkCol(lay, "名称") = 0 Then
        Application.ScreenUpdating = True
        MsgBox "CSVの列の並びが読み取れませんでした。" & vbCrLf & vbCrLf & _
               "読み取った並び" & vbCrLf & TkLayText(lay) & vbCrLf & vbCrLf & _
               "［特殊集計区分CSV構造チェック］で並びを確かめ、" & vbCrLf & _
               "「" & SH_CONFIG & "」シートの「特殊集計CSV列指定」に" & vbCrLf & _
               "　区分=14,コード=1,名称=2,単位=3,規格1=4,規格2=5,金額=11" & vbCrLf & _
               "のように列を書いてから、もう一度実行してください。", vbExclamation
        Exit Sub
    End If

    '--- CSVを突合キーごとの待ち行列にする ---
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = 1
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
            nm = Trim$(ColVal(arr, TkCol(lay, "名称")))

            If kind <> "" And nm <> "" Then
                k = TkKey(Trim$(ColVal(arr, TkCol(lay, "コード"))), nm, _
                          TkKikaku(arr, lay), Trim$(ColVal(arr, TkCol(lay, "単位"))))
                If Not map.Exists(k) Then map.Add k, New Collection
                Set q = map(k)
                q.Add Array(kind, ToNum(ColVal(arr, TkCol(lay, "金額"))), i, curKubun, _
                            Trim$(ColVal(arr, TkCol(lay, "コード"))), nm, _
                            TkKikaku(arr, lay), Trim$(ColVal(arr, TkCol(lay, "単位"))), _
                            ToNum(ColVal(arr, TkCol(lay, "数量"))))
            End If
        End If
    Next i

    '--- 結果シート ---
    Set rs = FreshSheet(SH_TK_RESULT)
    rs.Range("A1").Value = "特殊集計区分CSV 取込結果"
    rs.Range("A1").Font.Bold = True
    rs.Range("A2").Value = "ファイル：" & Mid$(CStr(v), InStrRev(CStr(v), Application.PathSeparator) + 1)
    rs.Range("A3").Value = "読み取った並び：" & TkLayText(lay)
    rs.Range("A5").Value = "CSV行"
    rs.Range("B5").Value = "特殊集計区分"
    rs.Range("C5").Value = "判定"
    rs.Range("D5").Value = "コード"
    rs.Range("E5").Value = "名称"
    rs.Range("F5").Value = "規格"
    rs.Range("G5").Value = "単位"
    rs.Range("H5").Value = "CSVの数量"
    rs.Range("I5").Value = "CSVの金額"
    rs.Range("J5").Value = "計算表の行"
    rs.Range("K5").Value = "結果"
    With rs.Range("A5:K5")
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With
    rr = 5

    '--- 明細行を上から見て、キーが合うものを1つずつ使う ---
    lastRow = ws.Cells(ws.Rows.Count, SC_CODE).End(xlUp).Row
    If lastRow < FIRST_ROW Then lastRow = FIRST_ROW

    If doClear Then
        For r = FIRST_ROW To lastRow
            If TkIsDetail(ws, r) Then
                ws.Cells(r, SC_MK_SHOBU).ClearContents
                ws.Cells(r, SC_MK_KANZA).ClearContents
                ws.Cells(r, SC_SHA_O).ClearContents
            End If
        Next r
    End If

    For r = FIRST_ROW To lastRow
        If TkIsDetail(ws, r) Then
            k = TkKey(CStr(ws.Cells(r, SC_CODE).Value), CStr(ws.Cells(r, SC_NAME).Value), _
                      CStr(ws.Cells(r, SC_KIKAKU).Value), CStr(ws.Cells(r, SC_TANI).Value))
            If map.Exists(k) Then
                Set q = map(k)
                If q.Count > 0 Then
                    e = q(1)
                    q.Remove 1

                    rr = rr + 1
                    TkResRow rs, rr, e, r

                    If CStr(e(0)) = "処分" Then
                        If IsScrapName(CStr(e(5))) Then
                            rs.Cells(rr, 11).Value = "スクラップなので処分費にしませんでした"
                            rs.Cells(rr, 11).Font.Color = RGB(192, 0, 0)
                            nSkip = nSkip + 1
                        Else
                            ws.Cells(r, SC_MK_SHOBU).Value = "〇"
                            nSho = nSho + 1
                            If useKin And CDbl(e(1)) <> 0 Then
                                ws.Cells(r, SC_SHA_O).Value = CDbl(e(1))
                                nKin = nKin + 1
                                rs.Cells(rr, 11).Value = "O列に〇／AA列に金額を入れました"
                            Else
                                rs.Cells(rr, 11).Value = "O列に〇を入れました（金額はCSVに無し）"
                            End If
                            ' CSVの数量と計算表のK列が違えば、金額の意味が合っていない疑い
                            If CDbl(e(8)) <> 0 And IsNumeric(ws.Cells(r, SC_Q_ALL).Value) Then
                                If Abs(CDbl(e(8)) - CDbl(ws.Cells(r, SC_Q_ALL).Value)) > 0.001 Then
                                    rs.Cells(rr, 8).Interior.Color = CLR_INPUT
                                    rs.Cells(rr, 11).Value = CStr(rs.Cells(rr, 11).Value) & _
                                        "　※数量が計算表(" & ws.Cells(r, SC_Q_ALL).Value & ")と違います"
                                    warnQ = warnQ + 1
                                End If
                            End If
                        End If
                    ElseIf CStr(e(0)) = "管材" Then
                        ws.Cells(r, SC_MK_KANZA).Value = "〇"
                        nKan = nKan + 1
                        rs.Cells(rr, 11).Value = "P列に〇を入れました"
                    End If
                End If
            End If
        End If
    Next r

    '--- 使われなかったCSV行 ---
    Set rest = New Collection
    For Each kAny In map.Keys
        Set q = map(CStr(kAny))
        For i = 1 To q.Count
            rest.Add q(i)
        Next i
    Next kAny

    If rest.Count > 0 Then
        rr = rr + 2
        rs.Cells(rr, 1).Value = "▼ 計算表に同じ明細が見つからなかったCSV行（" & rest.Count & "行）" & _
            "　…　代価表の中の行かもしれません。手でO列・P列に〇を付けてください"
        rs.Cells(rr, 1).Font.Bold = True
        rs.Cells(rr, 1).Font.Color = RGB(192, 0, 0)
        For i = 1 To rest.Count
            rr = rr + 1
            TkResRow rs, rr, rest(i), 0
            rs.Cells(rr, 11).Value = "見つかりませんでした"
            rs.Cells(rr, 11).Font.Color = RGB(192, 0, 0)
        Next i
    End If

    rs.Columns("A").ColumnWidth = 7
    rs.Columns("B").ColumnWidth = 18
    rs.Columns("C").ColumnWidth = 8
    rs.Columns("D").ColumnWidth = 12
    rs.Columns("E").ColumnWidth = 26
    rs.Columns("F").ColumnWidth = 26
    rs.Columns("G").ColumnWidth = 6
    rs.Columns("H").ColumnWidth = 11
    rs.Columns("I").ColumnWidth = 13
    rs.Columns("J").ColumnWidth = 10
    rs.Columns("K").ColumnWidth = 46
    rs.Columns("I").NumberFormat = "#,##0"

    Application.ScreenUpdating = True
    ws.Activate

    MsgBox "特殊集計区分CSVを取り込みました。" & vbCrLf & vbCrLf & _
           "処分費〇（O列）　：" & nSho & " 行" & vbCrLf & _
           "　うち金額もAA列へ：" & nKin & " 行" & vbCrLf & _
           "管材費〇（P列）　：" & nKan & " 行" & vbCrLf & _
           "スクラップで除外　：" & nSkip & " 行" & vbCrLf & _
           "見つからなかった行：" & rest.Count & " 行" & vbCrLf & _
           IIf(warnQ > 0, "数量が合わない行　：" & warnQ & " 行" & vbCrLf, "") & vbCrLf & _
           "内容は「" & SH_TK_RESULT & "」シートで確認できます。" & vbCrLf & _
           IIf(rest.Count > 0 Or warnQ > 0, _
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

' 結果シートに1行書く
Private Sub TkResRow(ByVal rs As Worksheet, ByVal rr As Long, ByVal e As Variant, _
                     ByVal sheetRow As Long)
    rs.Cells(rr, 1).Value = e(2)
    rs.Cells(rr, 2).Value = e(3)
    rs.Cells(rr, 3).Value = TkKindName(CStr(e(0)))
    rs.Cells(rr, 4).Value = "'" & CStr(e(4))
    rs.Cells(rr, 5).Value = e(5)
    rs.Cells(rr, 6).Value = e(6)
    rs.Cells(rr, 7).Value = e(7)
    If CDbl(e(8)) <> 0 Then rs.Cells(rr, 8).Value = CDbl(e(8))
    If CDbl(e(1)) <> 0 Then rs.Cells(rr, 9).Value = CDbl(e(1))
    If sheetRow > 0 Then rs.Cells(rr, 10).Value = sheetRow
    rs.Range(rs.Cells(rr, 1), rs.Cells(rr, 11)).Borders.LineStyle = xlContinuous
End Sub


' 明細行かどうか（AC列の処分費額の式は明細行だけが持っている）
Private Function TkIsDetail(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    If Not ws.Cells(r, SC_SA_ALL).HasFormula Then Exit Function
    If NormText(CStr(ws.Cells(r, SC_NAME).Value)) = "" Then Exit Function
    TkIsDetail = True
End Function


' 突合キー（旧CSVと新CSVの突合と同じ作り方）
Private Function TkKey(ByVal code As String, ByVal nm As String, _
                       ByVal kikaku As String, ByVal tani As String) As String
    TkKey = NormText(code) & "|" & NormText(nm) & "|" & NormText(kikaku) & "|" & NormText(tani)
End Function


' 規格1・規格2をつなぐ（M00_Main の JoinKikaku と同じ形）
Private Function TkKikaku(ByRef arr() As String, ByVal lay As Object) As String
    Dim k1 As String, k2 As String
    k1 = Trim$(ColVal(arr, TkCol(lay, "規格1")))
    k2 = Trim$(ColVal(arr, TkCol(lay, "規格2")))
    If k1 = "" Then
        TkKikaku = k2
    ElseIf k2 = "" Then
        TkKikaku = k1
    Else
        TkKikaku = k1 & "　" & k2
    End If
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


' 見出しが無く、スライド用csvと同じ並びだったときに使う金額列
Private Function TkKinDefault(ByVal cfg As Object) As Long
    If cfg Is Nothing Then TkKinDefault = CSV_KIN1: Exit Function
    If NormText(CStr(CfgVal(cfg, "使用系列", "変更"))) = NormText("当初") Then
        TkKinDefault = CSV_KIN2
    Else
        TkKinDefault = CSV_KIN1
    End If
End Function


'--------------------------------------------------------------
' 列の並びを読み取る
'   戻り値は Dictionary。"区分" "コード" "名称" "単位" "規格1" "規格2"
'   "数量" "単価" "金額" "摘要" に列番号（0＝無し）、ほかに
'   "見出し行"（0＝無し）／"列数"／"区分は見出し行" が入る。
'--------------------------------------------------------------
Private Function TkDetect(ByVal rows As Collection, ByVal spec As String, _
                          ByVal kinDefault As Long) As Object
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
    Dim slideShape As Boolean

    Set lay = CreateObject("Scripting.Dictionary")
    lay.CompareMode = 1
    ks = Array("区分", "コード", "名称", "単位", "規格1", "規格2", "数量", "単価", "金額", "摘要")
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
    If CLng(lay("区分")) = 0 Then
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
    If CLng(lay("区分")) > 0 Then
        For i = head0 + 1 To rows.Count
            arr = rows(i)
            If TkLooksKubun(ColVal(arr, CLng(lay("区分")))) Then
                If TkDataCells(arr, maxC, CLng(lay("区分"))) >= 3 Then
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
            If TkDataCells(arr, maxC, CLng(lay("区分"))) >= 3 Then
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
            ' スライド用csvと同じ13列の並びかどうか
            If maxC >= 12 Then
                If sCode(1) * 10 >= nData * 6 And sText(2) * 10 >= nData * 6 _
                   And sTani(3) * 10 >= nData * 3 Then slideShape = True
            End If

            If slideShape Then
                lay("コード") = CSV_CODE
                lay("名称") = CSV_NAME
                lay("単位") = CSV_TANI
                lay("規格1") = CSV_KIKAKU1
                lay("規格2") = CSV_KIKAKU2
                lay("数量") = CSV_SURYO1
                lay("金額") = kinDefault
            Else
                ReDim used(1 To maxC)
                If Not CBool(lay("区分は見出し行")) Then
                    If CLng(lay("区分")) > 0 Then used(CLng(lay("区分"))) = True
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
                    ' 名称より右にある文字列の列を規格1・規格2に
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
    End If

    '--- 5) 設定シートの列指定を最後にかぶせる（手で書いた指定が優先）---
    Set ov = TkParseSpec(spec)
    For Each kk In ov.Keys
        lay(CStr(kk)) = CLng(ov(CStr(kk)))
    Next kk

    Set TkDetect = lay
End Function


Private Function TkCol(ByVal lay As Object, ByVal name As String) As Long
    If lay Is Nothing Then Exit Function
    If Not lay.Exists(name) Then Exit Function
    TkCol = CLng(lay(name))
End Function


' 見出しの文字から項目名を決める（「コード」の判定を「単価」より先に置く）
Private Function TkHeadName(ByVal s As String) As String
    Dim t As String
    t = NormText(s)
    If t = "" Then Exit Function

    If InStr(1, t, "特殊集計") > 0 Or InStr(1, t, "集計区分") > 0 Then TkHeadName = "区分": Exit Function
    If t = "区分" Then TkHeadName = "区分": Exit Function
    If InStr(1, t, "コード") > 0 Then TkHeadName = "コード": Exit Function
    If InStr(1, t, "名称") > 0 Or InStr(1, t, "品名") > 0 Then TkHeadName = "名称": Exit Function
    If InStr(1, t, "規格") > 0 Or InStr(1, t, "仕様") > 0 Then
        If InStr(1, t, "2") > 0 Then TkHeadName = "規格2" Else TkHeadName = "規格1"
        Exit Function
    End If
    If InStr(1, t, "単位") > 0 Then TkHeadName = "単位": Exit Function
    If InStr(1, t, "数量") > 0 Then TkHeadName = "数量": Exit Function
    If InStr(1, t, "金額") > 0 Then TkHeadName = "金額": Exit Function
    If InStr(1, t, "単価") > 0 Then TkHeadName = "単価": Exit Function
    If InStr(1, t, "摘要") > 0 Then TkHeadName = "摘要": Exit Function
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
    t = NormText(s)
    If t = "" Or Len(t) > 4 Then Exit Function
    If InStr(1, "|" & TK_TANI_LIST & "|", "|" & t & "|") > 0 Then TkIsTaniish = True
End Function


' 金額・単価・数量の列か（構造チェックで数値を隠すのに使う）
Private Function TkIsMoneyCol(ByVal lay As Object, ByVal c As Long) As Boolean
    If c = 0 Then Exit Function
    If c = TkCol(lay, "金額") Then TkIsMoneyCol = True: Exit Function
    If c = TkCol(lay, "単価") Then TkIsMoneyCol = True: Exit Function
    If c = TkCol(lay, "数量") Then TkIsMoneyCol = True
End Function


Private Function TkColRole(ByVal lay As Object, ByVal c As Long) As String
    Dim ks As Variant, kk As Variant
    ks = Array("区分", "コード", "名称", "単位", "規格1", "規格2", "数量", "単価", "金額", "摘要")
    For Each kk In ks
        If TkCol(lay, CStr(kk)) = c Then TkColRole = vbLf & "＝" & CStr(kk): Exit Function
    Next kk
End Function


Private Function TkLayText(ByVal lay As Object) As String
    Dim ks As Variant, kk As Variant
    Dim s As String
    ks = Array("区分", "コード", "名称", "単位", "規格1", "規格2", "数量", "金額")
    For Each kk In ks
        If s <> "" Then s = s & "／"
        s = s & CStr(kk) & "=" & IIf(TkCol(lay, CStr(kk)) = 0, "なし", CStr(TkCol(lay, CStr(kk))) & "列")
    Next kk
    s = s & vbCrLf & "見出し行=" & IIf(CLng(lay("見出し行")) = 0, "なし", CStr(lay("見出し行"))) & _
        "／区分の引き継ぎ=" & IIf(CBool(lay("区分は見出し行")), "する", "しない")
    TkLayText = s
End Function


' 「区分=14,コード=1,…」を読む。列はA,B,…の記号でも数字でもよい
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
    Select Case NormText(s)
        Case "区分", "特殊集計区分": TkSpecName = "区分"
        Case "コード":               TkSpecName = "コード"
        Case "名称":                 TkSpecName = "名称"
        Case "単位":                 TkSpecName = "単位"
        Case "規格", "規格1":        TkSpecName = "規格1"
        Case "規格2":                TkSpecName = "規格2"
        Case "数量":                 TkSpecName = "数量"
        Case "金額":                 TkSpecName = "金額"
        Case "単価":                 TkSpecName = "単価"
        Case "摘要":                 TkSpecName = "摘要"
    End Select
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
