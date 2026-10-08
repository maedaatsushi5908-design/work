'==============================================================
' M50_Tokushu  ―  特殊集計区分一覧表CSVから処分費・管材費を自動入力
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M50_Tokushu" に。
'
' 積算システムの「特殊集計区分一覧表」CSVを【スライド前】【スライド後】の
' 2本読み、スライド計算表の
'   O 列（処分費〇）  ← 集計区分名称に「処分」が入る行
'   P 列（管材費〇）  ← 集計区分名称に「水道／管材／管財」が入る行
'   AA列（処分費額 スライド前）← 前のCSVの金額
'   AB列（処分費額 スライド後）← 後のCSVの金額
' を自動で入れます。手入力は出来形数量だけになります。
'
' マクロは2つです。
'   ・特殊集計区分CSV構造チェック … 金額を出さずに列の並びだけ書き出す（送付用）
'   ・特殊集計区分CSV取込         … 実際にスライド計算表へ入れる
'
' 一覧表は「内訳の明細」ではなく「代価表の中の基礎単価」を並べたものなので、
' 内訳の行は次の順にコードでたどります。
'   1) F列 代価表等コード番号   （内訳に出てくる代価表のコード。これが本筋）
'   2) G列 親施工単価コード番号 （代価表が「―」のとき内訳に出てくるコード）
'   3) 基礎単価コード           （上の2つで当たらないときの最後の手段）
' それぞれについて、内訳に無ければ「_単価表」シート（代価表の中身の索引）を使って
' 所属する代価表までさかのぼり、その代価表の内訳の行に当てます。
'   例）G列=V0013 は内訳に無いが、単価表部では代価表G0001の中にある
'       → 内訳のG0001 ／再掘削工 の行に当てる
' 同じコードの明細が複数あるときは、親施工数量・名称で絞り込みます。
' 同じ内訳の行に当たる行が複数あれば、金額を合算します。
'==============================================================
Option Explicit

' 単位らしさの判定に使う語彙（縦棒区切り）
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
'   「代価表」「親」「コード」の順を守ること（基礎単価コードが単価に取られないように）
Private Function TkHeadName(ByVal s As String) As String
    Dim t As String
    t = NormText(s)
    If t = "" Then Exit Function

    If InStr(1, t, "特殊集計") > 0 Or InStr(1, t, "集計区分") > 0 Then
        ' 「集計区分」は番号、「集計区分名称」が名前。判定には名前を使う
        If InStr(1, t, "名称") > 0 Then TkHeadName = "区分名称" Else TkHeadName = "区分"
        Exit Function
    End If
    If t = "区分" Then TkHeadName = "区分": Exit Function
    If InStr(1, t, "代価表") > 0 Then TkHeadName = "代価表コード": Exit Function
    If InStr(1, t, "親") > 0 Then
        If InStr(1, t, "数量") > 0 Then
            TkHeadName = "親数量"
        ElseIf InStr(1, t, "コード") > 0 Then
            TkHeadName = "親コード"
        End If
        Exit Function
    End If
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
    Select Case NormText(s)
        Case "区分", "集計区分", "特殊集計区分":       TkSpecName = "区分"
        Case "区分名称", "集計区分名称":                TkSpecName = "区分名称"
        Case "コード", "基礎単価コード":                TkSpecName = "コード"
        Case "名称":                                    TkSpecName = "名称"
        Case "単位":                                    TkSpecName = "単位"
        Case "規格", "規格1":                           TkSpecName = "規格1"
        Case "規格2":                                   TkSpecName = "規格2"
        Case "数量":                                    TkSpecName = "数量"
        Case "金額":                                    TkSpecName = "金額"
        Case "単価":                                    TkSpecName = "単価"
        Case "摘要":                                    TkSpecName = "摘要"
        Case "代価表コード", "代価表", "代価表等コード番号": TkSpecName = "代価表コード"
        Case "親コード", "親施工単価コード番号":        TkSpecName = "親コード"
        Case "親数量", "親施工数量":                    TkSpecName = "親数量"
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
