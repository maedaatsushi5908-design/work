Attribute VB_Name = "M_Link"
'==================================================================
' M_Link - 総括表のシート間リンクを、口径ひとつで切り替わる数式にする
'
' 総括表（土工事）は、試掘シート・管工シートから値を1本ずつ手で
' リンクしている。列が口径に対応しているので（J=50, M=400, R=50,
' S=75, T=400）、口径の部分だけを設定シートから読む数式に置き換えれば、
' 工事が変わっても設定の口径を直すだけで全部が繋ぎ変わる。
'
'   取り込み前:  ='試掘（舗400'!P4
'   置き換え後:  =IFERROR(INDIRECT("'試掘（舗"&リンク設定!$C$7&"'!P4"),0)
'
'   マクロ:
'     リンク設定を取り込む … 今あるリンクを読み取って設定シートを作る
'     数式を再生成       … 設定どおりに INDIRECT 数式を書き込む
'     直接参照に戻す     … 普通のリンク（='試掘（舗400'!P4）に戻す
'     リンクを点検       … 参照先が空・シート欠落・リンク漏れを調べる
'==================================================================
Option Explicit

Public Const LNK_SHEET As String = "リンク設定"
Private Const DIA_HDR As Long = 5        ' 口径対応表の見出し行
Private Const DIA_FIRST As Long = 6
Private Const DIA_LAST As Long = 25
Private Const LNK_HDR As Long = 28       ' リンク一覧の見出し行
Private Const LNK_FIRST As Long = 29

Private Const TERM_PAT As String = "'([^']+)'!(\$?[A-Z]{1,3}\$?[0-9]+)"

'==================================================================
' 1) 今あるリンクを読み取って設定シートを作る
'==================================================================
Public Sub リンク設定を取り込む()
    Dim srcName As String, ws As Worksheet, cfg As Worksheet
    Dim c As Range, f As String
    Dim colDia As Object, colHead As Object
    Dim items As Collection, it As Variant
    Dim r As Long, i As Long, n As Long, skipped As Long

    srcName = AskTargetSheet()
    If Len(srcName) = 0 Then Exit Sub

    Set ws = FindSheet(ThisWorkbook, srcName)
    If ws Is Nothing Then
        MsgBox "シートが見つかりません: " & srcName, vbExclamation
        Exit Sub
    End If

    Set colDia = CreateObject("Scripting.Dictionary")   ' 列 → 口径の集合(文字列)
    Set colHead = CreateObject("Scripting.Dictionary")  ' 列 → 見出し
    Set items = New Collection

    ' --- 全リンクを拾って、列ごとの口径を集める --------------------
    For Each c In ws.UsedRange
        If Not HasSheetRef(c) Then GoTo NextCell
        f = c.Formula
        Dim terms As Collection: Set terms = ParseTerms(f)
        If terms.Count = 0 Then GoTo NextCell

        ' 解析できた項の数と、数式中の "!" の数が合わなければ手を出さない。
        ' 引用符なしのシート名など、想定外の書き方が混ざっている合図。
        If terms.Count <> CountChar(f, "!") Then
            items.Add Array(c.Address(False, False), ColLetterOf(c), f, "?")
            GoTo NextCell
        End If

        Dim colL As String: colL = Split(c.Address(False, False), "$")(0)
        colL = ColLetterOf(c)

        Dim dset As String: dset = ""
        For i = 1 To terms.Count
            Dim d As String: d = TrailDigits(terms(i)(0))
            If Len(d) = 0 Then GoTo NextCell            ' 口径の付かない参照は対象外
            If InStr(1, "," & dset & ",", "," & d & ",") = 0 Then
                dset = IIf(Len(dset) = 0, d, dset & "," & d)
            End If
        Next i

        If Not colDia.Exists(colL) Then colDia(colL) = ""
        Dim cur As String: cur = colDia(colL)
        For Each it In Split(dset, ",")
            If InStr(1, "," & cur & ",", "," & it & ",") = 0 Then
                cur = IIf(Len(cur) = 0, CStr(it), cur & "," & CStr(it))
            End If
        Next it
        colDia(colL) = cur
        If Not colHead.Exists(colL) Then colHead(colL) = HeaderOf(ws, c.Column)

        items.Add Array(c.Address(False, False), colL, f, dset)
NextCell:
    Next c

    If items.Count = 0 Then
        MsgBox "シート間リンクが見つかりませんでした。" & vbCrLf & _
               "（口径の付いたシート名を参照しているものが対象です）", vbInformation
        Exit Sub
    End If

    ' --- 設定シートを作る ------------------------------------------
    Set cfg = ResetConfigSheet()
    cfg.Range("A2").Value = "対象シート"
    cfg.Range("B2").Value = ws.Name

    cfg.Range("A4").Value = "【口径対応表】　ここの口径を書き換えると、下のリンクが全部その口径に切り替わります"
    cfg.Range("A4").Font.Bold = True
    cfg.Cells(DIA_HDR, 1).Value = "列"
    cfg.Cells(DIA_HDR, 2).Value = "総括表の見出し"
    cfg.Cells(DIA_HDR, 3).Value = "口径"
    cfg.Cells(DIA_HDR, 4).Value = "備考"
    StyleHeader cfg.Range(cfg.Cells(DIA_HDR, 1), cfg.Cells(DIA_HDR, 4))

    r = DIA_FIRST
    Dim k As Variant
    For Each k In SortedKeys(colDia)
        If r > DIA_LAST Then Exit For
        cfg.Cells(r, 1).Value = k
        cfg.Cells(r, 2).Value = colHead(k)
        If InStr(colDia(k), ",") > 0 Then
            cfg.Cells(r, 3).Value = ""
            cfg.Cells(r, 4).Value = "★口径が一意でない(" & colDia(k) & ")。この列は自動化できません"
            cfg.Cells(r, 4).Font.Color = RGB(192, 0, 0)
        Else
            cfg.Cells(r, 3).Value = colDia(k)
            cfg.Cells(r, 3).Interior.Color = RGB(255, 242, 204)
        End If
        r = r + 1
    Next k

    ' --- リンク一覧 --------------------------------------------------
    cfg.Range("A27").Value = "【リンク一覧】　有効を空欄にするとその行は書き換えません"
    cfg.Range("A27").Font.Bold = True
    Dim hdr As Variant
    hdr = Array("No", "有効", "転記先セル", "列", "口径セル", "テンプレート", _
                "取り込み時の数式", "取り込み時の値", "備考")
    For i = 0 To UBound(hdr)
        cfg.Cells(LNK_HDR, i + 1).Value = hdr(i)
    Next i
    StyleHeader cfg.Range(cfg.Cells(LNK_HDR, 1), cfg.Cells(LNK_HDR, UBound(hdr) + 1))

    r = LNK_FIRST
    For i = 1 To items.Count
        Dim addr As String, cl As String, fml As String, ds As String
        addr = items(i)(0): cl = items(i)(1): fml = items(i)(2): ds = items(i)(3)

        n = n + 1
        cfg.Cells(r, 1).Value = n
        cfg.Cells(r, 3).Value = addr
        cfg.Cells(r, 4).Value = cl
        cfg.Cells(r, 7).Value = "'" & fml
        cfg.Cells(r, 8).Value = NumOrEmpty(ws.Range(addr))

        If ds = "?" Or InStr(ds, ",") > 0 Or InStr(colDia(cl), ",") > 0 Then
            ' 1つのセルに複数の口径が混ざる、または列の口径が一意でない
            cfg.Cells(r, 2).Value = ""
            If ds = "?" Then
                cfg.Cells(r, 9).Value = "★解析できない書き方です。手動のまま残します"
            Else
                cfg.Cells(r, 9).Value = "★口径が混在(" & ds & ")。手動のまま残します"
            End If
            cfg.Cells(r, 9).Font.Color = RGB(192, 0, 0)
            skipped = skipped + 1
        Else
            cfg.Cells(r, 2).Value = "○"
            cfg.Cells(r, 5).Value = DiaCellAddress(cfg, cl)
            cfg.Cells(r, 6).Value = "'" & MakeTemplate(fml, ds)
        End If
        r = r + 1
    Next i

    cfg.Columns.AutoFit
    If cfg.Columns(6).ColumnWidth > 55 Then cfg.Columns(6).ColumnWidth = 55
    If cfg.Columns(7).ColumnWidth > 45 Then cfg.Columns(7).ColumnWidth = 45
    cfg.Activate
    cfg.Range("A1").Select

    MsgBox "「" & LNK_SHEET & "」を作りました。" & vbCrLf & vbCrLf & _
           "対象シート : " & ws.Name & vbCrLf & _
           "リンク総数 : " & items.Count & " 本" & vbCrLf & _
           "自動化対象 : " & (items.Count - skipped) & " 本" & vbCrLf & _
           "手動のまま : " & skipped & " 本" & vbCrLf & vbCrLf & _
           "口径対応表を確認してから「数式を再生成」を実行してください。", _
           vbInformation, "取り込み完了"
End Sub

'==================================================================
' 2) 設定どおりに INDIRECT 数式を書き込む
'==================================================================
Public Sub 数式を再生成()
    WriteFormulas True
End Sub

'==================================================================
' 3) 普通のリンクに戻す
'==================================================================
Public Sub 直接参照に戻す()
    WriteFormulas False
End Sub

Private Sub WriteFormulas(ByVal useIndirect As Boolean)
    Dim cfg As Worksheet, ws As Worksheet
    Dim r As Long, lastRow As Long, n As Long, ng As Long
    Dim tmpl As String, addr As String, diaCell As String, dia As String
    Dim newF As String, msg As String
    Dim scr As Boolean, calc As XlCalculation

    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then
        MsgBox "「" & LNK_SHEET & "」がありません。" & vbCrLf & _
               "先に「リンク設定を取り込む」を実行してください。", vbExclamation
        Exit Sub
    End If

    Set ws = FindSheet(ThisWorkbook, CStr(cfg.Range("B2").Value))
    If ws Is Nothing Then
        MsgBox "対象シートが見つかりません: " & cfg.Range("B2").Value, vbExclamation
        Exit Sub
    End If

    If MsgBox(ws.Name & " の数式を" & _
              IIf(useIndirect, "INDIRECT 形式", "直接参照") & "に書き換えます。" & vbCrLf & vbCrLf & _
              "書き換える前にバックアップシートを作ります。続けますか？", _
              vbYesNo + vbQuestion, "確認") <> vbYes Then Exit Sub

    BackupSheet ws

    scr = Application.ScreenUpdating: calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastRow = cfg.Cells(cfg.Rows.Count, 3).End(xlUp).Row
    For r = LNK_FIRST To lastRow
        If Len(Norm(cfg.Cells(r, 2).Value)) = 0 Then GoTo NextRow   ' 有効でない
        addr = Trim$(CStr(cfg.Cells(r, 3).Value))
        tmpl = Trim$(CStr(cfg.Cells(r, 6).Value))
        diaCell = Trim$(CStr(cfg.Cells(r, 5).Value))
        If Len(addr) = 0 Or Len(tmpl) = 0 Or Len(diaCell) = 0 Then GoTo NextRow

        dia = Trim$(CStr(cfg.Range(diaCell).Value))
        If Len(dia) = 0 Then
            cfg.Cells(r, 9).Value = "口径が空欄のため書き換えていません"
            ng = ng + 1
            GoTo NextRow
        End If

        If useIndirect Then
            newF = BuildIndirect(tmpl, diaCell)
        Else
            newF = Replace(tmpl, "#", dia)
        End If

        On Error Resume Next
        ws.Range(addr).Formula = newF
        If Err.Number <> 0 Then
            cfg.Cells(r, 9).Value = "書き込み失敗: " & Err.Description
            Err.Clear
            ng = ng + 1
        Else
            cfg.Cells(r, 9).Value = ""
            n = n + 1
        End If
        On Error GoTo 0
NextRow:
    Next r

    Application.Calculation = calc
    Application.ScreenUpdating = scr
    Application.CalculateFull

    msg = n & " 本を書き換えました。"
    If ng > 0 Then msg = msg & vbCrLf & ng & " 本は書き換えていません（備考欄を確認）。"
    If useIndirect Then
        msg = msg & vbCrLf & vbCrLf & _
              "以後、口径を変えるときは「" & LNK_SHEET & "」の口径欄" & vbCrLf & _
              "（黄色いセル）を書き換えるだけで全部が繋ぎ変わります。"
    End If
    MsgBox msg, vbInformation, "完了"
End Sub

'------------------------------------------------------------------
' テンプレート → INDIRECT 数式
'   ='試掘（舗#'!P4
'     → =IFERROR(INDIRECT("'試掘（舗"&リンク設定!$C$7&"'!P4"),0)
'------------------------------------------------------------------
Private Function BuildIndirect(ByVal tmpl As String, ByVal diaCell As String) As String
    Dim re As Object, ms As Object, m As Object
    Dim out As String, pos As Long, rep As String
    Dim pre As String, addr As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "'([^']+)#'!(\$?[A-Z]{1,3}\$?[0-9]+)"

    Set ms = re.Execute(tmpl)
    If ms.Count = 0 Then BuildIndirect = Replace(tmpl, "#", ""): Exit Function

    pos = 1
    For Each m In ms
        pre = m.SubMatches(0)
        addr = m.SubMatches(1)
        rep = "IFERROR(INDIRECT(""'" & pre & """&" & _
              "'" & LNK_SHEET & "'!" & diaCell & "&""'!" & addr & """),0)"
        out = out & Mid$(tmpl, pos, m.FirstIndex + 1 - pos) & rep
        pos = m.FirstIndex + 1 + m.Length
    Next m
    out = out & Mid$(tmpl, pos)
    BuildIndirect = out
End Function

'==================================================================
' 4) 点検
'==================================================================
Public Sub リンクを点検()
    Dim cfg As Worksheet, ws As Worksheet, rep As Worksheet
    Dim r As Long, lastRow As Long, rr As Long
    Dim tmpl As String, dia As String, sn As String, addr As String
    Dim used As Object, terms As Collection, i As Long
    Dim nEmpty As Long, nMiss As Long, nOrphan As Long

    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then
        MsgBox "先に「リンク設定を取り込む」を実行してください。", vbExclamation
        Exit Sub
    End If
    Set ws = FindSheet(ThisWorkbook, CStr(cfg.Range("B2").Value))
    If ws Is Nothing Then MsgBox "対象シートがありません。", vbExclamation: Exit Sub

    Set used = CreateObject("Scripting.Dictionary")

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("リンク点検").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set rep = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    rep.Name = "リンク点検"

    rep.Range("A1").Value = "リンク点検　" & Format$(Now, "yyyy/mm/dd hh:nn") & "　対象: " & ws.Name
    rep.Range("A1").Font.Bold = True
    rep.Range("A3").Value = "転記先"
    rep.Range("B3").Value = "参照先シート"
    rep.Range("C3").Value = "参照セル"
    rep.Range("D3").Value = "値"
    rep.Range("E3").Value = "所見"
    StyleHeader rep.Range("A3:E3")

    rr = 4
    lastRow = cfg.Cells(cfg.Rows.Count, 3).End(xlUp).Row
    For r = LNK_FIRST To lastRow
        addr = Trim$(CStr(cfg.Cells(r, 3).Value))
        tmpl = Trim$(CStr(cfg.Cells(r, 6).Value))
        If Len(addr) = 0 Or Len(tmpl) = 0 Then GoTo NextRow
        dia = Trim$(CStr(cfg.Range(Trim$(CStr(cfg.Cells(r, 5).Value))).Value))
        If Len(dia) = 0 Then GoTo NextRow

        Set terms = ParseTerms(Replace(tmpl, "#", dia))
        For i = 1 To terms.Count
            sn = terms(i)(0)
            used(sn) = True
            rep.Cells(rr, 1).Value = addr
            rep.Cells(rr, 2).Value = sn
            rep.Cells(rr, 3).Value = terms(i)(1)

            Dim tgt As Worksheet
            Set tgt = FindSheet(ThisWorkbook, sn)
            If tgt Is Nothing Then
                rep.Cells(rr, 5).Value = "★シートがありません"
                rep.Cells(rr, 5).Interior.Color = RGB(255, 199, 206)
                nMiss = nMiss + 1
            Else
                Dim v As Variant
                v = NumOrEmpty(tgt.Range(Replace(terms(i)(1), "$", "")))
                If Not IsEmpty(v) Then rep.Cells(rr, 4).Value = v
                If IsEmpty(v) Or v = 0 Then
                    rep.Cells(rr, 5).Value = "参照先が空（この口径は未使用かもしれません）"
                    rep.Cells(rr, 5).Interior.Color = RGB(255, 235, 156)
                    nEmpty = nEmpty + 1
                Else
                    rep.Cells(rr, 5).Value = "OK"
                End If
            End If
            rr = rr + 1
        Next i
NextRow:
    Next r

    ' リンクされていないのに中身のあるシートを探す（張り忘れ）
    rr = rr + 1
    rep.Cells(rr, 1).Value = "【リンクされていないシート】"
    rep.Cells(rr, 1).Font.Bold = True
    rr = rr + 1
    Dim sh As Worksheet, pre As String, anyData As Boolean
    For Each sh In ThisWorkbook.Worksheets
        If used.Exists(sh.Name) Then GoTo NextSheet
        If Len(TrailDigits(sh.Name)) = 0 Then GoTo NextSheet   ' 口径付きシートのみ
        anyData = (Application.WorksheetFunction.Count(sh.UsedRange) > 0) And _
                  (Application.WorksheetFunction.Sum(sh.UsedRange) <> 0)
        rep.Cells(rr, 1).Value = sh.Name
        rep.Cells(rr, 2).Value = IIf(sh.Visible = xlSheetVisible, "表示", "非表示")
        If anyData Then
            rep.Cells(rr, 5).Value = "★中身があるのにリンクされていません。張り忘れの可能性"
            rep.Cells(rr, 5).Interior.Color = RGB(255, 199, 206)
            nOrphan = nOrphan + 1
        Else
            rep.Cells(rr, 5).Value = "空。未使用の雛形とみられます"
        End If
        rr = rr + 1
NextSheet:
    Next sh

    rep.Columns.AutoFit
    rep.Activate

    MsgBox "点検が終わりました。" & vbCrLf & vbCrLf & _
           "参照先が空     : " & nEmpty & " 件" & vbCrLf & _
           "シート欠落     : " & nMiss & " 件" & vbCrLf & _
           "張り忘れの疑い : " & nOrphan & " 件" & vbCrLf & vbCrLf & _
           "詳しくは「リンク点検」シートを見てください。", vbInformation, "点検結果"
End Sub

'==================================================================
' 補助
'==================================================================
Private Function AskTargetSheet() As String
    Dim s As String
    s = InputBox("リンクを取り込む総括表シートの名前を入れてください。" & vbCrLf & vbCrLf & _
                 "例: 総括表（土工事）", "対象シート", "総括表（土工事）")
    AskTargetSheet = Trim$(s)
End Function

Private Function HasSheetRef(ByVal c As Range) As Boolean
    Dim f As String
    If Not c.HasFormula Then Exit Function
    f = c.Formula
    HasSheetRef = (InStr(f, "!") > 0) And (InStr(f, "表紙") = 0)
End Function

' 数式から 'シート名'!セル の並びを取り出す
Private Function ParseTerms(ByVal f As String) As Collection
    Dim re As Object, ms As Object, m As Object, out As Collection
    Set out = New Collection
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = TERM_PAT
    Set ms = re.Execute(f)
    For Each m In ms
        out.Add Array(m.SubMatches(0), m.SubMatches(1))
    Next m
    Set ParseTerms = out
End Function

' 文字列の末尾に続く数字を返す（"試掘（舗400" → "400"）
Private Function TrailDigits(ByVal s As String) As String
    Dim i As Long, ch As String, buf As String
    For i = Len(s) To 1 Step -1
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            buf = ch & buf
        Else
            Exit For
        End If
    Next i
    TrailDigits = buf
End Function

' 数式の中の該当口径を # に置き換える
Private Function MakeTemplate(ByVal f As String, ByVal dia As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "'([^']+)" & dia & "'!"
    MakeTemplate = re.Replace(f, "'$1#'!")
End Function

Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    CountChar = Len(s) - Len(Replace(s, ch, ""))
End Function

Private Function ColLetterOf(ByVal c As Range) As String
    ColLetterOf = Split(c.Address(True, False), "$")(0)
End Function

Private Function HeaderOf(ByVal ws As Worksheet, ByVal col As Long) As String
    Dim r As Long, v As Variant
    For r = 1 To 12
        v = ws.Cells(r, col).Value
        If Not IsError(v) Then
            If Len(Norm(v)) > 0 And Left$(CStr(v), 1) <> "=" Then
                HeaderOf = Trim$(Replace(Replace(CStr(v), vbLf, " "), vbCr, " "))
                Exit Function
            End If
        End If
    Next r
End Function

Private Function DiaCellAddress(ByVal cfg As Worksheet, ByVal colL As String) As String
    Dim r As Long
    For r = DIA_FIRST To DIA_LAST
        If Norm(cfg.Cells(r, 1).Value) = Norm(colL) Then
            DiaCellAddress = "$C$" & r
            Exit Function
        End If
    Next r
End Function

Private Function SortedKeys(ByVal d As Object) As Variant
    Dim k As Variant, arr() As String, n As Long, i As Long, j As Long, t As String
    ReDim arr(0 To d.Count - 1)
    For Each k In d.Keys
        arr(n) = CStr(k): n = n + 1
    Next k
    ' 列文字の長さ→辞書順
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If Len(arr(j)) < Len(arr(i)) Or _
               (Len(arr(j)) = Len(arr(i)) And arr(j) < arr(i)) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    SortedKeys = arr
End Function

Private Function ResetConfigSheet() As Worksheet
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(LNK_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = LNK_SHEET
    ws.Range("A1").Value = "リンク設定"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    Set ResetConfigSheet = ws
End Function

Private Sub StyleHeader(ByVal rg As Range)
    rg.Font.Bold = True
    rg.Interior.Color = RGB(220, 230, 241)
    rg.HorizontalAlignment = xlCenter
End Sub

Private Sub BackupSheet(ByVal ws As Worksheet)
    Dim nm As String, bk As Worksheet
    nm = "BK_" & Format$(Now, "mmdd_hhnn")
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set bk = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    bk.Name = nm
    Application.DisplayAlerts = True
End Sub
