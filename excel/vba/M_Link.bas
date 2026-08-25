Attribute VB_Name = "M_Link"
'==================================================================
' M_Link - 総括表のシート間リンクを、口径で切り替わる数式にする
'
' 総括表（土工事）は、試掘シート・管工シートから値を1本ずつ手で
' リンクしている。列が口径に対応しているので（J=50, M=400, R=50,
' S=75, T=400）、口径の部分を設定シートから読む形にすれば、
' 工事が変わっても設定を直すだけで全部が繋ぎ変わる。
'
' さらに、いま空欄の口径列（K/L/N/U など）にも同じ数式を入れておけば、
' 次の工事でその口径を使ったときに自動で値が出る。ただし総括表には
' 「行ごとに口径が分かれている区間」があり、そこへ一律に入れると
' 二重計上になるため、取り込み時に展開の可否を判定して分ける。
'
'   マクロ:
'     リンク設定を取り込む      今あるリンクを読み取って設定シートを作る
'     直接リンクで数式を作る    ='試掘（舗400'!P4 の形。参照先が読める（推奨）
'     INDIRECTで数式を作る      設定の口径を変えるだけで繋ぎ変わる形
'     リンクを点検              参照先が空・シート欠落・張り忘れを調べる
'==================================================================
Option Explicit

Public Const LNK_SHEET As String = "リンク設定"
Private Const DIA_HDR As Long = 5        ' 口径対応表の見出し行
Private Const DIA_FIRST As Long = 6
Private Const DIA_LAST As Long = 30
Private Const LNK_HDR As Long = 33       ' リンク一覧の見出し行
Private Const LNK_FIRST As Long = 34

Private Const TERM_PAT As String = "'([^']+)'!(\$?[A-Z]{1,3}\$?[0-9]+)"

' 列1つ分の情報
Private Type TCol
    Letter  As String
    Header  As String
    Kei     As String     ' 系統（見出しの「（」より前）
    Dia      As String    ' 口径
    Row_     As Long      ' 口径対応表での行
    Guessed As Boolean    ' 推定値かどうか
End Type

'==================================================================
' 1) 今あるリンクを読み取って設定シートを作る
'==================================================================
Public Sub リンク設定を取り込む()
    Dim srcName As String, ws As Worksheet, cfg As Worksheet
    Dim c As Range, f As String
    Dim cols() As TCol, nCol As Long
    Dim grp As Object, keys As Object
    Dim i As Long, r As Long, n As Long
    Dim nExp As Long, nFix As Long, nSkip As Long

    srcName = AskTargetSheet()
    If Len(srcName) = 0 Then Exit Sub
    Set ws = FindSheet(ThisWorkbook, srcName)
    If ws Is Nothing Then
        MsgBox "シートが見つかりません: " & srcName, vbExclamation
        Exit Sub
    End If

    ' grp: "行|系統" → Array(列の並び, テンプレート, 口径, 状態, 理由)
    Set grp = CreateObject("Scripting.Dictionary")
    Set keys = CreateObject("Scripting.Dictionary")   ' "系統|接頭辞|セル" → 行の集合

    ReDim cols(0 To 60)

    ' --- 全リンクを走査 --------------------------------------------
    For Each c In ws.UsedRange
        If Not HasSheetRef(c) Then GoTo NextCell
        f = c.Formula

        ' 前回このマクロが入れた INDIRECT 形式なら、普通の参照に読み替えてから解析する
        f = Normalize(f)

        Dim terms As Collection: Set terms = ParseTerms(f)
        If terms.Count = 0 Then GoTo NextCell
        If terms.Count <> CountChar(f, "!") Then GoTo NextCell   ' 想定外の書き方

        Dim colL As String: colL = ColLetterOf(c)
        Dim kei As String: kei = KeiOf(ws, c.Column)

        ' 口径と、接頭辞+セルの並び（署名）を作る
        Dim dia As String, sig As String, bad As Boolean
        dia = "": sig = "": bad = False
        For i = 1 To terms.Count
            Dim pre As String, addr As String, d As String
            pre = terms(i)(0): addr = Replace(terms(i)(1), "$", "")
            d = TrailDigits(pre)
            If Len(d) = 0 Then bad = True: Exit For
            If Len(dia) = 0 Then
                dia = d
            ElseIf dia <> d Then
                bad = True: Exit For        ' 1セルに複数の口径が混ざる
            End If
            sig = sig & "|" & Left$(pre, Len(pre) - Len(d)) & "!" & addr
        Next i
        If bad Then GoTo NextCell

        ' 列の情報を控える
        AddCol cols, nCol, colL, HeaderOf(ws, c.Column), kei, dia

        ' 「系統|接頭辞|セル」がどの行で使われているかを記録
        Dim parts As Variant, p As Variant
        parts = Split(Mid$(sig, 2), "|")
        For Each p In parts
            Dim kk As String: kk = kei & "|" & p
            If Not keys.Exists(kk) Then keys(kk) = ""
            If InStr(1, "," & keys(kk) & ",", "," & c.Row & ",") = 0 Then
                keys(kk) = IIf(Len(keys(kk)) = 0, CStr(c.Row), keys(kk) & "," & c.Row)
            End If
        Next p

        ' 行×系統でまとめる
        Dim gk As String: gk = c.Row & "|" & kei
        If grp.Exists(gk) Then
            Dim cur As Variant: cur = grp(gk)
            If cur(1) <> sig Then cur(3) = "NG_SIG"          ' 列ごとに参照セルが違う
            cur(0) = cur(0) & "," & colL
            grp(gk) = cur
        Else
            grp(gk) = Array(colL, sig, MakeTemplate(f, dia), "", "")
        End If
NextCell:
    Next c

    If grp.Count = 0 Then
        MsgBox "口径の付いたシートを参照するリンクが見つかりませんでした。", vbInformation
        Exit Sub
    End If

    ' --- 展開の可否を判定 ------------------------------------------
    Dim gkv As Variant
    For Each gkv In grp.Keys
        Dim g As Variant: g = grp(gkv)
        If g(3) = "NG_SIG" Then
            g(4) = "列ごとに参照セルが違う（シートの行数が口径で異なる）"
        Else
            Dim shared_ As String: shared_ = ""
            Dim kei2 As String: kei2 = Split(gkv, "|")(1)
            For Each p In Split(Mid$(g(1), 2), "|")
                Dim rowsFor As String: rowsFor = keys(kei2 & "|" & p)
                Dim q As Variant
                For Each q In Split(rowsFor, ",")
                    If InStr(1, "," & shared_ & ",", "," & q & ",") = 0 Then
                        shared_ = IIf(Len(shared_) = 0, CStr(q), shared_ & "," & CStr(q))
                    End If
                Next q
            Next p
            If UBound(Split(shared_, ",")) > 0 Then
                g(3) = "NG_SHARE"
                g(4) = "同じ参照セルを行 " & shared_ & " が分け合う"
            Else
                g(3) = "OK"
            End If
        End If
        grp(gkv) = g
    Next gkv

    ' --- 未リンクの列も口径対応表に加える --------------------------
    AddSiblingCols ws, cols, nCol
    GuessDiameters cols, nCol

    ' --- 設定シートを作る ------------------------------------------
    Set cfg = ResetConfigSheet()
    cfg.Range("A2").Value = "対象シート"
    cfg.Range("B2").Value = ws.Name

    cfg.Range("A4").Value = "【口径対応表】　口径を書き換えると、その列のリンクが全部その口径に切り替わります。空欄の列には数式を入れません"
    cfg.Range("A4").Font.Bold = True
    WriteRow cfg, DIA_HDR, Array("列", "総括表の見出し", "口径", "系統", "備考")
    StyleHeader cfg.Range(cfg.Cells(DIA_HDR, 1), cfg.Cells(DIA_HDR, 5))

    r = DIA_FIRST
    For i = 0 To nCol - 1
        If r > DIA_LAST Then Exit For
        cols(i).Row_ = r
        cfg.Cells(r, 1).Value = cols(i).Letter
        cfg.Cells(r, 2).Value = cols(i).Header
        cfg.Cells(r, 3).Value = cols(i).Dia
        cfg.Cells(r, 3).Interior.Color = RGB(255, 242, 204)
        cfg.Cells(r, 4).Value = cols(i).Kei
        If cols(i).Guessed Then
            cfg.Cells(r, 5).Value = "★推定値。合っているか確かめてください"
            cfg.Cells(r, 5).Font.Color = RGB(192, 0, 0)
        ElseIf Len(cols(i).Dia) = 0 Then
            cfg.Cells(r, 5).Value = "口径が分からないため空欄。入れればこの列にも数式が入ります"
        End If
        r = r + 1
    Next i

    ' --- リンク一覧 --------------------------------------------------
    cfg.Cells(LNK_HDR - 2, 1).Value = "【リンク一覧】　展開に○が付いた行は、同じ系統の全ての口径列に数式を入れます"
    cfg.Cells(LNK_HDR - 2, 1).Font.Bold = True
    WriteRow cfg, LNK_HDR, Array("No", "有効", "展開", "対象行", "系統", "今ある列", _
                                 "テンプレート", "取り込み時の値", "判定", "備考")
    StyleHeader cfg.Range(cfg.Cells(LNK_HDR, 1), cfg.Cells(LNK_HDR, 10))

    r = LNK_FIRST
    Dim sk As Variant
    For Each sk In SortedGroupKeys(grp)
        Dim gg As Variant: gg = grp(sk)
        Dim gr As Long: gr = CLng(Split(sk, "|")(0))
        n = n + 1
        cfg.Cells(r, 1).Value = n
        ' 列ごとに参照セルが違う行は、1つのテンプレートで書くと値が壊れる。
        ' 書き込み対象から外し、手作業のまま残す。
        cfg.Cells(r, 2).Value = IIf(gg(3) = "NG_SIG", "", "○")
        cfg.Cells(r, 4).Value = gr
        cfg.Cells(r, 5).Value = Split(sk, "|")(1)
        cfg.Cells(r, 6).Value = gg(0)
        cfg.Cells(r, 7).Value = "'" & gg(2)
        cfg.Cells(r, 8).Value = SumOfCols(ws, gr, CStr(gg(0)))

        If gg(3) = "OK" Then
            cfg.Cells(r, 3).Value = "○"
            cfg.Cells(r, 9).Value = "展開可"
            nExp = nExp + 1
        ElseIf gg(3) = "NG_SIG" Then
            cfg.Cells(r, 3).Value = ""
            cfg.Cells(r, 9).Value = "対象外"
            cfg.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
            cfg.Cells(r, 10).Value = gg(4) & "。書き込みません（手作業のまま）"
            nSkip = nSkip + 1
        Else
            cfg.Cells(r, 3).Value = ""
            cfg.Cells(r, 9).Value = "展開不可"
            cfg.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
            cfg.Cells(r, 10).Value = gg(4) & "。今ある列だけに入れます"
            nFix = nFix + 1
        End If
        r = r + 1
    Next sk

    cfg.Columns.AutoFit
    If cfg.Columns(7).ColumnWidth > 46 Then cfg.Columns(7).ColumnWidth = 46
    If cfg.Columns(10).ColumnWidth > 44 Then cfg.Columns(10).ColumnWidth = 44
    cfg.Activate: cfg.Range("A1").Select

    MsgBox "「" & LNK_SHEET & "」を作りました。" & vbCrLf & vbCrLf & _
           "対象シート : " & ws.Name & vbCrLf & _
           "リンク群   : " & grp.Count & " 件（行×系統でまとめた数）" & vbCrLf & _
           "  展開可   : " & nExp & " 件 … 全ての口径列に数式を入れます" & vbCrLf & _
           "  展開不可 : " & nFix & " 件 … 今ある列だけに入れます" & vbCrLf & _
           "  対象外   : " & nSkip & " 件 … 列ごとに参照先が違うため手作業のまま" & vbCrLf & vbCrLf & _
           "口径対応表の空欄と、★の付いた推定値を確かめてから" & vbCrLf & _
           "「直接リンクで数式を作る」を実行してください。", _
           vbInformation, "取り込み完了"
End Sub

'==================================================================
' 2) 数式を書き込む
'==================================================================
Public Sub 直接リンクで数式を作る()
    WriteFormulas False
End Sub

Public Sub INDIRECTで数式を作る()
    WriteFormulas True
End Sub

Private Sub WriteFormulas(ByVal useIndirect As Boolean)
    Dim cfg As Worksheet, ws As Worksheet
    Dim r As Long, lastRow As Long, i As Long
    Dim nWrite As Long, nBlank As Long, nErr As Long
    Dim scr As Boolean, calc As XlCalculation
    Dim diaOf As Object, keiOf As Object, colsOfKei As Object

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

    ' 口径対応表を読む
    Set diaOf = CreateObject("Scripting.Dictionary")      ' 列 → 口径
    Set keiOf = CreateObject("Scripting.Dictionary")      ' 列 → 系統
    Set colsOfKei = CreateObject("Scripting.Dictionary")  ' 系統 → 列の並び
    For r = DIA_FIRST To DIA_LAST
        Dim cl As String: cl = Trim$(CStr(cfg.Cells(r, 1).Value))
        If Len(cl) = 0 Then GoTo NextDia
        Dim dv As String: dv = Trim$(CStr(cfg.Cells(r, 3).Value))
        Dim kv As String: kv = Trim$(CStr(cfg.Cells(r, 4).Value))
        diaOf(cl) = dv
        keiOf(cl) = kv
        If Len(dv) > 0 Then
            If Not colsOfKei.Exists(kv) Then colsOfKei(kv) = ""
            colsOfKei(kv) = IIf(Len(colsOfKei(kv)) = 0, cl, colsOfKei(kv) & "," & cl)
        End If
NextDia:
    Next r

    If MsgBox(ws.Name & " に数式を書き込みます（" & _
              IIf(useIndirect, "INDIRECT 形式", "直接リンク形式") & "）。" & vbCrLf & vbCrLf & _
              "書き込む前にバックアップシートを作ります。続けますか？", _
              vbYesNo + vbQuestion, "確認") <> vbYes Then Exit Sub

    BackupSheet ws

    scr = Application.ScreenUpdating: calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastRow = cfg.Cells(cfg.Rows.Count, 4).End(xlUp).Row
    For r = LNK_FIRST To lastRow
        If Len(Norm(cfg.Cells(r, 2).Value)) = 0 Then GoTo NextRow      ' 有効でない
        Dim tRow As Long: tRow = Val(cfg.Cells(r, 4).Value)
        Dim kei As String: kei = Trim$(CStr(cfg.Cells(r, 5).Value))
        Dim tmpl As String: tmpl = Trim$(CStr(cfg.Cells(r, 7).Value))
        Dim expand As Boolean: expand = (Len(Norm(cfg.Cells(r, 3).Value)) > 0)
        If tRow = 0 Or Len(tmpl) = 0 Then GoTo NextRow

        Dim targets As String
        If expand And colsOfKei.Exists(kei) Then
            targets = colsOfKei(kei)                    ' 系統の全ての口径列
        Else
            targets = Trim$(CStr(cfg.Cells(r, 6).Value))  ' 今ある列だけ
        End If

        Dim tc As Variant, wrote As Long, note As String
        note = ""
        For Each tc In Split(targets, ",")
            Dim colL As String: colL = Trim$(CStr(tc))
            If Len(colL) = 0 Then GoTo NextCol
            Dim dia As String: dia = ""
            If diaOf.Exists(colL) Then dia = diaOf(colL)
            If Len(dia) = 0 Then
                nBlank = nBlank + 1
                note = note & colL & "列は口径が空欄のため未記入。 "
                GoTo NextCol
            End If

            Dim newF As String
            If useIndirect Then
                newF = BuildIndirect(tmpl, "$C$" & DiaRowOf(cfg, colL))
            Else
                newF = Replace(tmpl, "#", dia)
            End If

            On Error Resume Next
            ws.Cells(tRow, ColToNum(colL)).Formula = newF
            If Err.Number <> 0 Then
                note = note & colL & "列: " & Err.Description & " "
                Err.Clear
                nErr = nErr + 1
            Else
                wrote = wrote + 1
            End If
            On Error GoTo 0
NextCol:
        Next tc

        nWrite = nWrite + wrote
        cfg.Cells(r, 10).Value = Trim$(note)
NextRow:
    Next r

    Application.Calculation = calc
    Application.ScreenUpdating = scr
    Application.CalculateFull

    Dim msg As String
    msg = nWrite & " 個のセルに数式を入れました。"
    If nBlank > 0 Then msg = msg & vbCrLf & nBlank & " 個は口径が空欄のため飛ばしました。"
    If nErr > 0 Then msg = msg & vbCrLf & nErr & " 個は書き込めませんでした（備考欄を確認）。"
    If useIndirect Then
        msg = msg & vbCrLf & vbCrLf & _
              "以後、口径を変えるときは口径対応表の黄色いセルを" & vbCrLf & _
              "書き換えるだけで繋ぎ変わります。"
    Else
        msg = msg & vbCrLf & vbCrLf & _
              "口径を変えたら、口径対応表を直してから" & vbCrLf & _
              "このマクロをもう一度実行してください。"
    End If
    MsgBox msg, vbInformation, "完了"
End Sub

'------------------------------------------------------------------
' テンプレート → INDIRECT 数式
'   ='試掘（舗#'!P4
'     → =IFERROR(INDIRECT("'試掘（舗"&'リンク設定'!$C$7&"'!P4"),0)
'------------------------------------------------------------------
Private Function BuildIndirect(ByVal tmpl As String, ByVal diaCell As String) As String
    Dim re As Object, ms As Object, m As Object
    Dim out As String, pos As Long, rep As String

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "'([^']+)#'!(\$?[A-Z]{1,3}\$?[0-9]+)"

    Set ms = re.Execute(tmpl)
    If ms.Count = 0 Then BuildIndirect = Replace(tmpl, "#", ""): Exit Function

    pos = 1
    For Each m In ms
        rep = "IFERROR(INDIRECT(""'" & m.SubMatches(0) & """&" & _
              "'" & LNK_SHEET & "'!" & diaCell & "&""'!" & m.SubMatches(1) & """),0)"
        out = out & Mid$(tmpl, pos, m.FirstIndex + 1 - pos) & rep
        pos = m.FirstIndex + 1 + m.Length
    Next m
    BuildIndirect = out & Mid$(tmpl, pos)
End Function

'==================================================================
' 3) 点検
'==================================================================
Public Sub リンクを点検()
    Dim cfg As Worksheet, ws As Worksheet, rep As Worksheet
    Dim r As Long, lastRow As Long, rr As Long
    Dim used As Object, diaOf As Object
    Dim nEmpty As Long, nMiss As Long, nOrphan As Long

    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then
        MsgBox "先に「リンク設定を取り込む」を実行してください。", vbExclamation
        Exit Sub
    End If
    Set ws = FindSheet(ThisWorkbook, CStr(cfg.Range("B2").Value))
    If ws Is Nothing Then MsgBox "対象シートがありません。", vbExclamation: Exit Sub

    Set used = CreateObject("Scripting.Dictionary")
    Set diaOf = CreateObject("Scripting.Dictionary")
    For r = DIA_FIRST To DIA_LAST
        Dim cl As String: cl = Trim$(CStr(cfg.Cells(r, 1).Value))
        If Len(cl) > 0 Then diaOf(cl) = Trim$(CStr(cfg.Cells(r, 3).Value))
    Next r

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("リンク点検").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set rep = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    rep.Name = "リンク点検"
    rep.Range("A1").Value = "リンク点検　" & Format$(Now, "yyyy/mm/dd hh:nn") & "　対象: " & ws.Name
    rep.Range("A1").Font.Bold = True
    WriteRow rep, 3, Array("対象行", "列", "参照先シート", "参照セル", "値", "所見")
    StyleHeader rep.Range("A3:F3")

    rr = 4
    lastRow = cfg.Cells(cfg.Rows.Count, 4).End(xlUp).Row
    For r = LNK_FIRST To lastRow
        Dim tRow As Long: tRow = Val(cfg.Cells(r, 4).Value)
        Dim tmpl As String: tmpl = Trim$(CStr(cfg.Cells(r, 7).Value))
        If tRow = 0 Or Len(tmpl) = 0 Then GoTo NextRow

        Dim tc As Variant
        For Each tc In Split(CStr(cfg.Cells(r, 6).Value), ",")
            Dim colL As String: colL = Trim$(CStr(tc))
            If Len(colL) = 0 Then GoTo NextCol
            Dim dia As String: dia = ""
            If diaOf.Exists(colL) Then dia = diaOf(colL)
            If Len(dia) = 0 Then GoTo NextCol

            Dim terms As Collection: Set terms = ParseTerms(Replace(tmpl, "#", dia))
            Dim i As Long
            For i = 1 To terms.Count
                Dim sn As String: sn = terms(i)(0)
                used(sn) = True
                rep.Cells(rr, 1).Value = tRow
                rep.Cells(rr, 2).Value = colL
                rep.Cells(rr, 3).Value = sn
                rep.Cells(rr, 4).Value = terms(i)(1)

                Dim tgt As Worksheet: Set tgt = FindSheet(ThisWorkbook, sn)
                If tgt Is Nothing Then
                    rep.Cells(rr, 6).Value = "★シートがありません"
                    rep.Cells(rr, 6).Interior.Color = RGB(255, 199, 206)
                    nMiss = nMiss + 1
                Else
                    Dim v As Variant
                    v = NumOrEmpty(tgt.Range(Replace(terms(i)(1), "$", "")))
                    If Not IsEmpty(v) Then rep.Cells(rr, 5).Value = v
                    If IsEmpty(v) Or v = 0 Then
                        rep.Cells(rr, 6).Value = "参照先が空（この口径は未使用かもしれません）"
                        rep.Cells(rr, 6).Interior.Color = RGB(255, 235, 156)
                        nEmpty = nEmpty + 1
                    Else
                        rep.Cells(rr, 6).Value = "OK"
                    End If
                End If
                rr = rr + 1
            Next i
NextCol:
        Next tc
NextRow:
    Next r

    rr = rr + 1
    rep.Cells(rr, 1).Value = "【リンクされていない口径付きシート】"
    rep.Cells(rr, 1).Font.Bold = True
    rr = rr + 1
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If used.Exists(sh.Name) Then GoTo NextSheet
        If Len(TrailDigits(sh.Name)) = 0 Then GoTo NextSheet
        Dim anyData As Boolean
        anyData = (Application.WorksheetFunction.Count(sh.UsedRange) > 0) And _
                  (Application.WorksheetFunction.Sum(sh.UsedRange) <> 0)
        rep.Cells(rr, 1).Value = sh.Name
        rep.Cells(rr, 2).Value = IIf(sh.Visible = xlSheetVisible, "表示", "非表示")
        If anyData Then
            rep.Cells(rr, 6).Value = "★中身があるのにリンクされていません。張り忘れの可能性"
            rep.Cells(rr, 6).Interior.Color = RGB(255, 199, 206)
            nOrphan = nOrphan + 1
        Else
            rep.Cells(rr, 6).Value = "空。未使用の雛形とみられます"
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
Private Sub AddCol(ByRef cols() As TCol, ByRef n As Long, ByVal letter As String, _
                   ByVal header As String, ByVal kei As String, ByVal dia As String)
    Dim i As Long
    For i = 0 To n - 1
        If cols(i).Letter = letter Then
            If Len(cols(i).Dia) = 0 Then cols(i).Dia = dia
            Exit Sub
        End If
    Next i
    If n > UBound(cols) Then ReDim Preserve cols(0 To n + 20)
    cols(n).Letter = letter
    cols(n).Header = header
    cols(n).Kei = kei
    cols(n).Dia = dia
    cols(n).Guessed = False
    n = n + 1
End Sub

' リンクのある列と同じ系統の列を、口径が空のまま加える
Private Sub AddSiblingCols(ByVal ws As Worksheet, ByRef cols() As TCol, ByRef n As Long)
    Dim keis As Object: Set keis = CreateObject("Scripting.Dictionary")
    Dim i As Long, c As Long
    For i = 0 To n - 1
        If Len(cols(i).Kei) > 0 Then keis(cols(i).Kei) = True
    Next i

    For c = 1 To LastUsedCol(ws)
        Dim kei As String: kei = KeiOf(ws, c)
        If Len(kei) = 0 Then GoTo NextCol
        If Not keis.Exists(kei) Then GoTo NextCol
        Dim letter As String: letter = ColLetterFromNum(c)
        For i = 0 To n - 1
            If cols(i).Letter = letter Then GoTo NextCol
        Next i
        AddCol cols, n, letter, HeaderOf(ws, c), kei, ""
NextCol:
    Next c
End Sub

' 口径が空の列に、シート名と見出しの数字から候補を当てる
Private Sub GuessDiameters(ByRef cols() As TCol, ByVal n As Long)
    Dim pass As Long, changed As Boolean, i As Long

    For pass = 1 To 5
        changed = False
        For i = 0 To n - 1
            If Len(cols(i).Dia) > 0 Then GoTo NextCol
            Dim cand As String: cand = ""
            Dim cnt As Long: cnt = 0
            Dim nums As Variant: nums = NumbersIn(cols(i).Header)
            Dim v As Variant
            For Each v In nums
                If SheetExistsForKei(cols(i).Kei, CStr(v)) Then
                    If Not DiaUsed(cols, n, cols(i).Kei, CStr(v)) Then
                        cand = CStr(v): cnt = cnt + 1
                    End If
                End If
            Next v
            If cnt = 1 Then
                cols(i).Dia = cand
                cols(i).Guessed = True
                changed = True
            End If
NextCol:
        Next i
        If Not changed Then Exit For
    Next pass
End Sub

Private Function DiaUsed(ByRef cols() As TCol, ByVal n As Long, _
                         ByVal kei As String, ByVal dia As String) As Boolean
    Dim i As Long
    For i = 0 To n - 1
        If cols(i).Kei = kei And cols(i).Dia = dia Then DiaUsed = True: Exit Function
    Next i
End Function

' その系統の接頭辞 + 口径 のシートがブックにあるか
Private Function SheetExistsForKei(ByVal kei As String, ByVal dia As String) As Boolean
    Dim sh As Worksheet, k As String
    k = Norm(kei)
    If Len(k) = 0 Then Exit Function
    For Each sh In ThisWorkbook.Worksheets
        If TrailDigits(sh.Name) = dia Then
            If InStr(1, Norm(sh.Name), Left$(k, 2), vbTextCompare) > 0 Then
                SheetExistsForKei = True
                Exit Function
            End If
        End If
    Next sh
End Function

Private Function NumbersIn(ByVal s As String) As Variant
    Dim re As Object, ms As Object, m As Object, out() As String, n As Long
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[0-9]+"
    Set ms = re.Execute(StrConv(s, vbNarrow))
    ReDim out(0 To ms.Count)
    For Each m In ms
        out(n) = m.Value: n = n + 1
    Next m
    If n = 0 Then
        NumbersIn = Array()
    Else
        ReDim Preserve out(0 To n - 1)
        NumbersIn = out
    End If
End Function

' 見出しの「（」より前を系統とみなす
Private Function KeiOf(ByVal ws As Worksheet, ByVal col As Long) As String
    Dim h As String, p As Long
    h = HeaderOf(ws, col)
    If Len(h) = 0 Then Exit Function
    p = InStr(h, ChrW(&HFF08))          ' （
    If p = 0 Then p = InStr(h, "(")
    If p > 0 Then h = Left$(h, p - 1)
    KeiOf = Trim$(Replace(Replace(h, " ", ""), ChrW(&H3000), ""))
End Function

Private Function SumOfCols(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As String) As Variant
    ' any は VBA の予約語なので使えない
    Dim tc As Variant, tot As Double, found As Boolean
    For Each tc In Split(cols, ",")
        Dim v As Variant
        v = NumOrEmpty(ws.Cells(r, ColToNum(Trim$(CStr(tc)))))
        If Not IsEmpty(v) Then
            tot = tot + CDbl(v)
            found = True
        End If
    Next tc
    If found Then SumOfCols = tot Else SumOfCols = Empty
End Function

Private Function DiaRowOf(ByVal cfg As Worksheet, ByVal colL As String) As Long
    Dim r As Long
    For r = DIA_FIRST To DIA_LAST
        If Norm(cfg.Cells(r, 1).Value) = Norm(colL) Then DiaRowOf = r: Exit Function
    Next r
End Function

Private Function SortedGroupKeys(ByVal d As Object) As Variant
    Dim k As Variant, arr() As String, n As Long, i As Long, j As Long, t As String
    ReDim arr(0 To d.Count - 1)
    For Each k In d.Keys
        arr(n) = CStr(k): n = n + 1
    Next k
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            Dim a As Long, b As Long
            a = CLng(Split(arr(i), "|")(0)): b = CLng(Split(arr(j), "|")(0))
            If b < a Or (b = a And Split(arr(j), "|")(1) < Split(arr(i), "|")(1)) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    SortedGroupKeys = arr
End Function

Private Function AskTargetSheet() As String
    AskTargetSheet = Trim$(InputBox( _
        "リンクを取り込む総括表シートの名前を入れてください。" & vbCrLf & vbCrLf & _
        "例: 総括表（土工事）", "対象シート", "総括表（土工事）"))
End Function

Private Function HasSheetRef(ByVal c As Range) As Boolean
    Dim f As String
    If Not c.HasFormula Then Exit Function
    f = c.Formula
    If InStr(f, "表紙") > 0 Then Exit Function
    ' INDIRECT 形式は「リンク設定」への参照しか持たないので、そちらも拾う
    HasSheetRef = (InStr(f, "!") > 0) Or (InStr(f, "INDIRECT") > 0)
End Function

'------------------------------------------------------------------
' 前回このマクロが入れた INDIRECT 形式を、普通のシート参照に戻す
'   =IFERROR(INDIRECT("'試掘（舗"&'リンク設定'!$C$7&"'!P4"),0)
'     → ='試掘（舗400'!P4
' 口径は、まだ残っている「リンク設定」シートから読む。
' これにより、一度変換した後でも取り込み直せる。
'------------------------------------------------------------------
Private Function Normalize(ByVal f As String) As String
    Dim re As Object, ms As Object, m As Object
    Dim cfg As Worksheet, out As String, pos As Long, dia As String

    Normalize = f
    If InStr(f, "INDIRECT") = 0 Then Exit Function

    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then Exit Function     ' 口径が分からないので触らない

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "IFERROR\(INDIRECT\(""'([^""]+)""&'?" & LNK_SHEET & _
                 "'?!(\$?[A-Z]{1,3}\$?[0-9]+)&""'!(\$?[A-Z]{1,3}\$?[0-9]+)""\),0\)"

    Set ms = re.Execute(f)
    If ms.Count = 0 Then Exit Function

    pos = 1
    For Each m In ms
        On Error Resume Next
        dia = Trim$(CStr(cfg.Range(m.SubMatches(1)).Value))
        If Err.Number <> 0 Then Err.Clear: dia = ""
        On Error GoTo 0
        If Len(dia) = 0 Then Exit Function   ' 1つでも解決できなければ諦める
        out = out & Mid$(f, pos, m.FirstIndex + 1 - pos) & _
              "'" & m.SubMatches(0) & dia & "'!" & m.SubMatches(2)
        pos = m.FirstIndex + 1 + m.Length
    Next m
    Normalize = out & Mid$(f, pos)
End Function

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

Private Function TrailDigits(ByVal s As String) As String
    Dim i As Long, ch As String, buf As String
    For i = Len(s) To 1 Step -1
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then buf = ch & buf Else Exit For
    Next i
    TrailDigits = buf
End Function

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

Private Function ColLetterFromNum(ByVal n As Long) As String
    ColLetterFromNum = Split(Cells(1, n).Address(True, False), "$")(0)
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

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If Err.Number <> 0 Or LastUsedCol < 1 Then LastUsedCol = 1
    On Error GoTo 0
End Function

Private Sub WriteRow(ByVal ws As Worksheet, ByVal r As Long, ByVal arr As Variant)
    Dim i As Long
    For i = 0 To UBound(arr)
        ws.Cells(r, i + 1).Value = arr(i)
    Next i
End Sub

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
