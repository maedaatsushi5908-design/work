Attribute VB_Name = "M_Link"
'==================================================================
' M_Link - 総括表のシート間リンクを数式にする
'
' マクロは1本だけ。
'
'     総括表の数式を作り直す()
'
' 中でやること:
'   1. 今あるリンクを読み取って「リンク設定」シートを作る
'      （前回の設定があれば、口径などの手直しは引き継ぐ）
'   2. 何をするかを表示して確認を取る
'   3. バックアップを作ってから数式を書き込む
'   4. 値が変わったセルを一覧にして表示し、
'      納得できなければその場で元に戻す
'
' 数式は2種類。取り込み時にどちらかを自動で割り当てる。
'
'   直接   ='試掘（舗400'!P4
'          舗装切断工のように、転記元の行が固定文字で並ぶ箇所。
'
'   条件式 =SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)
'          舗装版破砕工のブロックは、その工事で出てくる舗装厚だけが
'          詰めて並ぶ。同じセルを見ていると別の厚さの数量を拾うので、
'          総括表の厚さ欄と一致する行を探して合計する。
'==================================================================
Option Explicit

Public Const LNK_SHEET As String = "リンク設定"
Public Const REP_SHEET As String = "実行結果"

Private Const DIA_HDR As Long = 6
Private Const DIA_FIRST As Long = 7
Private Const DIA_LAST As Long = 26
Private Const LNK_HDR As Long = 29
Private Const LNK_FIRST As Long = 30

Private Const TERM_PAT As String = "'([^']+)'!(\$?[A-Z]{1,3}\$?[0-9]+)"

' 埋められなかった黄色いセル。実行結果シートに出す。
Private mUnfilled As Collection

' 列1つ分
Private Type TCol
    Letter  As String
    Header  As String
    Kei     As String
    Dia     As String
    Guessed As Boolean
End Type

' 書き換えた1セル分（元に戻すために控える）
Private Type TChange
    Row_    As Long
    Col_    As Long
    OldF    As String
    OldV    As Variant
    NewF    As String
    NewV    As Variant
End Type

'==================================================================
' 唯一の入口
'==================================================================
Public Sub 総括表の数式を作り直す()
    Dim ws As Worksheet, cfg As Worksheet
    Dim srcName As String, msg As String
    Dim nExp As Long, nFix As Long, nSkip As Long, nCond As Long, nGuess As Long
    Dim changes() As TChange, nChg As Long, nDiff As Long, nLeft As Long
    Dim scr As Boolean, calc As XlCalculation

    ' --- 対象シートを決める ------------------------------------------
    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then
        srcName = Trim$(InputBox( _
            "数式を作り直す総括表シートの名前を入れてください。" & vbCrLf & vbCrLf & _
            "例: 総括表（土工事）", "対象シート", "総括表（土工事）"))
        If Len(srcName) = 0 Then Exit Sub
    Else
        srcName = Trim$(CStr(cfg.Range("B2").Value))
    End If

    Set ws = FindSheet(ThisWorkbook, srcName)
    If ws Is Nothing Then
        MsgBox "シートが見つかりません: " & srcName, vbExclamation
        Exit Sub
    End If

    scr = Application.ScreenUpdating
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- 1. 設定を作る（前回の手直しは引き継ぐ）----------------------
    On Error GoTo Failed
    BuildConfig ws, nExp, nFix, nSkip, nCond, nGuess
    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then GoTo Failed

    Application.Calculation = calc
    Application.ScreenUpdating = scr

    ' --- 2. 何をするかを見せて確認 ------------------------------------
    msg = "対象シート : " & ws.Name & vbCrLf & vbCrLf & _
          "リンク群   : " & (nExp + nFix + nSkip) & " 件" & vbCrLf & _
          "  全口径列に広げる : " & nExp & " 件" & vbCrLf & _
          "  今ある列だけ     : " & nFix & " 件" & vbCrLf & _
          "  手作業のまま     : " & nSkip & " 件" & vbCrLf & vbCrLf & _
          "  うち条件式       : " & nCond & " 件（舗装厚で照合）" & vbCrLf & vbCrLf
    If nGuess > 0 Then
        msg = msg & "※ 口径を " & nGuess & " 件、見出しから推定しました。" & vbCrLf & _
                    "　 実行後に「" & LNK_SHEET & "」シートの★印を確かめてください。" & vbCrLf & vbCrLf
    End If
    msg = msg & "バックアップを作ってから書き込みます。続けますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "確認") <> vbYes Then
        MsgBox "書き込みは行いませんでした。" & vbCrLf & _
               "「" & LNK_SHEET & "」シートに、何をする予定だったかが残っています。", _
               vbInformation
        Exit Sub
    End If

    ' --- 3. バックアップ → 書き込み ------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    BackupSheet ws
    nChg = WriteAll(cfg, ws, changes, nLeft)

    Application.Calculation = calc      ' 元の計算モードに戻す
    Application.CalculateFull

    nDiff = CountDiff(changes, nChg)
    WriteReport ws, changes, nChg

    Application.ScreenUpdating = scr

    ' --- 4. 変わったところを見せて、取り消せるようにする ----------------
    If nDiff = 0 Then
        DropConfigIfClean nLeft
        MsgBox nChg & " 個のセルに数式を入れました。" & vbCrLf & _
               "値が変わったセルはありません。" & vbCrLf & vbCrLf & _
               "詳しくは「" & REP_SHEET & "」シートを見てください。", _
               vbInformation, "完了"
    Else
        If MsgBox(nChg & " 個のセルに数式を入れました。" & vbCrLf & _
                  "そのうち " & nDiff & " 個で値が変わっています。" & vbCrLf & vbCrLf & _
                  "「" & REP_SHEET & "」シートに変わったセルの一覧があります。" & vbCrLf & _
                  "内容を確かめてから答えてください。" & vbCrLf & vbCrLf & _
                  "このまま確定しますか？" & vbCrLf & _
                  "「いいえ」を選ぶと、すべて元に戻します。", _
                  vbYesNo + vbQuestion, "値が変わりました") <> vbYes Then
            Application.ScreenUpdating = False
            UndoAll ws, changes, nChg
            Application.ScreenUpdating = scr
            Application.CalculateFull
            MsgBox "元に戻しました。" & vbCrLf & _
                   "「" & LNK_SHEET & "」シートで設定を直してから、" & vbCrLf & _
                   "もう一度実行してください。", vbInformation, "取り消し"
            Exit Sub
        End If
        DropConfigIfClean nLeft
        MsgBox "確定しました。" & vbCrLf & _
               "元の状態は BK_ で始まるシートに残っています。", vbInformation, "完了"
    End If
    Exit Sub

Failed:
    Application.Calculation = calc
    Application.ScreenUpdating = scr
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical
End Sub

'==================================================================
' 1. 設定を作る
'==================================================================
Private Sub BuildConfig(ByVal ws As Worksheet, ByRef nExp As Long, ByRef nFix As Long, _
                        ByRef nSkip As Long, ByRef nCond As Long, ByRef nGuess As Long)
    Dim cfg As Worksheet, c As Range, f As String
    Dim cols() As TCol, nCol As Long
    Dim grp As Object, keys As Object, keep As Object, keepDia As Object
    Dim i As Long, r As Long, n As Long
    Dim thkCol As String, kndCol As String

    ' 前回の手直しを控えておく
    Set keep = CreateObject("Scripting.Dictionary")
    Set keepDia = CreateObject("Scripting.Dictionary")
    RememberEdits keep, keepDia, thkCol, kndCol

    If Len(thkCol) = 0 Then thkCol = FindLabelCol(ws, "舗装厚", "I")
    If Len(kndCol) = 0 Then kndCol = FindLabelCol(ws, "摘", "E")

    Set grp = CreateObject("Scripting.Dictionary")
    Set keys = CreateObject("Scripting.Dictionary")
    ReDim cols(0 To 60)

    ' --- 全リンクを走査 ---------------------------------------------
    For Each c In ws.UsedRange
        If Not HasSheetRef(c) Then GoTo NextCell
        f = Normalize(c.Formula)

        Dim terms As Collection: Set terms = ParseTerms(f)
        If terms.Count = 0 Then GoTo NextCell
        If terms.Count <> CountChar(f, "!") Then GoTo NextCell

        Dim colL As String: colL = ColLetterOf(c)
        Dim kei As String: kei = KeiOf(ws, c.Column)

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
                bad = True: Exit For
            End If
            sig = sig & "|" & Left$(pre, Len(pre) - Len(d)) & "!" & addr
        Next i
        If bad Then GoTo NextCell

        AddCol cols, nCol, colL, HeaderOf(ws, c.Column), kei, dia

        Dim p As Variant
        For Each p In Split(Mid$(sig, 2), "|")
            Dim kk As String: kk = kei & "|" & p
            If Not keys.Exists(kk) Then keys(kk) = ""
            If InStr(1, "," & keys(kk) & ",", "," & c.Row & ",") = 0 Then
                keys(kk) = IIf(Len(keys(kk)) = 0, CStr(c.Row), keys(kk) & "," & c.Row)
            End If
        Next p

        Dim gk As String: gk = c.Row & "|" & kei
        If grp.Exists(gk) Then
            Dim cur As Variant: cur = grp(gk)
            If cur(1) <> sig Then cur(3) = "NG_SIG"
            cur(0) = cur(0) & "," & colL
            grp(gk) = cur
        Else
            grp(gk) = Array(colL, sig, MakeTemplate(f, dia), "", "")
        End If
NextCell:
    Next c

    If grp.Count = 0 Then Err.Raise 5, , "口径の付いたシートを参照するリンクが見つかりませんでした。"

    ' --- 展開の可否 -------------------------------------------------
    Dim gkv As Variant
    For Each gkv In grp.Keys
        Dim g As Variant: g = grp(gkv)
        If g(3) = "NG_SIG" Then
            g(4) = "列ごとに参照セルが違う（シートの行数が口径で異なる）"
        Else
            Dim shared_ As String: shared_ = ""
            Dim kei2 As String: kei2 = Split(gkv, "|")(1)
            For Each p In Split(Mid$(g(1), 2), "|")
                Dim q As Variant
                For Each q In Split(keys(kei2 & "|" & p), ",")
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

    AddSiblingCols ws, cols, nCol
    ApplyKeptDia cols, nCol, keepDia
    GuessDiameters cols, nCol, nGuess

    ' --- 設定シートを作る --------------------------------------------
    Set cfg = ResetConfigSheet()
    cfg.Range("A2").Value = "対象シート"
    cfg.Range("B2").Value = ws.Name
    cfg.Range("A3").Value = "厚さ列"
    cfg.Range("B3").Value = thkCol
    cfg.Range("C3").Value = "種別列"
    cfg.Range("D3").Value = kndCol
    cfg.Range("B3").Interior.Color = RGB(255, 242, 204)
    cfg.Range("D3").Interior.Color = RGB(255, 242, 204)
    cfg.Range("A4").Value = "この表を書き換えてから、もう一度マクロを実行すると反映されます"
    cfg.Range("A4").Font.Italic = True

    cfg.Range("A5").Value = "【口径対応表】　黄色いセルを書き換えると、その列の参照先が変わります。空欄の列には数式を入れません"
    cfg.Range("A5").Font.Bold = True
    WriteRow cfg, DIA_HDR, Array("列", "総括表の見出し", "口径", "系統", "備考")
    StyleHeader cfg.Range(cfg.Cells(DIA_HDR, 1), cfg.Cells(DIA_HDR, 5))

    r = DIA_FIRST
    For i = 0 To nCol - 1
        If r > DIA_LAST Then Exit For
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

    cfg.Cells(LNK_HDR - 1, 1).Value = "【リンク一覧】　有効を空欄にするとその行は書き換えません"
    cfg.Cells(LNK_HDR - 1, 1).Font.Bold = True
    WriteRow cfg, LNK_HDR, Array("No", "有効", "展開", "方式", "対象行", "厚さ", "種別", "系統", _
                                 "今ある列", "テンプレート", "実行前の値", "判定", "備考")
    StyleHeader cfg.Range(cfg.Cells(LNK_HDR, 1), cfg.Cells(LNK_HDR, 13))

    r = LNK_FIRST
    Dim sk As Variant
    For Each sk In SortedGroupKeys(grp)
        Dim gg As Variant: gg = grp(sk)
        Dim gr As Long: gr = CLng(Split(sk, "|")(0))
        n = n + 1

        Dim thk As Variant, knd As String, firstCol As String, firstDia As String
        thk = ws.Cells(gr, ColToNum(thkCol)).Value
        knd = KindOf(ws, gr, ColToNum(kndCol))
        firstCol = Split(CStr(gg(0)), ",")(0)
        firstDia = DiaOfCol(cols, nCol, firstCol)

        Dim way As String
        way = "直接"
        If IsNumeric(thk) And Len(firstDia) > 0 And Len(knd) > 0 Then
            If HasBreakBlock(PrefixOf(CStr(gg(2))) & firstDia) Then way = "条件式"
        End If

        cfg.Cells(r, 1).Value = n
        cfg.Cells(r, 2).Value = IIf(gg(3) = "NG_SIG", "", "○")
        cfg.Cells(r, 4).Value = way
        cfg.Cells(r, 5).Value = gr
        If IsNumeric(thk) Then cfg.Cells(r, 6).Value = thk
        cfg.Cells(r, 7).Value = knd
        cfg.Cells(r, 8).Value = Split(sk, "|")(1)
        cfg.Cells(r, 9).Value = gg(0)
        cfg.Cells(r, 10).Value = "'" & gg(2)
        cfg.Cells(r, 11).Value = SumOfCols(ws, gr, CStr(gg(0)))

        If gg(3) = "OK" Then
            cfg.Cells(r, 3).Value = "○"
            cfg.Cells(r, 12).Value = "展開可"
            nExp = nExp + 1
        ElseIf gg(3) = "NG_SIG" Then
            cfg.Cells(r, 3).Value = ""
            cfg.Cells(r, 12).Value = "対象外"
            cfg.Cells(r, 12).Interior.Color = RGB(255, 199, 206)
            cfg.Cells(r, 13).Value = gg(4) & "。書き込みません（手作業のまま）"
            nSkip = nSkip + 1
        Else
            cfg.Cells(r, 3).Value = ""
            cfg.Cells(r, 12).Value = "展開不可"
            cfg.Cells(r, 12).Interior.Color = RGB(255, 235, 156)
            cfg.Cells(r, 13).Value = gg(4) & "。今ある列だけに入れます"
            nFix = nFix + 1
        End If

        ' 前回の手直しを引き継ぐ
        Dim kk2 As String: kk2 = CStr(sk)
        If keep.Exists(kk2) Then
            Dim kv As Variant: kv = keep(kk2)
            cfg.Cells(r, 2).Value = kv(0)
            cfg.Cells(r, 3).Value = kv(1)
            cfg.Cells(r, 4).Value = kv(2)
            way = CStr(kv(2))
        End If
        If way = "条件式" Then nCond = nCond + 1
        r = r + 1
    Next sk

    cfg.Columns.AutoFit
    If cfg.Columns(10).ColumnWidth > 46 Then cfg.Columns(10).ColumnWidth = 46
    If cfg.Columns(13).ColumnWidth > 44 Then cfg.Columns(13).ColumnWidth = 44
End Sub

' 前回の設定シートから、手直しされうる欄を控える
Private Sub RememberEdits(ByVal keep As Object, ByVal keepDia As Object, _
                          ByRef thkCol As String, ByRef kndCol As String)
    Dim cfg As Worksheet, r As Long
    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then Exit Sub

    thkCol = Trim$(CStr(cfg.Range("B3").Value))
    kndCol = Trim$(CStr(cfg.Range("D3").Value))

    For r = DIA_FIRST To DIA_LAST
        Dim cl As String: cl = Trim$(CStr(cfg.Cells(r, 1).Value))
        If Len(cl) > 0 Then keepDia(cl) = Trim$(CStr(cfg.Cells(r, 3).Value))
    Next r

    Dim last As Long: last = cfg.Cells(cfg.Rows.Count, 5).End(xlUp).Row
    For r = LNK_FIRST To last
        Dim gr As Long: gr = Val(cfg.Cells(r, 5).Value)
        Dim kei As String: kei = Trim$(CStr(cfg.Cells(r, 8).Value))
        If gr > 0 And Len(kei) > 0 Then
            keep(gr & "|" & kei) = Array(cfg.Cells(r, 2).Value, _
                                         cfg.Cells(r, 3).Value, _
                                         cfg.Cells(r, 4).Value)
        End If
    Next r
End Sub

Private Sub ApplyKeptDia(ByRef cols() As TCol, ByVal n As Long, ByVal keptDia As Object)
    Dim i As Long
    For i = 0 To n - 1
        If keptDia.Exists(cols(i).Letter) Then
            Dim v As String: v = CStr(keptDia(cols(i).Letter))
            If Len(v) > 0 Then
                cols(i).Dia = v
                cols(i).Guessed = False
            End If
        End If
    Next i
End Sub

'==================================================================
' 3. 書き込み
'==================================================================
Private Function WriteAll(ByVal cfg As Worksheet, ByVal ws As Worksheet, _
                          ByRef changes() As TChange, ByRef nLeft As Long) As Long
    Dim r As Long, lastRow As Long, n As Long
    Dim diaOf As Object, colsOfKei As Object
    Dim thkColW As String

    thkColW = Trim$(CStr(cfg.Range("B3").Value))
    If Len(thkColW) = 0 Then thkColW = "I"

    Set diaOf = CreateObject("Scripting.Dictionary")
    Set colsOfKei = CreateObject("Scripting.Dictionary")
    For r = DIA_FIRST To DIA_LAST
        Dim cl As String: cl = Trim$(CStr(cfg.Cells(r, 1).Value))
        If Len(cl) = 0 Then GoTo NextDia
        Dim dv As String: dv = Trim$(CStr(cfg.Cells(r, 3).Value))
        Dim kv As String: kv = Trim$(CStr(cfg.Cells(r, 4).Value))
        diaOf(cl) = dv
        If Len(dv) > 0 Then
            If Not colsOfKei.Exists(kv) Then colsOfKei(kv) = ""
            colsOfKei(kv) = IIf(Len(colsOfKei(kv)) = 0, cl, colsOfKei(kv) & "," & cl)
        End If
NextDia:
    Next r

    ReDim changes(0 To 400)

    lastRow = cfg.Cells(cfg.Rows.Count, 5).End(xlUp).Row
    For r = LNK_FIRST To lastRow
        If Len(Norm(cfg.Cells(r, 2).Value)) = 0 Then GoTo NextRow

        Dim way As String: way = Trim$(CStr(cfg.Cells(r, 4).Value))
        Dim tRow As Long: tRow = Val(cfg.Cells(r, 5).Value)
        Dim thick As String: thick = Trim$(CStr(cfg.Cells(r, 6).Value))
        Dim kind As String: kind = Trim$(CStr(cfg.Cells(r, 7).Value))
        Dim kei As String: kei = Trim$(CStr(cfg.Cells(r, 8).Value))
        Dim tmpl As String: tmpl = Trim$(CStr(cfg.Cells(r, 10).Value))
        Dim expand As Boolean: expand = (Len(Norm(cfg.Cells(r, 3).Value)) > 0)
        If tRow = 0 Or Len(tmpl) = 0 Then GoTo NextRow

        Dim targets As String
        If expand And colsOfKei.Exists(kei) Then
            targets = colsOfKei(kei)
        Else
            targets = Trim$(CStr(cfg.Cells(r, 9).Value))
        End If

        Dim tc As Variant, note As String
        note = ""
        For Each tc In Split(targets, ",")
            Dim colL As String: colL = Trim$(CStr(tc))
            If Len(colL) = 0 Then GoTo NextCol
            Dim dia As String: dia = ""
            If diaOf.Exists(colL) Then dia = diaOf(colL)
            If Len(dia) = 0 Then
                note = note & colL & "列は口径が空欄のため未記入。 "
                GoTo NextCol
            End If

            Dim newF As String, why As String
            If way = "条件式" And Len(thick) = 0 Then way = "直接"
            If way = "条件式" Then
                newF = BuildSumif(tmpl, dia, "'" & ws.Name & "'!$" & thkColW & tRow, kind, why)
                If Len(newF) = 0 Then
                    note = note & colL & "列: " & why & " "
                    GoTo NextCol
                End If
            Else
                newF = Replace(tmpl, "#", dia)
            End If

            Dim cell_ As Range
            Set cell_ = ws.Cells(tRow, ColToNum(colL))

            If n > UBound(changes) Then ReDim Preserve changes(0 To n + 200)
            changes(n).Row_ = tRow
            changes(n).Col_ = cell_.Column
            changes(n).OldF = cell_.Formula
            changes(n).OldV = NumOrEmpty(cell_)
            changes(n).NewF = newF

            On Error Resume Next
            cell_.Formula = newF
            If Err.Number <> 0 Then
                note = note & colL & "列: " & Err.Description & " "
                Err.Clear
                On Error GoTo 0
                GoTo NextCol
            End If
            On Error GoTo 0
            n = n + 1
NextCol:
        Next tc

        cfg.Cells(r, 13).Value = Trim$(Trim$(CStr(cfg.Cells(r, 13).Value)) & " " & note)
NextRow:
    Next r

    ' 黄色い「入力セル」も、条件式で埋められるものは埋める
    FillYellow cfg, ws, changes, n, thkColW, diaOf, nLeft

    ' 書き込み後の値を控える
    Application.Calculate
    Dim i As Long
    For i = 0 To n - 1
        changes(i).NewV = NumOrEmpty(ws.Cells(changes(i).Row_, changes(i).Col_))
    Next i

    WriteAll = n
End Function

'------------------------------------------------------------------
' 黄色い「入力セル」を条件式で埋める
'
' 総括表では、各計算書から値を写す場所が黄色く塗ってある（凡例の
' 「入力セル（各数量計算書の数量を転記間違いないように！）」）。
' つまり黄色は「ここに値が入る」という作成者自身の指定なので、
' 推測せずにそこだけを埋める。
'
' 転記元シートは、行の近くにある同じ系統のリンクから接頭辞を借り、
' 口径は列から取る。舗装版破砕は 舗 シート、土量は 土 シートと
' 行によって系統の中でも参照先が変わるため、列だけでは決められない。
'------------------------------------------------------------------
Private Sub FillYellow(ByVal cfg As Worksheet, ByVal ws As Worksheet, _
                       ByRef changes() As TChange, ByRef n As Long, _
                       ByVal thkColW As String, ByVal diaOf As Object, _
                       ByRef nLeft As Long)
    Dim c As Range, r As Long, lastRow As Long
    Dim kndColW As String, kei As String, colL As String
    Dim rows_ As Collection, keis As Collection, pres As Collection

    kndColW = Trim$(CStr(cfg.Range("D3").Value))
    If Len(kndColW) = 0 Then kndColW = "E"

    ' 設定から「行 → 系統 → 接頭辞」を集める
    Set rows_ = New Collection: Set keis = New Collection: Set pres = New Collection
    lastRow = cfg.Cells(cfg.Rows.Count, 5).End(xlUp).Row
    For r = LNK_FIRST To lastRow
        Dim pre As String
        pre = PrefixOf(Trim$(CStr(cfg.Cells(r, 10).Value)))
        If Len(pre) > 0 Then
            rows_.Add Val(cfg.Cells(r, 5).Value)
            keis.Add Trim$(CStr(cfg.Cells(r, 8).Value))
            pres.Add pre
        End If
    Next r
    If rows_.Count = 0 Then Exit Sub

    Set mUnfilled = New Collection

    For Each c In ws.UsedRange
        If Not IsYellow(c) Then GoTo NextCell
        If c.HasFormula Then GoTo NextCell

        colL = Split(c.Address(True, False), "$")(0)

        Dim thk As Variant: thk = ws.Cells(c.Row, ColToNum(thkColW)).Value
        If Not IsNumeric(thk) Then GoTo NextCell
        Dim kind As String: kind = KindOf(ws, c.Row, ColToNum(kndColW))
        If Len(kind) = 0 Then GoTo NextCell

        kei = KeiOf(ws, c.Column)

        ' 転記元シートの決め方は2通り。
        '   口径のある系統 … 近くの同系統リンクから接頭辞を借り、口径を足す
        '   口径の無い系統 … 見出しの系統名でシートを直に探す（仮配管 → 仮配（舗）
        Dim sn As String
        sn = ""
        If diaOf.Exists(colL) Then
            If Len(CStr(diaOf(colL))) > 0 Then
                pre = NearestPrefix(rows_, keis, pres, c.Row, kei)
                If Len(pre) > 0 Then sn = pre & CStr(diaOf(colL))
            End If
        End If
        If Len(sn) = 0 Then sn = SheetByKei(kei)

        If Len(sn) = 0 Then
            mUnfilled.Add Array(c.Address(False, False), kei, _
                "転記元シートを決められません（同じ系統のリンクも、名前の合うシートも無い）")
            nLeft = nLeft + 1
            GoTo NextCell
        End If

        Dim why As String, newF As String
        newF = BuildSumif("='" & sn & "'!A1", "", "'" & ws.Name & "'!$" & thkColW & c.Row, kind, why)
        If Len(newF) = 0 Then
            mUnfilled.Add Array(c.Address(False, False), kei, why)
            nLeft = nLeft + 1
            GoTo NextCell
        End If

        If n > UBound(changes) Then ReDim Preserve changes(0 To n + 200)
        changes(n).Row_ = c.Row
        changes(n).Col_ = c.Column
        changes(n).OldF = c.Formula
        changes(n).OldV = NumOrEmpty(c)
        changes(n).NewF = newF

        On Error Resume Next
        c.Formula = newF
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextCell
        End If
        On Error GoTo 0
        n = n + 1
        If Len(why) > 0 Then mUnfilled.Add Array(c.Address(False, False), kei, why)
NextCell:
    Next c
End Sub

' 純粋な黄色（入力セルの目印）か
Private Function IsYellow(ByVal c As Range) As Boolean
    On Error Resume Next
    If c.Interior.Pattern = xlNone Then Exit Function
    IsYellow = (c.Interior.Color = RGB(255, 255, 0))
    On Error GoTo 0
End Function

' 同じ系統で、行がいちばん近いリンクの接頭辞を返す
Private Function NearestPrefix(ByVal rows_ As Collection, ByVal keis As Collection, _
                               ByVal pres As Collection, ByVal r As Long, _
                               ByVal kei As String) As String
    Dim i As Long, best As Long, bestPre As String
    best = 999999
    For i = 1 To rows_.Count
        If Norm(keis(i)) = Norm(kei) Then
            Dim d As Long: d = Abs(CLng(rows_(i)) - r)
            If d < best Then
                best = d
                bestPre = CStr(pres(i))
            End If
        End If
    Next i
    NearestPrefix = bestPre
End Function

'------------------------------------------------------------------
' 手直しが要らないなら「リンク設定」シートは残さない。
' 毎回いまの数式から作り直せるので、置いておく意味がないため。
' 埋められなかった箇所があるときだけ、手がかりとして残す。
'------------------------------------------------------------------
Private Sub DropConfigIfClean(ByVal nLeft As Long)
    If nLeft > 0 Then Exit Sub
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(LNK_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Function CountDiff(ByRef changes() As TChange, ByVal n As Long) As Long
    Dim i As Long, c As Long
    For i = 0 To n - 1
        If Differs(changes(i)) Then c = c + 1
    Next i
    CountDiff = c
End Function

Private Function Differs(ByRef ch As TChange) As Boolean
    If IsEmpty(ch.OldV) And IsEmpty(ch.NewV) Then Exit Function
    If IsEmpty(ch.OldV) Or IsEmpty(ch.NewV) Then Differs = True: Exit Function
    Differs = (Abs(CDbl(ch.OldV) - CDbl(ch.NewV)) > 0.00000001)
End Function

Private Sub UndoAll(ByVal ws As Worksheet, ByRef changes() As TChange, ByVal n As Long)
    Dim i As Long
    For i = n - 1 To 0 Step -1
        On Error Resume Next
        If Len(changes(i).OldF) = 0 Then
            ws.Cells(changes(i).Row_, changes(i).Col_).ClearContents
        Else
            ws.Cells(changes(i).Row_, changes(i).Col_).Formula = changes(i).OldF
        End If
        On Error GoTo 0
    Next i
End Sub

'==================================================================
' 4. 結果を書き出す（変わったセル ＋ 点検）
'==================================================================
Private Sub WriteReport(ByVal ws As Worksheet, ByRef changes() As TChange, ByVal n As Long)
    Dim rep As Worksheet, i As Long, r As Long
    Dim used As Object, sh As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(REP_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set rep = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    rep.Name = REP_SHEET
    rep.Range("A1").Value = "実行結果　" & Format$(Now, "yyyy/mm/dd hh:nn") & "　対象: " & ws.Name
    rep.Range("A1").Font.Bold = True
    rep.Range("A1").Font.Size = 12

    ' --- 値が変わったセル --------------------------------------------
    rep.Range("A3").Value = "【値が変わったセル】　内容を確かめてください"
    rep.Range("A3").Font.Bold = True
    WriteRow rep, 4, Array("セル", "実行前", "実行後", "差", "入れた数式")
    StyleHeader rep.Range("A4:E4")

    r = 5
    Set used = CreateObject("Scripting.Dictionary")
    For i = 0 To n - 1
        used(SheetOfFormula(changes(i).NewF)) = True
        If Not Differs(changes(i)) Then GoTo NextChange
        rep.Cells(r, 1).Value = ws.Cells(changes(i).Row_, changes(i).Col_).Address(False, False)
        If Not IsEmpty(changes(i).OldV) Then rep.Cells(r, 2).Value = changes(i).OldV
        If Not IsEmpty(changes(i).NewV) Then rep.Cells(r, 3).Value = changes(i).NewV
        If Not IsEmpty(changes(i).OldV) And Not IsEmpty(changes(i).NewV) Then
            rep.Cells(r, 4).Value = CDbl(changes(i).NewV) - CDbl(changes(i).OldV)
        End If
        rep.Cells(r, 5).Value = "'" & changes(i).NewF
        rep.Cells(r, 1).Interior.Color = RGB(255, 235, 156)
        r = r + 1
NextChange:
    Next i
    If r = 5 Then
        rep.Cells(5, 1).Value = "（値が変わったセルはありません）"
        r = 6
    End If

    ' --- 埋められなかった黄色いセル -----------------------------------
    If Not mUnfilled Is Nothing Then
        If mUnfilled.Count > 0 Then
            r = r + 2
            rep.Cells(r, 1).Value = "【入力セル（黄色）で気になったもの】"
            rep.Cells(r, 1).Font.Bold = True
            r = r + 1
            WriteRow rep, r, Array("セル", "系統", "理由")
            StyleHeader rep.Range(rep.Cells(r, 1), rep.Cells(r, 3))
            r = r + 1
            For i = 1 To mUnfilled.Count
                rep.Cells(r, 1).Value = mUnfilled(i)(0)
                rep.Cells(r, 2).Value = mUnfilled(i)(1)
                rep.Cells(r, 3).Value = mUnfilled(i)(2)
                rep.Cells(r, 3).Interior.Color = RGB(255, 235, 156)
                r = r + 1
            Next i
        End If
    End If

    ' --- 点検：中身があるのにリンクされていないシート --------------------
    r = r + 2
    rep.Cells(r, 1).Value = "【リンクされていない口径付きシート】"
    rep.Cells(r, 1).Font.Bold = True
    r = r + 1
    WriteRow rep, r, Array("シート", "表示", "所見")
    StyleHeader rep.Range(rep.Cells(r, 1), rep.Cells(r, 3))
    r = r + 1

    For Each sh In ThisWorkbook.Worksheets
        If used.Exists(sh.Name) Then GoTo NextSheet
        If Len(TrailDigits(sh.Name)) = 0 Then GoTo NextSheet
        Dim hasData As Boolean
        hasData = (Application.WorksheetFunction.Count(sh.UsedRange) > 0) And _
                  (Application.WorksheetFunction.Sum(sh.UsedRange) <> 0)
        rep.Cells(r, 1).Value = sh.Name
        rep.Cells(r, 2).Value = IIf(sh.Visible = xlSheetVisible, "表示", "非表示")
        If hasData Then
            rep.Cells(r, 3).Value = "★中身があるのにリンクされていません。張り忘れの可能性"
            rep.Cells(r, 3).Interior.Color = RGB(255, 199, 206)
        Else
            rep.Cells(r, 3).Value = "空。未使用の雛形とみられます"
        End If
        r = r + 1
NextSheet:
    Next sh

    rep.Columns.AutoFit
    If rep.Columns(5).ColumnWidth > 60 Then rep.Columns(5).ColumnWidth = 60
    rep.Activate
    rep.Range("A5").Select
End Sub

Private Function SheetOfFormula(ByVal f As String) As String
    Dim terms As Collection
    Set terms = ParseTerms(f)
    If terms.Count > 0 Then SheetOfFormula = terms(1)(0)
End Function

'==================================================================
' 舗装版破砕工ブロックの検出と、条件式の組み立て
'
' 転記元の舗装厚は工事ごとに詰めて並ぶ（4cm が無ければ 5cm が先頭に来る）。
' 同じセルを見ていると別の厚さの数量を拾うので、厚さで照合する。
'==================================================================
Private Function FindBlock(ByVal sn As String, ByRef r0 As Long, ByRef r1 As Long, _
                           ByRef asThk As Long, ByRef asSum As Long, _
                           ByRef coThk As Long, ByRef coSum As Long) As Boolean
    Dim ws As Worksheet, r As Long, c As Long, sect As Long, hdr As Long
    Dim kinds As String, totals As String, i As Long

    Set ws = FindSheet(ThisWorkbook, sn)
    If ws Is Nothing Then Exit Function

    For r = 1 To 60
        For c = 1 To 60
            If InStr(1, Norm(ws.Cells(r, c).Value), "□舗装版破砕", vbTextCompare) > 0 Then
                sect = r
                Exit For
            End If
        Next c
        If sect > 0 Then Exit For
    Next r
    If sect = 0 Then Exit Function

    For r = sect + 1 To sect + 3
        kinds = "": totals = ""
        For c = 1 To 60
            Dim v As String: v = Norm(ws.Cells(r, c).Value)
            If v = "種別・舗装厚" Then kinds = kinds & c & ","
            If v = "合計" Then totals = totals & c & ","
        Next c
        If Len(kinds) > 0 And Len(totals) > 0 Then
            hdr = r
            Exit For
        End If
    Next r
    If hdr = 0 Then Exit Function

    Dim kArr As Variant, tArr As Variant
    kArr = Split(Left$(kinds, Len(kinds) - 1), ",")
    tArr = Split(Left$(totals, Len(totals) - 1), ",")

    asThk = 0: asSum = 0: coThk = 0: coSum = 0
    For i = 0 To UBound(kArr)
        Dim kc As Long: kc = CLng(kArr(i))
        Dim sc As Long: sc = NextTotal(tArr, kc)
        If sc = 0 Then GoTo NextPair
        If asSum = 0 Then
            asThk = kc + 1: asSum = sc
        ElseIf coSum = 0 Then
            coThk = kc + 1: coSum = sc
            Exit For
        End If
NextPair:
    Next i
    If asSum = 0 Then Exit Function

    r0 = hdr + 1
    r1 = r0 - 1
    For r = r0 To r0 + 40
        Dim k As String: k = Norm(ws.Cells(r, asThk - 1).Value)
        If k <> "AS" And k <> "CO" Then Exit For
        r1 = r
    Next r
    If r1 < r0 Then Exit Function

    ' 予備を1行足す。次の工事で舗装厚が1種類増えても式を直さずに済む。
    ' SUMIF は空欄に当たらないので、余分に含めても結果は変わらない。
    r1 = r1 + 1

    FindBlock = True
End Function

Private Function NextTotal(ByVal tArr As Variant, ByVal kc As Long) As Long
    Dim i As Long, best As Long
    For i = 0 To UBound(tArr)
        Dim t As Long: t = CLng(tArr(i))
        If t > kc Then
            If best = 0 Or t < best Then best = t
        End If
    Next i
    NextTotal = best
End Function

Private Function HasBreakBlock(ByVal sn As String) As Boolean
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long
    HasBreakBlock = FindBlock(sn, a, b, c, d, e, f)
End Function

Private Function BuildSumif(ByVal tmpl As String, ByVal dia As String, _
                             ByVal thkRef As String, ByVal kind As String, _
                             ByRef why As String) As String
    Dim sn As String, pre As String, out As String
    Dim r0 As Long, r1 As Long, aT As Long, aS As Long, cT As Long, cS As Long

    why = ""
    If Len(dia) = 0 Then
        ' 口径が無い系統は、テンプレートのシート名をそのまま使う
        sn = SheetInFormula(tmpl)
        If Len(sn) = 0 Then why = "シート名を取り出せません": Exit Function
    Else
        pre = PrefixOf(tmpl)
        If Len(pre) = 0 Then why = "シート名を取り出せません": Exit Function
        sn = pre & dia
    End If
    If Not FindBlock(sn, r0, r1, aT, aS, cT, cS) Then
        why = "舗装版破砕工のブロックが見つかりません(" & sn & ")"
        Exit Function
    End If

    If InStr(kind, "As") > 0 And aS > 0 Then out = SumifTerm(sn, aS, aT, r0, r1, thkRef)
    If InStr(kind, "Co") > 0 And cS > 0 Then
        If Len(out) > 0 Then out = out & "+"
        out = out & SumifTerm(sn, cS, cT, r0, r1, thkRef)
    End If

    If Len(out) = 0 Then
        If InStr(kind, "Co") > 0 And cS = 0 Then
            why = sn & " に Co 側の欄が無いため埋められません"
        Else
            why = "種別が As でも Co でもありません(" & kind & ")"
        End If
        Exit Function
    End If
    If InStr(kind, "Co") > 0 And cS = 0 Then
        why = "注意: " & sn & " に Co 側の欄が無いため As だけを合計しています"
    End If
    BuildSumif = "=" & out
End Function

' 系統名に合う転記元シートを1つだけ見つける
'
'   1. 名前がそのまま一致するシート（給水2度 → 給水2度）
'   2. 舗装版破砕のブロックを持つシートのうち、括弧より前が
'      系統名の先頭に一致するもの（仮配管 → 仮配（舗）
'
' 2文字だけで照合すると 給水2度 と 給水付替 が同じシートに当たるため、
' 括弧より前の全体で照合し、1つに絞れたときだけ採用する。
Private Function SheetByKei(ByVal kei As String) As String
    Dim sh As Worksheet, k As String, hit As String, n As Long, tok As String
    k = Norm(kei)
    If Len(k) = 0 Then Exit Function

    For Each sh In ThisWorkbook.Worksheets
        If Norm(sh.Name) = k Then SheetByKei = sh.Name: Exit Function
    Next sh

    For Each sh In ThisWorkbook.Worksheets
        tok = LeadToken(sh.Name)
        If Len(tok) >= 2 Then
            If Left$(k, Len(tok)) = tok Then
                If HasBreakBlock(sh.Name) Then
                    hit = sh.Name
                    n = n + 1
                End If
            End If
        End If
    Next sh
    If n = 1 Then SheetByKei = hit
End Function

' シート名の括弧より前  仮配（舗 → 仮配
Private Function LeadToken(ByVal s As String) As String
    Dim t As String, p As Long
    t = Norm(s)
    p = InStr(t, "(")
    If p = 0 Then p = InStr(t, ChrW(&HFF08))
    If p > 0 Then t = Left$(t, p - 1)
    LeadToken = t
End Function

Private Function SheetInFormula(ByVal f As String) As String
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "'([^']+)'!"
    Set ms = re.Execute(f)
    If ms.Count > 0 Then SheetInFormula = ms(0).SubMatches(0)
End Function

'------------------------------------------------------------------
' 1項ぶんの SUMIF を組み立てる
'   =SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)
'           └ 転記元の舗装厚        └ 総括表の厚さ欄      └ 転記元の合計
'------------------------------------------------------------------
Private Function SumifTerm(ByVal sn As String, ByVal sumCol As Long, ByVal thkCol As Long, _
                           ByVal r0 As Long, ByVal r1 As Long, ByVal thkRef As String) As String
    Dim q As String
    q = "'" & sn & "'!"
    SumifTerm = "SUMIF(" & q & Rng(thkCol, r0, r1) & "," & thkRef & "," & q & Rng(sumCol, r0, r1) & ")"
End Function

Private Function Rng(ByVal col As Long, ByVal r0 As Long, ByVal r1 As Long) As String
    Dim L As String
    L = ColLetterFromNum(col)
    Rng = "$" & L & "$" & r0 & ":$" & L & "$" & r1
End Function

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

Private Sub GuessDiameters(ByRef cols() As TCol, ByVal n As Long, ByRef nGuess As Long)
    Dim pass As Long, changed As Boolean, i As Long
    For pass = 1 To 5
        changed = False
        For i = 0 To n - 1
            If Len(cols(i).Dia) > 0 Then GoTo NextCol
            Dim cand As String: cand = ""
            Dim cnt As Long: cnt = 0
            Dim v As Variant
            For Each v In NumbersIn(cols(i).Header)
                If SheetExistsForKei(cols(i).Kei, CStr(v)) Then
                    If Not DiaUsed(cols, n, cols(i).Kei, CStr(v)) Then
                        cand = CStr(v): cnt = cnt + 1
                    End If
                End If
            Next v
            If cnt = 1 Then
                cols(i).Dia = cand
                cols(i).Guessed = True
                nGuess = nGuess + 1
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
    If ms.Count = 0 Then NumbersIn = Array(): Exit Function
    ReDim out(0 To ms.Count - 1)
    For Each m In ms
        out(n) = m.Value: n = n + 1
    Next m
    NumbersIn = out
End Function

Private Function KeiOf(ByVal ws As Worksheet, ByVal col As Long) As String
    Dim h As String, p As Long
    h = HeaderOf(ws, col)
    If Len(h) = 0 Then Exit Function
    p = InStr(h, ChrW(&HFF08))
    If p = 0 Then p = InStr(h, "(")
    If p > 0 Then h = Left$(h, p - 1)
    KeiOf = Trim$(Replace(Replace(h, " ", ""), ChrW(&H3000), ""))
End Function

Private Function KindOf(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As String, out As String
    v = Norm(MergedValue(ws, r, c))
    If InStr(v, "AS") > 0 Then out = "As"
    If InStr(v, "CO") > 0 Then out = out & "Co"
    KindOf = out
End Function

Private Function MergedValue(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim cell_ As Range
    Set cell_ = ws.Cells(r, c)
    If cell_.MergeCells Then
        MergedValue = cell_.MergeArea.Cells(1, 1).Value
    Else
        MergedValue = cell_.Value
    End If
End Function

Private Function FindLabelCol(ByVal ws As Worksheet, ByVal label As String, _
                              ByVal defaultCol As String) As String
    Dim r As Long, c As Long
    For r = 1 To 12
        For c = 1 To 30
            If InStr(1, Norm(ws.Cells(r, c).Value), Norm(label), vbTextCompare) > 0 Then
                FindLabelCol = ColLetterFromNum(c)
                Exit Function
            End If
        Next c
    Next r
    FindLabelCol = defaultCol
End Function

Private Function DiaOfCol(ByRef cols() As TCol, ByVal n As Long, ByVal letter As String) As String
    Dim i As Long
    For i = 0 To n - 1
        If cols(i).Letter = letter Then DiaOfCol = cols(i).Dia: Exit Function
    Next i
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

Private Function HasSheetRef(ByVal c As Range) As Boolean
    Dim f As String
    If Not c.HasFormula Then Exit Function
    f = c.Formula
    If InStr(f, "表紙") > 0 Then Exit Function
    HasSheetRef = (InStr(f, "!") > 0) Or (InStr(f, "INDIRECT") > 0)
End Function

'------------------------------------------------------------------
' 以前このマクロが入れた INDIRECT 形式を、普通のシート参照に戻す。
' 一度変換した後でも読み直せるようにするため。
'------------------------------------------------------------------
Private Function Normalize(ByVal f As String) As String
    Dim re As Object, ms As Object, m As Object
    Dim cfg As Worksheet, out As String, pos As Long, dia As String

    Normalize = f
    If InStr(f, "INDIRECT") = 0 Then Exit Function

    Set cfg = FindSheet(ThisWorkbook, LNK_SHEET)
    If cfg Is Nothing Then Exit Function

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
        If Len(dia) = 0 Then Exit Function
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

Private Function PrefixOf(ByVal tmpl As String) As String
    Dim re As Object, ms As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "'([^']+)#'!"
    Set ms = re.Execute(tmpl)
    If ms.Count > 0 Then PrefixOf = ms(0).SubMatches(0)
End Function

Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    CountChar = Len(s) - Len(Replace(s, ch, ""))
End Function

Private Function ColLetterOf(ByVal c As Range) As String
    ColLetterOf = Split(c.Address(True, False), "$")(0)
End Function

Private Function ColLetterFromNum(ByVal n As Long) As String
    ColLetterFromNum = Split(ThisWorkbook.Worksheets(1).Cells(1, n).Address(True, False), "$")(0)
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
