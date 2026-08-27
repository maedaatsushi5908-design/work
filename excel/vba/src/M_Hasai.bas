Attribute VB_Name = "M_Hasai"
'==================================================================
' M_Hasai － 舗装版破砕の数量を、舗装厚で照合して転記する
'
' マクロは1本だけ。
'
'     舗装版破砕を転記する()
'
' 入れる数式はこれだけ。
'
'   =SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)
'           └ 転記元の舗装厚欄      └ 総括表の厚さ欄      └ 転記元の合計欄
'
' 転記元の舗装厚は、その工事に出てくる厚さだけが詰めて並ぶ。
' 4cm が無い工事では 5cm が先頭に来るので、同じセルを見ていると
' 別の厚さの数量を拾ってしまう。だから厚さで照合する。
'
' 書き込む先は、総括表で黄色く塗ってある入力セルだけ。
' As の行は転記元の As 側、Co の行は Co 側、As・Co の行は両方を足す。
'==================================================================
Option Explicit

'==================================================================
' ここだけ工事に合わせて直す
'==================================================================

' 総括表のシート名
Private Const TARGET_SHEET As String = "総括表（土工事）"

' 転記元シートが別のブックにあるときの、そのブックのファイル名。
' 空にしておくと、開いているブックの中から自動で探す。
' 見つけたブック名は実行前の確認画面に出る。
Private Const SOURCE_BOOK As String = ""

' 総括表の列 → 転記元シート名。区切りは縦棒。
' 使わない列は書かなくてよい。
' 「給水2度」は舗装版破砕のブロックが無いので入れていない。
'
' 工事が変わったらここを差し替える。下書きは次で作れる。
'     python3 excel/docs/scan_work.py excel/works/<工事のフォルダ>
'
' シート名の数字は口径ではなく様式の枠の名前。延長表のどの区分を
' 読んでいるかで決まる。1枠目(50)＝新設部、3枠目(75)＝一般部・撤去部。
' 「(75-300)新設」が 管工（舗50 になるのはそのため。
Private Const COL_MAP As String = _
    "J=試掘（舗50|" & _
    "K=試掘（舗300|" & _
    "L=試掘（舗75|" & _
    "M=試掘（舗400|" & _
    "N=試掘（舗600|" & _
    "P=仮配（舗|" & _
    "Q=給水(舗|" & _
    "R=管工（舗50|" & _
    "S=管工（舗200|" & _
    "T=管工（舗75|" & _
    "U=管工（舗400|" & _
    "V=管工（舗600"

' 東白川特２高層配水池（管工事が4列、総括表が 06 のブックの中にある工事）
' はこちら。上の COL_MAP と入れ替えて使う。
'    "J=試掘（舗50|K=試掘（舗300|L=試掘（舗75|M=試掘（舗400|N=試掘（舗600|"
'    "P=仮配（舗|Q=給水(舗|"
'    "R=管工（舗50|S=管工（舗75|T=管工（舗400|U=管工（舗600"

' 総括表で、いちばん左の工種名がこの言葉を含む行だけを対象にする
Private Const SECTION_LABEL As String = "舗装版破砕"

' 転記元シートで、ブロックの先頭にある見出し
Private Const BLOCK_LABEL As String = "□舗装版破砕"

' 入力セルの色（黄色）。総括表の凡例と同じ色
Private Const INPUT_COLOR As Long = 65535

' 入れた式の一覧を残すシート。要らなければ消してよい
Private Const REPORT_SHEET As String = "転記結果"

' 書き込んだ1セル分の控え
Private Type TRec
    Addr_   As String
    Thick   As String
    Kind    As String
    Src     As String
    Expr    As String
    OldV    As Variant
    NewV    As Variant
End Type

'==================================================================
' ここから下は工事が変わっても直さない
'==================================================================

' 転記元シートのブロック位置。同じシートを何度も探さないよう覚えておく
Private mBlock As Object

' 転記元のブック。総括表と同じブックのこともあれば、別ブックのこともある
Private mSrc As Workbook

'==================================================================
' 唯一の入口
'==================================================================
Public Sub 舗装版破砕を転記する()
    Dim ws As Worksheet
    Dim cols() As String, srcs() As String, nCol As Long
    Dim recs() As TRec
    Dim msg As String, i As Long, firstCol As Long, thkCol As Long
    Dim nWrite As Long, nSkip As Long, nRow As Long, why As String
    Dim scr As Boolean, calc As XlCalculation, bk As String

    Set mBlock = CreateObject("Scripting.Dictionary")

    Set ws = FindSheet(TARGET_SHEET)
    If ws Is Nothing Then
        MsgBox "シートが見つかりません: " & TARGET_SHEET & vbCrLf & vbCrLf & _
               "マクロの先頭にある TARGET_SHEET を、" & vbCrLf & _
               "実際のシート名に書き換えてください。", vbExclamation, "舗装版破砕の転記"
        Exit Sub
    End If

    ParseMap cols, srcs, nCol
    If nCol = 0 Then
        MsgBox "COL_MAP が読み取れません。マクロの先頭を確かめてください。", _
               vbExclamation, "舗装版破砕の転記"
        Exit Sub
    End If

    If Not ResolveSource(srcs, nCol, why) Then
        MsgBox why, vbExclamation, "舗装版破砕の転記"
        Exit Sub
    End If

    firstCol = FirstMapCol(cols, nCol)
    nRow = CountSectionRows(ws, firstCol)
    thkCol = ThickCol(ws, firstCol)
    If thkCol = 0 Then
        MsgBox SECTION_LABEL & " の行に舗装厚の数値が見つかりません。" & vbCrLf & _
               "対象シートが違うかもしれません。", vbExclamation, "舗装版破砕の転記"
        Exit Sub
    End If

    ' --- 何をするかを見せて確認 --------------------------------------
    msg = "対象シート: " & ws.Name & "（" & ThisWorkbook.Name & "）" & vbCrLf & _
          "転記元ブック: " & mSrc.Name & _
          IIf(mSrc Is ThisWorkbook, "（同じブック）", "（別のブック）") & vbCrLf & _
          SECTION_LABEL & " の行: " & nRow & " 行" & vbCrLf & _
          "舗装厚の列: " & ColLetter(thkCol) & " 列" & vbCrLf & vbCrLf & _
          "転記元の対応" & vbCrLf & _
          "----------------------------------------" & vbCrLf
    For i = 0 To nCol - 1
        msg = msg & " " & cols(i) & "列 → " & srcs(i) & _
              "  " & BlockNote(srcs(i)) & vbCrLf
    Next i
    msg = msg & "----------------------------------------" & vbCrLf & vbCrLf & _
          "黄色い入力セルにだけ SUMIF を入れます。" & vbCrLf & _
          "書き込む前にバックアップを取ります。" & vbCrLf & vbCrLf & _
          "続けますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "舗装版破砕の転記") <> vbYes Then Exit Sub

    ' --- バックアップ → 書き込み --------------------------------------
    scr = Application.ScreenUpdating
    calc = Application.Calculation
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    bk = MakeBackup(ws)
    nWrite = WriteAll(ws, cols, srcs, nCol, firstCol, thkCol, nSkip, why, recs)

    Application.Calculation = calc
    Application.CalculateFull
    If nWrite > 0 Then WriteReport ws, recs, nWrite
    Application.ScreenUpdating = scr
    On Error GoTo 0

    ' --- 結果を伝える。何もしなかったときは理由を出す ------------------
    If nWrite = 0 Then
        MsgBox "1つも書き込みませんでした。" & vbCrLf & vbCrLf & _
               Diagnose(ws, cols, srcs, nCol, firstCol) & vbCrLf & _
               "バックアップ " & bk & " は消してかまいません。", _
               vbExclamation, "舗装版破砕の転記"
    Else
        msg = nWrite & " 個のセルに数式を入れました。" & vbCrLf
        If nSkip > 0 Then
            msg = msg & "見送ったセル: " & nSkip & " 個" & vbCrLf & why & vbCrLf
        End If
        msg = msg & vbCrLf & _
              "★このシートは「ゼロ値を表示しない」設定です。" & vbCrLf & _
              "　数式を入れても、結果が 0 のセルは画面上は空欄に見えます。" & vbCrLf & _
              "　入れた式は「" & REPORT_SHEET & "」シートに一覧で出しました。" & vbCrLf & _
              "　セル上で見たいときは Ctrl と Shift と @ で数式表示に切り替えます。" & vbCrLf & vbCrLf & _
              "元の状態は " & bk & " シートに残しています。" & vbCrLf & _
              "「計」の行の値が前と変わっていないか確かめてください。"
        MsgBox msg, vbInformation, "完了"
    End If
    Exit Sub

Failed:
    Application.Calculation = calc
    Application.ScreenUpdating = scr
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical, "舗装版破砕の転記"
End Sub

'==================================================================
' 書き込み
'==================================================================
Private Function WriteAll(ByVal ws As Worksheet, ByRef cols() As String, _
                          ByRef srcs() As String, ByVal nCol As Long, _
                          ByVal firstCol As Long, ByVal thkCol As Long, _
                          ByRef nSkip As Long, ByRef why As String, _
                          ByRef recs() As TRec) As Long
    Dim r As Long, i As Long, lastRow As Long, n As Long
    Dim thkRef As String, kind As String, f As String, note As String
    Dim cel As Range, thkCell As Range

    why = ""
    lastRow = LastUsedRow(ws)
    ReDim recs(0 To 400)

    For r = 1 To lastRow
        If IsSectionRow(ws, r, firstCol) Then
            Set thkCell = ws.Cells(r, thkCol).MergeArea.Cells(1, 1)

            ' 「計」などの文字が入っている行は合計欄なので触らない。
            ' 空欄の行は書き込む。予備の行に厚さを入れたとき、
            ' そのまま数量が出るようにするため（SUMIF は空欄に当たらない）。
            If Not IsTextCell(thkCell) Then
                thkRef = "$" & ColLetter(thkCell.Column) & thkCell.Row
                kind = KindOfRow(ws, r, firstCol)
                If Len(kind) > 0 Then
                    For i = 0 To nCol - 1
                        Set cel = ws.Cells(r, ColNum(cols(i)))
                        If IsInputCell(cel) Then
                            note = ""
                            f = BuildSumif(srcs(i), ws.Name, thkRef, kind, note)
                            If Len(f) = 0 Then
                                nSkip = nSkip + 1
                            Else
                                If n > UBound(recs) Then ReDim Preserve recs(0 To n + 200)
                                recs(n).Addr_ = cel.Address(False, False)
                                recs(n).Thick = CStr(thkCell.Text)
                                recs(n).Kind = kind
                                recs(n).Src = srcs(i)
                                recs(n).Expr = f
                                recs(n).OldV = NumOrEmpty(cel)
                                cel.Formula = f
                                n = n + 1
                            End If
                            If Len(note) > 0 And InStr(why, note) = 0 Then
                                why = why & "　" & note & vbCrLf
                            End If
                        End If
                    Next i
                End If
            End If
        End If
    Next r
    WriteAll = n
End Function

' 数値でも空欄でもない、文字の入ったセルか（「計」など）
Private Function IsTextCell(ByVal c As Range) As Boolean
    Dim v As Variant
    v = c.Value
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then IsTextCell = True: Exit Function
    If VarType(v) = vbString Then IsTextCell = (Len(Trim$(CStr(v))) > 0)
End Function

Private Function NumOrEmpty(ByVal c As Range) As Variant
    If IsNum(c.Value) Then NumOrEmpty = c.Value
End Function

'------------------------------------------------------------------
' 舗装厚が入っている列。区間の行のうち、数値が現れた回数がいちばん
' 多い列を採る。空欄の行があっても列を見失わないようにするため。
'------------------------------------------------------------------
Private Function ThickCol(ByVal ws As Worksheet, ByVal firstCol As Long) As Long
    Dim r As Long, c As Long, lastRow As Long, best As Long
    Dim hits() As Long, tr As Range

    lastRow = LastUsedRow(ws)
    ReDim hits(0 To firstCol)

    For r = 1 To lastRow
        If IsSectionRow(ws, r, firstCol) Then
            For c = firstCol - 1 To 1 Step -1
                Set tr = ws.Cells(r, c).MergeArea.Cells(1, 1)
                If IsNum(tr.Value) Then
                    If tr.Column <= firstCol Then hits(tr.Column) = hits(tr.Column) + 1
                    Exit For
                End If
            Next c
        End If
    Next r

    For c = 1 To firstCol
        If hits(c) > best Then
            best = hits(c)
            ThickCol = c
        End If
    Next c
End Function

'------------------------------------------------------------------
' 入れた式の一覧を残す。
' このシートはゼロ値を表示しない設定なので、0 になったセルは
' 画面上は空欄に見える。入ったかどうかはここで確かめる。
'------------------------------------------------------------------
Private Sub WriteReport(ByVal ws As Worksheet, ByRef recs() As TRec, ByVal n As Long)
    Dim rp As Worksheet, i As Long, r As Long

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(REPORT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set rp = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    rp.Name = REPORT_SHEET

    rp.Range("A1").Value = ws.Name & " の " & SECTION_LABEL & " に入れた式　" & _
                           Format$(Now, "yyyy/mm/dd hh:nn")
    rp.Range("A1").Font.Bold = True
    rp.Range("A2").Value = "※ このシートは要らなければ消してかまいません"
    rp.Range("A2").Font.Italic = True

    r = 4
    rp.Cells(r, 1).Value = "セル"
    rp.Cells(r, 2).Value = "舗装厚"
    rp.Cells(r, 3).Value = "種別"
    rp.Cells(r, 4).Value = "転記元"
    rp.Cells(r, 5).Value = "前の値"
    rp.Cells(r, 6).Value = "後の値"
    rp.Cells(r, 7).Value = "入れた式"
    rp.Range(rp.Cells(r, 1), rp.Cells(r, 7)).Font.Bold = True
    rp.Range(rp.Cells(r, 1), rp.Cells(r, 7)).Interior.Color = RGB(226, 239, 218)

    For i = 0 To n - 1
        r = r + 1
        rp.Cells(r, 1).Value = recs(i).Addr_
        rp.Cells(r, 2).Value = recs(i).Thick
        rp.Cells(r, 3).Value = recs(i).Kind
        rp.Cells(r, 4).Value = recs(i).Src
        If Not IsEmpty(recs(i).OldV) Then rp.Cells(r, 5).Value = recs(i).OldV
        rp.Cells(r, 6).Value = ws.Range(recs(i).Addr_).Value
        rp.Cells(r, 7).Value = "'" & recs(i).Expr
        ' 値が変わったところに色を付ける
        If Differs(recs(i).OldV, ws.Range(recs(i).Addr_).Value) Then
            rp.Range(rp.Cells(r, 1), rp.Cells(r, 7)).Interior.Color = RGB(255, 235, 156)
        End If
    Next i

    rp.Columns.AutoFit
    If rp.Columns(7).ColumnWidth > 70 Then rp.Columns(7).ColumnWidth = 70
    On Error Resume Next
    rp.Activate
    rp.Rows(5).Select
    ActiveWindow.FreezePanes = True
    rp.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function Differs(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim x As Double, y As Double
    If IsNum(a) Then x = CDbl(a)
    If IsNum(b) Then y = CDbl(b)
    Differs = (Abs(x - y) > 0.0000001)
End Function

'------------------------------------------------------------------
' 1セル分の数式を組み立てる
'   =SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)
'------------------------------------------------------------------
Private Function BuildSumif(ByVal sn As String, ByVal tName As String, _
                            ByVal thkRef As String, ByVal kind As String, _
                            ByRef note As String) As String
    Dim r0 As Long, r1 As Long
    ' aS は予約語 As と衝突する（VBA は大文字小文字を区別しない）ので使わない
    Dim asThk As Long, asSum As Long, coThk As Long, coSum As Long
    Dim out As String, crit As String

    If Not FindBlock(sn, r0, r1, asThk, asSum, coThk, coSum) Then
        note = sn & " に " & BLOCK_LABEL & " のブロックがありません"
        Exit Function
    End If

    crit = "'" & tName & "'!" & thkRef

    If InStr(kind, "As") > 0 And asSum > 0 Then
        out = SumifTerm(sn, asThk, asSum, r0, r1, crit)
    End If
    If InStr(kind, "Co") > 0 And coSum > 0 Then
        If Len(out) > 0 Then out = out & "+"
        out = out & SumifTerm(sn, coThk, coSum, r0, r1, crit)
    End If

    If Len(out) = 0 Then
        note = sn & " に " & kind & " 側の欄がありません"
        Exit Function
    End If
    If InStr(kind, "Co") > 0 And coSum = 0 Then
        note = sn & " に Co 側の欄が無いため As だけを合計しました"
    End If
    BuildSumif = "=" & out
End Function

Private Function SumifTerm(ByVal sn As String, ByVal thkCol As Long, ByVal sumCol As Long, _
                           ByVal r0 As Long, ByVal r1 As Long, ByVal crit As String) As String
    Dim q As String
    q = Qual(sn)
    SumifTerm = "SUMIF(" & q & Rng(thkCol, r0, r1) & "," & crit & "," & q & Rng(sumCol, r0, r1) & ")"
End Function

Private Function Rng(ByVal col As Long, ByVal r0 As Long, ByVal r1 As Long) As String
    Dim s As String
    s = ColLetter(col)
    Rng = "$" & s & "$" & r0 & ":$" & s & "$" & r1
End Function

'==================================================================
' 転記元シートの舗装版破砕ブロックを探す
'
'   □舗装版破砕工
'   種別・舗装厚 | 面積 | 合 計 |  種別・舗装厚 | 面積 | 合 計
'   As     4    |      |  1.2  |  Co     15   |      |  0.3
'
' 「種別・舗装厚」の1つ右が厚さ、その右で最初に来る「合 計」が数量。
' 行は As / Co が続くところまで。次の工事で厚さが1種類増えても
' 式を直さずに済むよう、予備を1行足す。SUMIF は空欄に当たらない。
'==================================================================
Private Function FindBlock(ByVal sn As String, ByRef r0 As Long, ByRef r1 As Long, _
                           ByRef asThk As Long, ByRef asSum As Long, _
                           ByRef coThk As Long, ByRef coSum As Long) As Boolean
    Dim ws As Worksheet, r As Long, c As Long, sect As Long, hdr As Long
    Dim kinds As String, totals As String, i As Long, v As String
    Dim kArr As Variant, tArr As Variant, cached As String

    If mBlock Is Nothing Then Set mBlock = CreateObject("Scripting.Dictionary")
    If mBlock.Exists(sn) Then
        cached = mBlock(sn)
        If Len(cached) = 0 Then Exit Function
        r0 = CLng(Split(cached, "|")(0))
        r1 = CLng(Split(cached, "|")(1))
        asThk = CLng(Split(cached, "|")(2))
        asSum = CLng(Split(cached, "|")(3))
        coThk = CLng(Split(cached, "|")(4))
        coSum = CLng(Split(cached, "|")(5))
        FindBlock = True
        Exit Function
    End If
    mBlock(sn) = ""

    Set ws = FindSheetIn(mSrc, sn)
    If ws Is Nothing Then Exit Function

    ' ブロックの見出しを探す
    For r = 1 To 60
        For c = 1 To 60
            If InStr(Norm(ws.Cells(r, c).Value), Norm(BLOCK_LABEL)) > 0 Then
                sect = r
                Exit For
            End If
        Next c
        If sect > 0 Then Exit For
    Next r
    If sect = 0 Then Exit Function

    ' その下3行以内にある「種別・舗装厚」と「合 計」の行
    For r = sect + 1 To sect + 3
        kinds = "": totals = ""
        For c = 1 To 60
            v = Norm(ws.Cells(r, c).Value)
            If v = Norm("種別・舗装厚") Then kinds = kinds & c & ","
            If v = Norm("合 計") Then totals = totals & c & ","
        Next c
        If Len(kinds) > 0 And Len(totals) > 0 Then
            hdr = r
            Exit For
        End If
    Next r
    If hdr = 0 Then Exit Function

    kArr = Split(Left$(kinds, Len(kinds) - 1), ",")
    tArr = Split(Left$(totals, Len(totals) - 1), ",")

    asThk = 0: asSum = 0: coThk = 0: coSum = 0
    For i = 0 To UBound(kArr)
        Dim kc As Long, sc As Long
        kc = CLng(kArr(i))
        sc = NextTotal(tArr, kc)
        If sc > 0 Then
            If asSum = 0 Then
                asThk = kc + 1: asSum = sc
            ElseIf coSum = 0 Then
                coThk = kc + 1: coSum = sc
                Exit For
            End If
        End If
    Next i
    If asSum = 0 Then Exit Function

    ' 種別が As / Co と書いてある行が続くところまで
    r0 = hdr + 1
    r1 = r0 - 1
    For r = r0 To r0 + 40
        v = Norm(ws.Cells(r, asThk - 1).Value)
        If v <> "AS" And v <> "CO" Then Exit For
        r1 = r
    Next r
    If r1 < r0 Then Exit Function
    r1 = r1 + 1                       ' 予備を1行

    mBlock(sn) = r0 & "|" & r1 & "|" & asThk & "|" & asSum & "|" & coThk & "|" & coSum
    FindBlock = True
End Function

Private Function NextTotal(ByVal tArr As Variant, ByVal kc As Long) As Long
    Dim i As Long, best As Long, t As Long
    For i = 0 To UBound(tArr)
        t = CLng(tArr(i))
        If t > kc Then
            If best = 0 Or t < best Then best = t
        End If
    Next i
    NextTotal = best
End Function

' 確認画面に出す一言
Private Function BlockNote(ByVal sn As String) As String
    Dim r0 As Long, r1 As Long
    Dim asThk As Long, asSum As Long, coThk As Long, coSum As Long

    If FindSheetIn(mSrc, sn) Is Nothing Then
        BlockNote = "★シートがありません"
        Exit Function
    End If
    If Not FindBlock(sn, r0, r1, asThk, asSum, coThk, coSum) Then
        BlockNote = "★破砕のブロックがありません"
        Exit Function
    End If
    BlockNote = "(" & r0 & "-" & r1 & "行 As=" & ColLetter(asThk) & "/" & ColLetter(asSum)
    If coSum > 0 Then
        BlockNote = BlockNote & " Co=" & ColLetter(coThk) & "/" & ColLetter(coSum) & ")"
    Else
        BlockNote = BlockNote & " Coなし)"
    End If
End Function

'==================================================================
' 総括表side の読み取り
'==================================================================

' いちばん左の工種名が SECTION_LABEL を含む行か
Private Function IsSectionRow(ByVal ws As Worksheet, ByVal r As Long, _
                              ByVal firstCol As Long) As Boolean
    Dim c As Long, v As Variant, t As String
    For c = 1 To firstCol - 1
        v = MergedValue(ws, r, c)
        If Not IsEmpty(v) Then
            t = CStr(v)
            If Len(Trim$(t)) > 0 And Left$(t, 1) <> "=" Then
                IsSectionRow = (InStr(Norm(t), Norm(SECTION_LABEL)) > 0)
                Exit Function
            End If
        End If
    Next c
End Function

Private Function CountSectionRows(ByVal ws As Worksheet, ByVal firstCol As Long) As Long
    Dim r As Long, n As Long, lastRow As Long
    lastRow = LastUsedRow(ws)
    For r = 1 To lastRow
        If IsSectionRow(ws, r, firstCol) Then n = n + 1
    Next r
    CountSectionRows = n
End Function

' 舗装厚が入っているセルへの参照。列は固定、行は相対（$I14 の形）
' 数値が入っているか。「計」などの文字や空欄を弾く
' （IsNumeric は Empty に対して True を返すので、そのままでは使えない）
Private Function IsNum(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    IsNum = IsNumeric(v)
End Function

' その行の種別。As / Co / AsCo
Private Function KindOfRow(ByVal ws As Worksheet, ByVal r As Long, _
                           ByVal firstCol As Long) As String
    Dim c As Long, t As String, out As String
    For c = 1 To firstCol - 1
        t = Norm(MergedValue(ws, r, c))
        If InStr(t, "AS") > 0 And InStr(out, "As") = 0 Then out = out & "As"
        If InStr(t, "CO") > 0 And InStr(out, "Co") = 0 Then out = out & "Co"
    Next c
    KindOfRow = out
End Function

' 黄色く塗ってある入力セルか
Private Function IsInputCell(ByVal c As Range) As Boolean
    On Error Resume Next
    If c.Interior.Pattern = xlNone Then Exit Function
    IsInputCell = (c.Interior.Color = INPUT_COLOR)
    On Error GoTo 0
End Function

'==================================================================
' 何も書き込めなかったときに、どこで止まったかを出す
'==================================================================
Private Function Diagnose(ByVal ws As Worksheet, ByRef cols() As String, _
                          ByRef srcs() As String, ByVal nCol As Long, _
                          ByVal firstCol As Long) As String
    Dim r As Long, i As Long, lastRow As Long, thkCol As Long
    Dim nSect As Long, nThk As Long, nKind As Long, nYellow As Long, nBlock As Long
    Dim r0 As Long, r1 As Long, a1 As Long, a2 As Long, c1 As Long, c2 As Long

    lastRow = LastUsedRow(ws)
    thkCol = ThickCol(ws, firstCol)
    For r = 1 To lastRow
        If IsSectionRow(ws, r, firstCol) Then
            nSect = nSect + 1
            If thkCol > 0 Then
                If Not IsTextCell(ws.Cells(r, thkCol).MergeArea.Cells(1, 1)) Then
                    nThk = nThk + 1
                End If
            End If
            If Len(KindOfRow(ws, r, firstCol)) > 0 Then nKind = nKind + 1
            For i = 0 To nCol - 1
                If IsInputCell(ws.Cells(r, ColNum(cols(i)))) Then nYellow = nYellow + 1
            Next i
        End If
    Next r
    For i = 0 To nCol - 1
        If FindBlock(srcs(i), r0, r1, a1, a2, c1, c2) Then nBlock = nBlock + 1
    Next i

    Diagnose = _
        "調べたところ" & vbCrLf & _
        "　" & SECTION_LABEL & " の行 … " & nSect & " 行" & vbCrLf & _
        "　うち舗装厚の欄が使える行 … " & nThk & " 行" & "（厚さ列 " & IIf(thkCol > 0, ColLetter(thkCol), "?") & "）" & vbCrLf & _
        "　うち種別(As/Co)が読めた行 … " & nKind & " 行" & vbCrLf & _
        "　黄色い入力セル … " & nYellow & " 個" & vbCrLf & _
        "　ブロックが見つかった転記元 … " & nBlock & " / " & nCol & " シート" & vbCrLf & vbCrLf & _
        "0 になっている項目が原因です。"
End Function

'==================================================================
' こまごま
'==================================================================
Private Sub ParseMap(ByRef cols() As String, ByRef srcs() As String, ByRef n As Long)
    Dim p As Variant, kv As Variant
    ReDim cols(0 To 60)
    ReDim srcs(0 To 60)
    n = 0
    For Each p In Split(COL_MAP, "|")
        If InStr(CStr(p), "=") > 0 Then
            kv = Split(CStr(p), "=")
            If Len(Trim$(CStr(kv(0)))) > 0 And Len(Trim$(CStr(kv(1)))) > 0 Then
                cols(n) = UCase$(Trim$(CStr(kv(0))))
                srcs(n) = Trim$(CStr(kv(1)))
                n = n + 1
            End If
        End If
    Next p
End Sub

Private Function FirstMapCol(ByRef cols() As String, ByVal n As Long) As Long
    Dim i As Long, c As Long
    FirstMapCol = 999
    For i = 0 To n - 1
        c = ColNum(cols(i))
        If c < FirstMapCol Then FirstMapCol = c
    Next i
    If FirstMapCol > 900 Then FirstMapCol = 2
End Function

Private Function FindSheet(ByVal nm As String) As Worksheet
    Set FindSheet = FindSheetIn(ThisWorkbook, nm)
End Function

Private Function FindSheetIn(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    If wb Is Nothing Then Exit Function
    For Each sh In wb.Worksheets
        If Norm(sh.Name) = Norm(nm) Then
            Set FindSheetIn = sh
            Exit Function
        End If
    Next sh
End Function

'------------------------------------------------------------------
' 転記元のブックを決める
'
' 総括表と転記元が同じブックのこともあれば（1件目の工事）、
' 総括表だけ 01 総括表.xlsx に分かれていることもある（2件目の工事）。
' COL_MAP の1枚目のシートがどこにあるかで見分ける。
'------------------------------------------------------------------
Private Function ResolveSource(ByRef srcs() As String, ByVal nCol As Long, _
                               ByRef why As String) As Boolean
    Dim wb As Workbook, hit As Workbook, n As Long, probe As String, i As Long

    why = ""
    Set mSrc = Nothing

    ' 名前を書いてあるなら、それを開いているか確かめる
    If Len(SOURCE_BOOK) > 0 Then
        For Each wb In Application.Workbooks
            If Norm(wb.Name) = Norm(SOURCE_BOOK) Then Set mSrc = wb: Exit For
        Next wb
        If mSrc Is Nothing Then
            why = "転記元のブック「" & SOURCE_BOOK & "」が開いていません。" & vbCrLf & _
                  "先に開いてから、もう一度実行してください。"
            Exit Function
        End If
        ResolveSource = True
        Exit Function
    End If

    ' 空欄なら、COL_MAP のシートを持つブックを探す
    For i = 0 To nCol - 1
        If Len(srcs(i)) > 0 Then probe = srcs(i): Exit For
    Next i
    If Len(probe) = 0 Then
        why = "COL_MAP が空です。"
        Exit Function
    End If

    If Not FindSheetIn(ThisWorkbook, probe) Is Nothing Then
        Set mSrc = ThisWorkbook              ' 同じブックの中にある
        ResolveSource = True
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If Not FindSheetIn(wb, probe) Is Nothing Then
            Set hit = wb
            n = n + 1
        End If
    Next wb

    If n = 0 Then
        why = "転記元のシート「" & probe & "」が、開いているどのブックにもありません。" & vbCrLf & vbCrLf & _
              "転記元のブック（試掘や管工のシートが入っているもの）を" & vbCrLf & _
              "開いてから、もう一度実行してください。"
        Exit Function
    ElseIf n > 1 Then
        why = "「" & probe & "」というシートを持つブックが " & n & " つ開いています。" & vbCrLf & _
              "どれを使うか決められません。余分なほうを閉じるか、" & vbCrLf & _
              "マクロ先頭の SOURCE_BOOK にファイル名を書いてください。"
        Exit Function
    End If

    Set mSrc = hit
    ResolveSource = True
End Function

' 数式に書くシートの頭。別ブックなら [ファイル名] を付ける
'   同じブック : '試掘（舗50'!
'   別のブック : '[06 土工事・舗装復旧.xlsx]試掘（舗50'!
Private Function Qual(ByVal sn As String) As String
    If mSrc Is Nothing Then
        Qual = "'" & sn & "'!"
    ElseIf mSrc Is ThisWorkbook Then
        Qual = "'" & sn & "'!"
    Else
        Qual = "'[" & mSrc.Name & "]" & sn & "'!"
    End If
End Function

' 使われている最後の行。変数名 lastRow と同じ名前にすると
' VBA は大文字小文字を区別しないため衝突するので、別の名前にしている。
Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    LastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If LastUsedRow > 500 Then LastUsedRow = 500
    If LastUsedRow < 1 Then LastUsedRow = 1
End Function

Private Function MergedValue(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    MergedValue = ws.Cells(r, c).MergeArea.Cells(1, 1).Value
End Function

' 全角を半角に直し、空白を落として大文字にそろえる
Private Function Norm(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As Long, out As String
    If IsError(v) Then Exit Function
    s = CStr(v)
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch >= &HFF01 And ch <= &HFF5E Then ch = ch - &HFEE0   ' 全角英数記号
        If ch <> 32 And ch <> &H3000 And ch <> 9 And ch <> 10 And ch <> 13 Then
            out = out & ChrW(ch)
        End If
    Next i
    Norm = UCase$(out)
End Function

Private Function ColNum(ByVal letter As String) As Long
    Dim i As Long, s As String
    s = UCase$(Trim$(letter))
    For i = 1 To Len(s)
        ColNum = ColNum * 26 + (Asc(Mid$(s, i, 1)) - 64)
    Next i
End Function

Private Function ColLetter(ByVal n As Long) As String
    Dim k As Long, s As String
    k = n
    Do While k > 0
        s = Chr$(65 + (k - 1) Mod 26) & s
        k = (k - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function MakeBackup(ByVal ws As Worksheet) As String
    Dim nm As String, bk As Worksheet
    nm = "BK_" & Format$(Now, "mmdd_hhnn")

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set bk = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    bk.Name = nm
    bk.Visible = xlSheetVisible
    MakeBackup = nm
End Function
