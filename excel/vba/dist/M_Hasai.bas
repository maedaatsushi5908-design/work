Attribute VB_Name = "M_Hasai"
'==================================================================
' 06 土工事・舗装復旧 数量計算書 用
'
' このファイル1つだけを標準モジュールに貼り付ければ動く。
' マクロは「総括表の数量を転記する」1本。
'
' 元は excel/vba/src/ の M_Hasai.bas を
' つなげたもの。直すときは src 側を直して build_vba.py を実行する。
'==================================================================
Option Explicit

'--- M_Hasai.bas の宣言 ---------------------------------------
'==================================================================
' M_Hasai － 総括表（土工事）へ数量を転記する
'
' マクロは1本だけ。
'
'     総括表の数量を転記する()
'
' いまのところ2つの工種を扱う。転記元の並び方が違うので式も違う。
'
' ● 舗装版破砕工 … SUMIF で厚さを照合する
'
'   =SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)
'           └ 転記元の舗装厚欄      └ 総括表の厚さ欄      └ 転記元の合計欄
'
'   転記元の舗装厚は、その工事に出てくる厚さだけが詰めて並ぶ。
'   4cm が無い工事では 5cm が先頭に来るので、同じセルを見ていると
'   別の厚さの数量を拾ってしまう。だから厚さで照合する。
'
' ● 舗装切断工 … 行を探して直接参照する
'
'   ='試掘（舗50'!P5
'
'   こちらの厚さ区分は t≦15 / 15<t≦30 / 30<t≦40 と固定の文字で、
'   詰めて動くことがない。だから直接参照でよい。
'   ただし行の位置はシートによって違う（試掘は4行目から、管工は10行目から）
'   ので、マクロを流すたびに区分を照合して行を決め直す。
'
' 書き込む先は、総括表で黄色く塗ってある入力セルだけ。
' As の行は転記元の As 側、Co の行は Co 側、As・Co の行は両方を足す。
'==================================================================

'==================================================================
' ここだけ工事に合わせて直す
'==================================================================

' 総括表のシート名
Private Const TARGET_SHEET As String = "総括表（土工事）"

' 総括表の列 → 転記元シート名。区切りは縦棒。
' 使わない列は書かなくてよい。
Private Const COL_MAP As String = _
    "J=試掘（舗50|" & _
    "K=試掘（舗300|" & _
    "L=試掘（舗75|" & _
    "M=試掘（舗400|" & _
    "N=試掘（舗600|" & _
    "O=給水2度|" & _
    "P=仮配（舗|" & _
    "Q=給水(舗|" & _
    "R=管工（舗50|" & _
    "S=管工（舗75|" & _
    "T=管工（舗400|" & _
    "U=管工（舗600"

' --- 別の工事に持っていくときの覚え書き ------------------------------
' シート名の数字は口径ではなく、様式の枠の名前。延長表のどの区分を
' 読んでいるかで決まる。2つの工事で確かめたところ、どちらも同じ。
'
'     1枠目 （舗50   … ■新設部
'     3枠目 （舗75   … ■一般部・■撤去部
'
' だから「(75-300)新設」という見出しでも転記元は 管工（舗50 になる。
' 見出しの数字だけで決めると 新設 と 取替・撤去 を取り違える。
'
' 仮配管・給水付替は口径で分かれないので、シート名は工事が変わっても
' 仮配（舗 / 給水(舗 のまま（給水( は半角カッコ。全角ではない）。
' 給水2度 は見出しが □舗装版取壊工 で並びも違うが、マクロが見分ける。
'
' 長田(花山町2丁目他)配水管取替工事は管工事が5列に増えている。
' そちらへ持っていくときの COL_MAP:
'    "J=試掘（舗50|K=試掘（舗300|L=試掘（舗75|M=試掘（舗400|N=試掘（舗600|"
'    "P=仮配（舗|Q=給水(舗|"
'    "R=管工（舗50|S=管工（舗200|T=管工（舗75|U=管工（舗400|V=管工（舗600"
' ただし長田は総括表が別ブック（01 総括表）にあるので、このままでは動かない。
' 下書きは python3 excel/docs/scan_work.py excel/works/<工事> で作れる。
' --------------------------------------------------------------------

' 扱う工種。総括表のいちばん左の工種名と、転記元のブロック見出し。
' 工種を増やすときはここに足して、WriteAll の振り分けに1行足す。
Private Const SECTION_LABEL As String = "舗装版破砕"
Private Const BLOCK_LABEL As String = "□舗装版破砕"
Private Const CUT_LABEL As String = "舗装切断工"
Private Const CUT_BLOCK As String = "□舗装切断工"

' 給水2度だけは並びが違う。「種別・舗装厚／合 計」ではなく
' 「車道 5号工」「歩道 5号工」の2枠で、どちらも As。
' 総括表では2枠を足す。Co の枠は無い。
Private Const TORI_BLOCK As String = "□舗装版取壊工"
Private Const TORI_KIND As String = "As"

' 入力セルの色（黄色）。総括表の凡例と同じ色
Private Const INPUT_COLOR As Long = 65535

' 入れた式の一覧を残すシート。要らなければ消してよい
Private Const REPORT_SHEET As String = "転記結果"

' 書き込んだ1セル分の控え
Private Type TRec
    Addr_   As String
    Sect    As String
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

'==================================================================
' ここから M_Hasai.bas
'==================================================================
'==================================================================
' 唯一の入口
'==================================================================
Public Sub 総括表の数量を転記する()
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

    firstCol = FirstMapCol(cols, nCol)
    nRow = CountSectionRows(ws, firstCol)
    thkCol = ThickCol(ws, firstCol)
    If thkCol = 0 Then
        MsgBox SECTION_LABEL & " の行に舗装厚の数値が見つかりません。" & vbCrLf & _
               "対象シートが違うかもしれません。", vbExclamation, "舗装版破砕の転記"
        Exit Sub
    End If

    ' --- 何をするかを見せて確認 --------------------------------------
    msg = "対象シート: " & ws.Name & vbCrLf & _
          "対象の行: " & nRow & " 行（" & SECTION_LABEL & "・" & CUT_LABEL & "）" & vbCrLf & _
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
    Dim sect As String
    Dim cel As Range, thkCell As Range

    why = ""
    lastRow = LastUsedRow(ws)
    ReDim recs(0 To 400)

    For r = 1 To lastRow
        sect = SectionOf(ws, r, firstCol)
        If Len(sect) > 0 Then
            Set thkCell = ws.Cells(r, thkCol).MergeArea.Cells(1, 1)
            kind = KindOfRow(ws, r, firstCol)

            ' 舗装版破砕は厚さが数値。「計」などの文字が入っている行は
            ' 合計欄なので触らない。空欄の行は書き込む。予備の行に厚さを
            ' 入れたとき、そのまま数量が出るようにするため。
            ' 舗装切断工は厚さが区分の文字（t≦15㎝）なので、文字でよい。
            Dim ok As Boolean
            If sect = SECTION_LABEL Then
                ok = Not IsTextCell(thkCell)
            Else
                ok = IsTextCell(thkCell)
            End If

            If ok And Len(kind) > 0 Then
                thkRef = "$" & ColLetter(thkCell.Column) & thkCell.Row
                For i = 0 To nCol - 1
                    Set cel = ws.Cells(r, ColNum(cols(i)))
                    If IsInputCell(cel) Then
                        note = ""
                        If sect = SECTION_LABEL Then
                            f = BuildSumif(srcs(i), ws.Name, thkRef, kind, note)
                        Else
                            f = CutRef(srcs(i), CStr(thkCell.Text), kind, note)
                        End If
                        If Len(f) = 0 Then
                            nSkip = nSkip + 1
                        Else
                            If n > UBound(recs) Then ReDim Preserve recs(0 To n + 200)
                            recs(n).Addr_ = cel.Address(False, False)
                            recs(n).Sect = sect
                            recs(n).Thick = CStr(thkCell.Text)
                            recs(n).Kind = kind
                            recs(n).Src = srcs(i)
                            recs(n).OldV = NumOrEmpty(cel)
                            On Error Resume Next
                            cel.Formula = f
                            If Err.Number <> 0 Then
                                note = cel.Address(False, False) & " に入れられません: " & Err.Description
                                Err.Clear
                                On Error GoTo 0
                                nSkip = nSkip + 1
                                If InStr(why, note) = 0 Then why = why & "　" & note & vbCrLf
                                GoTo NextCol
                            End If
                            On Error GoTo 0
                            ' Excel が書き換えた後の姿を控える（不要な引用符は落ちる）
                            recs(n).Expr = cel.Formula
                            n = n + 1
                        End If
                        If Len(note) > 0 And InStr(why, note) = 0 Then
                            why = why & "　" & note & vbCrLf
                        End If
                    End If
NextCol:
                Next i
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
        If SectionOf(ws, r, firstCol) = SECTION_LABEL Then
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

    rp.Range("A1").Value = ws.Name & " に入れた式　" & _
                           Format$(Now, "yyyy/mm/dd hh:nn")
    rp.Range("A1").Font.Bold = True
    rp.Range("A2").Value = "※ このシートは要らなければ消してかまいません"
    rp.Range("A2").Font.Italic = True

    r = 4
    rp.Cells(r, 1).Value = "セル"
    rp.Cells(r, 2).Value = "工種"
    rp.Cells(r, 3).Value = "舗装厚"
    rp.Cells(r, 4).Value = "種別"
    rp.Cells(r, 5).Value = "転記元"
    rp.Cells(r, 6).Value = "前の値"
    rp.Cells(r, 7).Value = "後の値"
    rp.Cells(r, 8).Value = "入れた式"
    rp.Range(rp.Cells(r, 1), rp.Cells(r, 8)).Font.Bold = True
    rp.Range(rp.Cells(r, 1), rp.Cells(r, 8)).Interior.Color = RGB(226, 239, 218)

    For i = 0 To n - 1
        r = r + 1
        rp.Cells(r, 1).Value = recs(i).Addr_
        rp.Cells(r, 2).Value = recs(i).Sect
        rp.Cells(r, 3).Value = recs(i).Thick
        rp.Cells(r, 4).Value = recs(i).Kind
        rp.Cells(r, 5).Value = recs(i).Src
        If Not IsEmpty(recs(i).OldV) Then rp.Cells(r, 6).Value = recs(i).OldV
        rp.Cells(r, 7).Value = ws.Range(recs(i).Addr_).Value
        rp.Cells(r, 8).Value = "'" & recs(i).Expr
        ' 値が変わったところに色を付ける
        If Differs(recs(i).OldV, ws.Range(recs(i).Addr_).Value) Then
            rp.Range(rp.Cells(r, 1), rp.Cells(r, 8)).Interior.Color = RGB(255, 235, 156)
        End If
    Next i

    rp.Columns.AutoFit
    If rp.Columns(8).ColumnWidth > 70 Then rp.Columns(8).ColumnWidth = 70
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
    Dim r0 As Long, r1 As Long, pairs As String, p As Variant
    Dim out As String, crit As String, a As Variant

    If Not HasaiBlock(sn, r0, r1, pairs) Then
        note = sn & " に " & BLOCK_LABEL & " のブロックがありません"
        Exit Function
    End If

    crit = SheetRef(tName) & thkRef

    ' 種別の合う枠だけを足す。給水2度のように As の枠が2つある
    ' シートでは、どちらも足す（車道＋歩道）。
    For Each p In Split(pairs, "|")
        a = Split(CStr(p), ",")
        If InStr(kind, CStr(a(0))) > 0 Then
            If Len(out) > 0 Then out = out & "+"
            out = out & SumifTerm(sn, CLng(a(1)), CLng(a(2)), CLng(a(3)), r0, r1, crit)
        End If
    Next p

    If Len(out) = 0 Then
        note = sn & " に " & kind & " 側の欄がありません"
        Exit Function
    End If
    If InStr(kind, "Co") > 0 And InStr(pairs, "Co") = 0 Then
        note = sn & " に Co 側の欄が無いため As だけを合計しました"
    End If
    BuildSumif = "=" & out
End Function

' 舗装版破砕の転記元ブロック。給水2度だけ見出しが違うので順に試す
Private Function HasaiBlock(ByVal sn As String, ByRef r0 As Long, ByRef r1 As Long, _
                            ByRef pairs As String) As Boolean
    If FindBlock(sn, BLOCK_LABEL, r0, r1, pairs) Then
        HasaiBlock = True
    ElseIf FindBlock(sn, TORI_BLOCK, r0, r1, pairs) Then
        HasaiBlock = True
    End If
End Function

Private Function SumifTerm(ByVal sn As String, ByVal thkCol As Long, ByVal sumCol As Long, _
                           ByVal sumWide As Long, ByVal r0 As Long, ByVal r1 As Long, _
                           ByVal crit As String) As String
    Dim q As String
    q = SheetRef(sn)
    SumifTerm = "SUMIF(" & q & Rng(thkCol, 1, r0, r1) & "," & crit & "," & _
                q & Rng(sumCol, sumWide, r0, r1) & ")"
End Function

' 数式に書くシート名。Excel と同じで、囲む必要のある名前だけ ' で囲む。
'   試掘（舗50 → '試掘（舗50'!      （ を含むので囲む
'   給水2度    → 給水2度!            囲む必要がない
Private Function SheetRef(ByVal sn As String) As String
    If NeedsQuote(sn) Then
        SheetRef = "'" & Replace(sn, "'", "''") & "'!"
    Else
        SheetRef = sn & "!"
    End If
End Function

Private Function NeedsQuote(ByVal sn As String) As Boolean
    Dim i As Long, ch As Long
    If Len(sn) = 0 Then NeedsQuote = True: Exit Function
    ch = AscW(Left$(sn, 1))
    If ch >= 48 And ch <= 57 Then NeedsQuote = True: Exit Function     ' 数字で始まる
    For i = 1 To Len(sn)
        ch = AscW(Mid$(sn, i, 1))
        If ch < 0 Then ch = ch + 65536
        If Not SafeChar(ch) Then NeedsQuote = True: Exit Function
    Next i
End Function

' 囲まなくてよい文字か。迷ったら囲む側に倒す（囲んでも式は正しい）
Private Function SafeChar(ByVal ch As Long) As Boolean
    If ch >= 48 And ch <= 57 Then SafeChar = True: Exit Function       ' 0-9
    If ch >= 65 And ch <= 90 Then SafeChar = True: Exit Function       ' A-Z
    If ch >= 97 And ch <= 122 Then SafeChar = True: Exit Function      ' a-z
    If ch = 95 Then SafeChar = True: Exit Function                     ' _
    If ch = &H3000 Or ch = &H30FB Then Exit Function                   ' 全角空白・中黒は囲む
    If ch >= &H3041 And ch <= &H30FF Then SafeChar = True: Exit Function   ' かな
    If ch >= &H4E00 And ch <= &H9FFF Then SafeChar = True: Exit Function   ' 漢字
    If ch >= &HFF66 And ch <= &HFF9F Then SafeChar = True: Exit Function   ' 半角カナ
End Function

' 数式に書く範囲。合計側の幅は転記元の結合セルに合わせる。
' SUMIF は条件範囲と同じ形の分しか見ないので、広くても結果は同じ。
Private Function Rng(ByVal col As Long, ByVal wide As Long, _
                     ByVal r0 As Long, ByVal r1 As Long) As String
    Rng = "$" & ColLetter(col) & "$" & r0 & ":$" & ColLetter(col + wide - 1) & "$" & r1
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
Private Function FindBlock(ByVal sn As String, ByVal anchor As String, _
                           ByRef r0 As Long, ByRef r1 As Long, _
                           ByRef pairs As String) As Boolean
    Dim ws As Worksheet, r As Long, c As Long, sect As Long
    Dim key As String, cached As String, v As String

    key = anchor & "|" & sn
    If mBlock Is Nothing Then Set mBlock = CreateObject("Scripting.Dictionary")
    If mBlock.Exists(key) Then
        cached = mBlock(key)
        If Len(cached) = 0 Then Exit Function
        r0 = CLng(Split(cached, ";")(0))
        r1 = CLng(Split(cached, ";")(1))
        pairs = Split(cached, ";")(2)
        FindBlock = True
        Exit Function
    End If
    mBlock(key) = ""

    Set ws = FindSheet(sn)
    If ws Is Nothing Then Exit Function

    ' ブロックの見出しを探す
    For r = 1 To 60
        For c = 1 To 60
            If InStr(Norm(ws.Cells(r, c).Value), Norm(anchor)) > 0 Then
                sect = r
                Exit For
            End If
        Next c
        If sect > 0 Then Exit For
    Next r
    If sect = 0 Then Exit Function

    pairs = ""
    If Not ScanKindHeader(ws, sect, r0, r1, pairs) Then
        If Not ScanGokouHeader(ws, sect, r0, r1, pairs) Then Exit Function
    End If

    mBlock(key) = r0 & ";" & r1 & ";" & pairs
    FindBlock = True
End Function

'------------------------------------------------------------------
' 並び その1 －「種別・舗装厚」と「合 計」が並ぶ形
'
'   種別・舗装厚 | 合 計 ‖ 種別・舗装厚 | 合 計
'   As    4      |  0.6  ‖ Co    15     |  0.3
'
' 「種別・舗装厚」の1つ右が厚さ、その右で最初に来る「合 計」が数量。
' 行は As / Co が続くところまで。次の工事で舗装厚が1種類増えても
' 式を直さずに済むよう、予備を1行足す。SUMIF は空欄に当たらない。
'------------------------------------------------------------------
Private Function ScanKindHeader(ByVal ws As Worksheet, ByVal sect As Long, _
                                ByRef r0 As Long, ByRef r1 As Long, _
                                ByRef pairs As String) As Boolean
    Dim r As Long, c As Long, hdr As Long, i As Long, v As String
    Dim kinds As String, totals As String, kArr As Variant, tArr As Variant
    Dim kc As Long, sc As Long, firstThk As Long, kind As String

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

    ' 種別が As / Co と書いてある行が続くところまで
    firstThk = CLng(kArr(0)) + 1
    r0 = hdr + 1
    r1 = r0 - 1
    For r = r0 To r0 + 40
        v = Norm(ws.Cells(r, firstThk - 1).Value)
        If v <> "AS" And v <> "CO" Then Exit For
        r1 = r
    Next r
    If r1 < r0 Then Exit Function

    For i = 0 To UBound(kArr)
        kc = CLng(kArr(i))
        sc = NextTotal(tArr, kc)
        If sc > 0 Then
            kind = Norm(ws.Cells(r0, kc).Value)
            If kind <> "AS" And kind <> "CO" Then
                kind = IIf(Len(pairs) = 0, "As", "Co")   ' 読めなければ左から As, Co
            Else
                kind = IIf(kind = "AS", "As", "Co")
            End If
            pairs = pairs & IIf(Len(pairs) = 0, "", "|") & _
                    kind & "," & (kc + 1) & "," & sc & "," & MergeWide(ws, r0, sc)
        End If
        If UBound(Split(pairs, "|")) >= 1 Then Exit For   ' As と Co の2枠まで
    Next i
    If Len(pairs) = 0 Then Exit Function

    r1 = r1 + 1                       ' 予備を1行
    ScanKindHeader = True
End Function

'------------------------------------------------------------------
' 並び その2 －「車道 5号工」「歩道 5号工」が並ぶ形（給水2度）
'
'   車道　5号工（K:N 結合）  ‖ 歩道　5号工（O:R 結合）
'   K=舗装厚 | L:N=数量      ‖ O=舗装厚 | P:R=数量
'
' 結合の左端が厚さ、その次から右端までが数量。どちらの枠も As なので、
' 総括表では2枠を足す。行は厚さ欄が埋まっているところまで（予備は付けない）。
'------------------------------------------------------------------
Private Function ScanGokouHeader(ByVal ws As Worksheet, ByVal sect As Long, _
                                 ByRef r0 As Long, ByRef r1 As Long, _
                                 ByRef pairs As String) As Boolean
    Dim r As Long, c As Long, hdr As Long, cols As String
    Dim p As Variant, mc As Range, thk As Long, sum_ As Long, w As Long

    For r = sect To sect + 3
        cols = ""
        For c = 1 To 60
            If InStr(Norm(ws.Cells(r, c).Value), Norm("号工")) > 0 Then
                cols = cols & c & ","
            End If
        Next c
        If Len(cols) > 0 Then
            hdr = r
            Exit For
        End If
    Next r
    If hdr = 0 Then Exit Function

    r0 = hdr + 1
    r1 = r0 - 1
    thk = CLng(Split(cols, ",")(0))
    For r = r0 To r0 + 40
        If IsEmpty(ws.Cells(r, thk).Value) Then Exit For
        r1 = r
    Next r
    If r1 < r0 Then Exit Function

    For Each p In Split(Left$(cols, Len(cols) - 1), ",")
        Set mc = ws.Cells(hdr, CLng(p)).MergeArea
        thk = mc.Column
        sum_ = thk + 1
        w = mc.Columns.Count - 1
        If w < 1 Then w = 1
        pairs = pairs & IIf(Len(pairs) = 0, "", "|") & _
                TORI_KIND & "," & thk & "," & sum_ & "," & w
    Next p
    If Len(pairs) = 0 Then Exit Function
    ScanGokouHeader = True
End Function

' 結合しているセルの横幅（列数）
Private Function MergeWide(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Long
    MergeWide = ws.Cells(r, c).MergeArea.Columns.Count
    If MergeWide < 1 Then MergeWide = 1
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
    If FindSheet(sn) Is Nothing Then
        BlockNote = "★シートがありません"
        Exit Function
    End If
    Dim r0 As Long, r1 As Long, pairs As String, h As String
    If HasaiBlock(sn, r0, r1, pairs) Then
        h = r0 & "-" & r1 & "行(" & KindsOf(pairs) & ")"
    Else
        h = "★なし"
    End If
    BlockNote = "破砕" & h & "  切断" & Where(sn, CUT_BLOCK)
End Function

Private Function Where(ByVal sn As String, ByVal anchor As String) As String
    Dim r0 As Long, r1 As Long, pairs As String
    If FindBlock(sn, anchor, r0, r1, pairs) Then
        Where = r0 & "-" & r1 & "行(" & KindsOf(pairs) & ")"
    Else
        Where = "★なし"
    End If
End Function

' 枠の一覧を「As+Co」「As×2」のように短く表す
Private Function KindsOf(ByVal pairs As String) As String
    Dim p As Variant, out As String, n As Long
    For Each p In Split(pairs, "|")
        out = out & IIf(Len(out) = 0, "", "+") & Split(CStr(p), ",")(0)
        n = n + 1
    Next p
    If n = 2 And Split(pairs, "|")(0) = Split(pairs, "|")(1) Then out = "As×2"
    KindsOf = out
End Function

'------------------------------------------------------------------
' 舗装切断工の1セル分。行を探して直接参照にする
'
'   ='試掘（舗50'!P5
'
' 厚さ区分は固定の文字なので、詰めて動くことはない。ただし行の位置は
' シートによって違う（試掘は4行目から、管工は10行目から）ので、
' 区分を照合して行を決める。総括表は「t≦15㎝」、転記元は「t≦15」と
' ㎝ の有無が違うため、㎝ を落として突き合わせる。
'------------------------------------------------------------------
Private Function CutRef(ByVal sn As String, ByVal label As String, _
                        ByVal kind As String, ByRef note As String) As String
    Dim r0 As Long, r1 As Long, pairs As String, p As Variant, a As Variant
    Dim thkCol As Long, sumCol As Long, r As Long, ws As Worksheet, want As String

    If Not FindBlock(sn, CUT_BLOCK, r0, r1, pairs) Then
        note = sn & " に " & CUT_BLOCK & " のブロックがありません"
        Exit Function
    End If

    For Each p In Split(pairs, "|")
        a = Split(CStr(p), ",")
        If InStr(kind, CStr(a(0))) > 0 Then
            thkCol = CLng(a(1)): sumCol = CLng(a(2))
            Exit For
        End If
    Next p
    If sumCol = 0 Then
        note = sn & " に " & kind & " 側の欄がありません"
        Exit Function
    End If

    want = NormLabel(label)
    If Len(want) = 0 Then
        note = "総括表の厚さ区分が読めません"
        Exit Function
    End If

    Set ws = FindSheet(sn)
    For r = r0 To r1
        If NormLabel(ws.Cells(r, thkCol).Value) = want Then
            CutRef = "=" & SheetRef(sn) & ColLetter(sumCol) & r
            Exit Function
        End If
    Next r
    note = sn & " に「" & label & "」の行がありません"
End Function

' 厚さ区分の文字をそろえる。t≦15㎝ も t≦15 も同じものとして扱う
Private Function NormLabel(ByVal v As Variant) As String
    Dim t As String
    t = Norm(v)
    t = Replace(t, ChrW(&H339D), "")      ' ㎝（U+339D）
    t = Replace(t, "CM", "")
    NormLabel = t
End Function

'==================================================================
' 総括表side の読み取り
'==================================================================

' いちばん左の工種名で、どの工種の行かを返す。対象外なら空文字
Private Function SectionOf(ByVal ws As Worksheet, ByVal r As Long, _
                           ByVal firstCol As Long) As String
    Dim c As Long, v As Variant, t As String
    For c = 1 To firstCol - 1
        v = MergedValue(ws, r, c)
        If Not IsEmpty(v) Then
            t = CStr(v)
            If Len(Trim$(t)) > 0 And Left$(t, 1) <> "=" Then
                t = Norm(t)
                If InStr(t, Norm(SECTION_LABEL)) > 0 Then
                    SectionOf = SECTION_LABEL
                ElseIf InStr(t, Norm(CUT_LABEL)) > 0 Then
                    SectionOf = CUT_LABEL
                End If
                Exit Function
            End If
        End If
    Next c
End Function

Private Function IsSectionRow(ByVal ws As Worksheet, ByVal r As Long, _
                              ByVal firstCol As Long) As Boolean
    IsSectionRow = (Len(SectionOf(ws, r, firstCol)) > 0)
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
    Dim nHas As Long, nCut As Long, nKind As Long, nYellow As Long
    Dim nB1 As Long, nB2 As Long, sect As String
    Dim r0 As Long, r1 As Long, pairs As String

    lastRow = LastUsedRow(ws)
    thkCol = ThickCol(ws, firstCol)
    For r = 1 To lastRow
        sect = SectionOf(ws, r, firstCol)
        If Len(sect) > 0 Then
            If sect = SECTION_LABEL Then nHas = nHas + 1 Else nCut = nCut + 1
            If Len(KindOfRow(ws, r, firstCol)) > 0 Then nKind = nKind + 1
            For i = 0 To nCol - 1
                If IsInputCell(ws.Cells(r, ColNum(cols(i)))) Then nYellow = nYellow + 1
            Next i
        End If
    Next r
    For i = 0 To nCol - 1
        If HasaiBlock(srcs(i), r0, r1, pairs) Then nB1 = nB1 + 1
        If FindBlock(srcs(i), CUT_BLOCK, r0, r1, pairs) Then nB2 = nB2 + 1
    Next i

    Diagnose = _
        "調べたところ" & vbCrLf & _
        "　" & SECTION_LABEL & " の行 … " & nHas & " 行" & vbCrLf & _
        "　" & CUT_LABEL & " の行 … " & nCut & " 行" & vbCrLf & _
        "　厚さ列 … " & IIf(thkCol > 0, ColLetter(thkCol) & " 列", "見つからない") & vbCrLf & _
        "　種別(As/Co)が読めた行 … " & nKind & " 行" & vbCrLf & _
        "　黄色い入力セル … " & nYellow & " 個" & vbCrLf & _
        "　" & BLOCK_LABEL & " が見つかった転記元 … " & nB1 & " / " & nCol & " シート" & vbCrLf & _
        "　" & CUT_BLOCK & " が見つかった転記元 … " & nB2 & " / " & nCol & " シート" & vbCrLf & vbCrLf & _
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
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If Norm(sh.Name) = Norm(nm) Then
            Set FindSheet = sh
            Exit Function
        End If
    Next sh
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
