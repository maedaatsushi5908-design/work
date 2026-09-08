Attribute VB_Name = "M_HasaiHosou"
'==================================================================
' 06 総括表（舗装工事） 用
'
' このファイル1つだけを標準モジュールに貼り付ければ動く。
' マクロは「総括表（舗装工事）の数量を転記する」1本。
' 将来的には M_Hasai（総括表（土工事）用）と合わせる予定。
'
' 元は excel/vba/src/ の M_HasaiHosou.bas を
' つなげたもの。直すときは src 側を直して build_vba.py を実行する。
'==================================================================
Option Explicit

'--- M_HasaiHosou.bas の宣言 ----------------------------------
'==================================================================
' M_HasaiHosou － 総括表（舗装工事）へ数量を転記する
'
' M_Hasai（総括表（土工事）用）とは別の、独立して動くマクロ。
' 将来的には M_Hasai と合わせる予定だが、いまはこの1本で
' 総括表（舗装工事）だけを扱う。
'
' マクロは1本だけ。
'
'     総括表（舗装工事）の数量を転記する()
'
' いまのところ対応しているのは CELL_MAP に載っている数セルだけ。
' 総括表（舗装工事）の「舗装版切断（舗装工事）」は、転記元
' （舗装（集計）の「□舗装切断工」の枠）から行を探して直接参照する。
'
'   I7  = As・t≦15    → 舗装（集計）!M4
'   I8  = As・15<t≦30 → 舗装（集計）!M5
'   I9  = Co・t≦15    → 舗装（集計）!P4
'   I10 = Co・15<t≦30 → 舗装（集計）!P5
'
' 書き込む先は、総括表で黄色く塗ってある入力セルだけ。
'==================================================================

'==================================================================
' ここだけ工事に合わせて直す
'==================================================================

' 総括表のシート名
Private Const TARGET_SHEET As String = "総括表（舗装工事）"

' 転記元シート名
Private Const PAVE_SRC As String = "舗装（集計）"

' 総括表のセル → 転記元セルの対応。「セル=セル|セル=セル|…」を並べる。
' 舗装版切断（舗装工事）の並びと厚さ区分は工事によって変わらないはずだが、
' 念のためこの工事で確かめた対応をそのまま書く。別の工事に持っていくときは
' 総括表（舗装工事）と 舗装（集計）の並びを見比べて書き直すこと。
Private Const CELL_MAP As String = "I7=M4|I8=M5|I9=P4|I10=P5"

' 入力セルの色（黄色）。総括表の凡例と同じ色
Private Const INPUT_COLOR As Long = 65535

'==================================================================
' ここから M_HasaiHosou.bas
'==================================================================
'==================================================================
Public Sub 総括表（舗装工事）の数量を転記する()
    Dim ws As Worksheet, srcWs As Worksheet
    Dim p As Variant, kv As Variant
    Dim addr As String, cellRef As String, f As String, msg As String
    Dim nWrite As Long, nSkip As Long, why As String
    Dim scr As Boolean, calc As XlCalculation, bk As String
    Dim cel As Range

    Set ws = FindSheet(TARGET_SHEET)
    If ws Is Nothing Then
        MsgBox "シートが見つかりません: " & TARGET_SHEET & vbCrLf & vbCrLf & _
               "マクロの先頭にある TARGET_SHEET を、" & vbCrLf & _
               "実際のシート名に書き換えてください。", vbExclamation, _
               "総括表（舗装工事）の数量を転記"
        Exit Sub
    End If

    Set srcWs = FindSheet(PAVE_SRC)
    If srcWs Is Nothing Then
        MsgBox "転記元シートが見つかりません: " & PAVE_SRC, vbExclamation, _
               "総括表（舗装工事）の数量を転記"
        Exit Sub
    End If

    On Error GoTo Failed

    ' --- 何をするかを見せて確認 --------------------------------------
    msg = "対象シート: " & ws.Name & vbCrLf & _
          "転記元シート: " & srcWs.Name & vbCrLf & vbCrLf & _
          "黄色い入力セルにだけ、次の直接参照を入れます。" & vbCrLf & _
          "----------------------------------------" & vbCrLf
    For Each p In Split(CELL_MAP, "|")
        kv = Split(CStr(p), "=")
        If UBound(kv) = 1 Then
            msg = msg & "  " & CStr(kv(0)) & " = " & PAVE_SRC & "!" & CStr(kv(1)) & vbCrLf
        End If
    Next p
    msg = msg & "----------------------------------------" & vbCrLf & vbCrLf & _
          "書き込む前にバックアップを取ります。続けますか？"

    If MsgBox(msg, vbYesNo + vbQuestion, "総括表（舗装工事）の数量を転記") <> vbYes Then Exit Sub

    ' --- バックアップ → 書き込み --------------------------------------
    scr = Application.ScreenUpdating
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    bk = MakeBackup(ws)

    For Each p In Split(CELL_MAP, "|")
        kv = Split(CStr(p), "=")
        If UBound(kv) = 1 Then
            addr = CStr(kv(0))
            cellRef = CStr(kv(1))
            Set cel = ws.Range(addr)
            If IsInputCell(cel) Then
                f = "=" & SheetRef(PAVE_SRC) & cellRef
                cel.Formula = f
                nWrite = nWrite + 1
            Else
                nSkip = nSkip + 1
                why = why & "　" & addr & " は黄色でないので見送りました" & vbCrLf
            End If
        End If
    Next p

    Application.Calculation = calc
    Application.CalculateFull
    Application.ScreenUpdating = scr
    On Error GoTo 0

    ' --- 結果を伝える ------------------------------------------------
    If nWrite = 0 Then
        MsgBox "1つも書き込みませんでした。" & vbCrLf & vbCrLf & why & vbCrLf & _
               "バックアップ " & bk & " は消してかまいません。", _
               vbExclamation, "総括表（舗装工事）の数量を転記"
    Else
        msg = nWrite & " 個のセルに数式を入れました。" & vbCrLf
        If nSkip > 0 Then msg = msg & "見送ったセル: " & nSkip & " 個" & vbCrLf & why & vbCrLf
        msg = msg & vbCrLf & "元の状態は " & bk & " シートに残しています。"
        MsgBox msg, vbInformation, "完了"
    End If
    Exit Sub

Failed:
    Application.Calculation = calc
    Application.ScreenUpdating = scr
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical, "総括表（舗装工事）の数量を転記"
End Sub

Private Function FindSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If Norm(sh.Name) = Norm(nm) Then
            Set FindSheet = sh
            Exit Function
        End If
    Next sh
End Function

Private Function IsInputCell(ByVal c As Range) As Boolean
    IsInputCell = (c.Interior.Color = INPUT_COLOR)
End Function

' 数式に書くシート名。囲む必要のある名前だけ ' で囲む
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
