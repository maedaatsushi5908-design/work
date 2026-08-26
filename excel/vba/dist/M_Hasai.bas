Attribute VB_Name = "M_Hasai"
'==================================================================
' 舗装版破砕の転記を SUMIF に置き換える（個別指示ぶん）
'
' 総括表（土工事）の 32 セルに、舗装厚で照合する式を入れる。
' 解析はしない。下の一覧をそのまま書き込むだけなので、
' 何がどこに入るかはコードを読めば全部わかる。
'
'   =SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)
'          └ 転記元の舗装厚      └ 総括表の厚さ欄   └ 転記元の合計
'
' 転記元の舗装厚は工事ごとに詰めて並ぶため、同じセルを見ていると
' 別の厚さの数量を拾う。厚さが一致する行だけを合計すれば間違えない。
'
' 実行すると、書き込む前にバックアップシートを作る。
' 値が変わったセルを一覧で見せ、納得できなければ元に戻せる。
'==================================================================
Option Explicit

Private Const DEST As String = "総括表（土工事）"

Public Sub 舗装版破砕の式を入れる()
    Dim ws As Worksheet, i As Long, n As Long
    Dim cells_() As String, fmls() As String
    Dim oldF() As String, oldV() As Variant, newV() As Variant
    Dim nDiff As Long, scr As Boolean, calc As XlCalculation

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEST)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シートが見つかりません: " & DEST, vbExclamation
        Exit Sub
    End If

    Defs cells_, fmls, n

    If MsgBox(DEST & " の " & n & " セルに、舗装厚で照合する式を入れます。" & vbCrLf & vbCrLf & _
              "書き込む前にバックアップシートを作ります。続けますか？", _
              vbYesNo + vbQuestion, "確認") <> vbYes Then Exit Sub

    scr = Application.ScreenUpdating
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Backup ws

    ReDim oldF(0 To n - 1): ReDim oldV(0 To n - 1): ReDim newV(0 To n - 1)
    For i = 0 To n - 1
        With ws.Range(cells_(i))
            oldF(i) = .Formula
            oldV(i) = CellNum(ws, cells_(i))
            .Formula = fmls(i)
        End With
    Next i

    Application.Calculation = calc
    Application.CalculateFull
    For i = 0 To n - 1
        newV(i) = CellNum(ws, cells_(i))
    Next i

    nDiff = Report(ws, cells_, fmls, oldV, newV, n)
    Application.ScreenUpdating = scr

    If nDiff = 0 Then
        MsgBox n & " セルに式を入れました。" & vbCrLf & _
               "値が変わったセルはありません。", vbInformation, "完了"
        Exit Sub
    End If

    If MsgBox(n & " セルに式を入れました。" & vbCrLf & _
              "そのうち " & nDiff & " セルで値が変わっています。" & vbCrLf & vbCrLf & _
              "「破砕の式 結果」シートで内容を確かめてください。" & vbCrLf & vbCrLf & _
              "このまま確定しますか？" & vbCrLf & _
              "「いいえ」ですべて元に戻します。", _
              vbYesNo + vbQuestion, "値が変わりました") = vbYes Then
        MsgBox "確定しました。", vbInformation, "完了"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For i = n - 1 To 0 Step -1
        If Len(oldF(i)) = 0 Then
            ws.Range(cells_(i)).ClearContents
        Else
            ws.Range(cells_(i)).Formula = oldF(i)
        End If
    Next i
    Application.ScreenUpdating = scr
    Application.CalculateFull
    MsgBox "元に戻しました。", vbInformation, "取り消し"
End Sub

'==================================================================
' 入れる式の一覧
'   セル      厚さ 種別   転記元
'==================================================================
Private Sub Defs(ByRef cells_() As String, ByRef fmls() As String, ByRef n As Long)
    ReDim cells_(0 To 200): ReDim fmls(0 To 200)
    n = 0

    ' --- 仮配（舗 ----------------------------------------
    ' P51    厚さ 15  AsCo   Co側の欄が無い
    A cells_, fmls, n, "P51", _
      "=SUMIF('仮配（舗'!$R$10:$R$17,'総括表（土工事）'!$I51,'仮配（舗'!$W$10:$W$17)"
    ' P52    厚さ 14  AsCo   Co側の欄が無い
    A cells_, fmls, n, "P52", _
      "=SUMIF('仮配（舗'!$R$10:$R$17,'総括表（土工事）'!$I52,'仮配（舗'!$W$10:$W$17)"
    ' P66    厚さ 20  As   
    A cells_, fmls, n, "P66", _
      "=SUMIF('仮配（舗'!$R$10:$R$17,'総括表（土工事）'!$I66,'仮配（舗'!$W$10:$W$17)"

    ' --- 管工（舗400 -------------------------------------
    ' T41    厚さ  4  AsCo 
    A cells_, fmls, n, "T41", _
      "=SUMIF('管工（舗400'!$R$16:$R$23,'総括表（土工事）'!$I41,'管工（舗400'!$T$16:$T$23)+SUMIF('管工（舗400'!$V$16:$V$23,'総括表（土工事）'!$I41,'管工（舗400'!$X$16:$X$23)"
    ' T42    厚さ  5  AsCo 
    A cells_, fmls, n, "T42", _
      "=SUMIF('管工（舗400'!$R$16:$R$23,'総括表（土工事）'!$I42,'管工（舗400'!$T$16:$T$23)+SUMIF('管工（舗400'!$V$16:$V$23,'総括表（土工事）'!$I42,'管工（舗400'!$X$16:$X$23)"
    ' T43    厚さ 10  AsCo 
    A cells_, fmls, n, "T43", _
      "=SUMIF('管工（舗400'!$R$16:$R$23,'総括表（土工事）'!$I43,'管工（舗400'!$T$16:$T$23)+SUMIF('管工（舗400'!$V$16:$V$23,'総括表（土工事）'!$I43,'管工（舗400'!$X$16:$X$23)"
    ' T60    厚さ 15  AsCo 
    A cells_, fmls, n, "T60", _
      "=SUMIF('管工（舗400'!$R$16:$R$23,'総括表（土工事）'!$I60,'管工（舗400'!$T$16:$T$23)+SUMIF('管工（舗400'!$V$16:$V$23,'総括表（土工事）'!$I60,'管工（舗400'!$X$16:$X$23)"

    ' --- 管工（舗50 --------------------------------------
    ' R31    厚さ  4  AsCo 
    A cells_, fmls, n, "R31", _
      "=SUMIF('管工（舗50'!$R$16:$R$23,'総括表（土工事）'!$I31,'管工（舗50'!$T$16:$T$23)+SUMIF('管工（舗50'!$V$16:$V$23,'総括表（土工事）'!$I31,'管工（舗50'!$X$16:$X$23)"
    ' R32    厚さ  5  AsCo 
    A cells_, fmls, n, "R32", _
      "=SUMIF('管工（舗50'!$R$16:$R$23,'総括表（土工事）'!$I32,'管工（舗50'!$T$16:$T$23)+SUMIF('管工（舗50'!$V$16:$V$23,'総括表（土工事）'!$I32,'管工（舗50'!$X$16:$X$23)"
    ' R33    厚さ 10  AsCo 
    A cells_, fmls, n, "R33", _
      "=SUMIF('管工（舗50'!$R$16:$R$23,'総括表（土工事）'!$I33,'管工（舗50'!$T$16:$T$23)+SUMIF('管工（舗50'!$V$16:$V$23,'総括表（土工事）'!$I33,'管工（舗50'!$X$16:$X$23)"
    ' R54    厚さ 15  AsCo 
    A cells_, fmls, n, "R54", _
      "=SUMIF('管工（舗50'!$R$16:$R$23,'総括表（土工事）'!$I54,'管工（舗50'!$T$16:$T$23)+SUMIF('管工（舗50'!$V$16:$V$23,'総括表（土工事）'!$I54,'管工（舗50'!$X$16:$X$23)"
    ' R55    厚さ 14  AsCo 
    A cells_, fmls, n, "R55", _
      "=SUMIF('管工（舗50'!$R$16:$R$23,'総括表（土工事）'!$I55,'管工（舗50'!$T$16:$T$23)+SUMIF('管工（舗50'!$V$16:$V$23,'総括表（土工事）'!$I55,'管工（舗50'!$X$16:$X$23)"

    ' --- 管工（舗600 -------------------------------------
    ' U46    厚さ  4  AsCo 
    A cells_, fmls, n, "U46", _
      "=SUMIF('管工（舗600'!$R$16:$R$23,'総括表（土工事）'!$I46,'管工（舗600'!$T$16:$T$23)+SUMIF('管工（舗600'!$V$16:$V$23,'総括表（土工事）'!$I46,'管工（舗600'!$X$16:$X$23)"
    ' U47    厚さ  5  AsCo 
    A cells_, fmls, n, "U47", _
      "=SUMIF('管工（舗600'!$R$16:$R$23,'総括表（土工事）'!$I47,'管工（舗600'!$T$16:$T$23)+SUMIF('管工（舗600'!$V$16:$V$23,'総括表（土工事）'!$I47,'管工（舗600'!$X$16:$X$23)"
    ' U48    厚さ 10  AsCo 
    A cells_, fmls, n, "U48", _
      "=SUMIF('管工（舗600'!$R$16:$R$23,'総括表（土工事）'!$I48,'管工（舗600'!$T$16:$T$23)+SUMIF('管工（舗600'!$V$16:$V$23,'総括表（土工事）'!$I48,'管工（舗600'!$X$16:$X$23)"

    ' --- 管工（舗75 --------------------------------------
    ' S36    厚さ  4  AsCo 
    A cells_, fmls, n, "S36", _
      "=SUMIF('管工（舗75'!$R$16:$R$23,'総括表（土工事）'!$I36,'管工（舗75'!$T$16:$T$23)+SUMIF('管工（舗75'!$V$16:$V$23,'総括表（土工事）'!$I36,'管工（舗75'!$X$16:$X$23)"
    ' S37    厚さ  5  AsCo 
    A cells_, fmls, n, "S37", _
      "=SUMIF('管工（舗75'!$R$16:$R$23,'総括表（土工事）'!$I37,'管工（舗75'!$T$16:$T$23)+SUMIF('管工（舗75'!$V$16:$V$23,'総括表（土工事）'!$I37,'管工（舗75'!$X$16:$X$23)"
    ' S38    厚さ 10  AsCo 
    A cells_, fmls, n, "S38", _
      "=SUMIF('管工（舗75'!$R$16:$R$23,'総括表（土工事）'!$I38,'管工（舗75'!$T$16:$T$23)+SUMIF('管工（舗75'!$V$16:$V$23,'総括表（土工事）'!$I38,'管工（舗75'!$X$16:$X$23)"
    ' S57    厚さ 15  AsCo 
    A cells_, fmls, n, "S57", _
      "=SUMIF('管工（舗75'!$R$16:$R$23,'総括表（土工事）'!$I57,'管工（舗75'!$T$16:$T$23)+SUMIF('管工（舗75'!$V$16:$V$23,'総括表（土工事）'!$I57,'管工（舗75'!$X$16:$X$23)"

    ' --- 給水(舗 ----------------------------------------
    ' Q31    厚さ  4  AsCo 
    A cells_, fmls, n, "Q31", _
      "=SUMIF('給水(舗'!$R$10:$R$16,'総括表（土工事）'!$I31,'給水(舗'!$T$10:$T$16)+SUMIF('給水(舗'!$W$10:$W$16,'総括表（土工事）'!$I31,'給水(舗'!$Y$10:$Y$16)"
    ' Q32    厚さ  5  AsCo 
    A cells_, fmls, n, "Q32", _
      "=SUMIF('給水(舗'!$R$10:$R$16,'総括表（土工事）'!$I32,'給水(舗'!$T$10:$T$16)+SUMIF('給水(舗'!$W$10:$W$16,'総括表（土工事）'!$I32,'給水(舗'!$Y$10:$Y$16)"
    ' Q33    厚さ 10  AsCo 
    A cells_, fmls, n, "Q33", _
      "=SUMIF('給水(舗'!$R$10:$R$16,'総括表（土工事）'!$I33,'給水(舗'!$T$10:$T$16)+SUMIF('給水(舗'!$W$10:$W$16,'総括表（土工事）'!$I33,'給水(舗'!$Y$10:$Y$16)"
    ' Q54    厚さ 15  AsCo 
    A cells_, fmls, n, "Q54", _
      "=SUMIF('給水(舗'!$R$10:$R$16,'総括表（土工事）'!$I54,'給水(舗'!$T$10:$T$16)+SUMIF('給水(舗'!$W$10:$W$16,'総括表（土工事）'!$I54,'給水(舗'!$Y$10:$Y$16)"
    ' Q55    厚さ 14  AsCo 
    A cells_, fmls, n, "Q55", _
      "=SUMIF('給水(舗'!$R$10:$R$16,'総括表（土工事）'!$I55,'給水(舗'!$T$10:$T$16)+SUMIF('給水(舗'!$W$10:$W$16,'総括表（土工事）'!$I55,'給水(舗'!$Y$10:$Y$16)"
    ' Q66    厚さ 20  As   
    A cells_, fmls, n, "Q66", _
      "=SUMIF('給水(舗'!$R$10:$R$16,'総括表（土工事）'!$I66,'給水(舗'!$T$10:$T$16)"
    ' Q67    厚さ 20  Co   
    A cells_, fmls, n, "Q67", _
      "=SUMIF('給水(舗'!$W$10:$W$16,'総括表（土工事）'!$I67,'給水(舗'!$Y$10:$Y$16)"

    ' --- 試掘（舗400 -------------------------------------
    ' M13    厚さ  4  As   
    A cells_, fmls, n, "M13", _
      "=SUMIF('試掘（舗400'!$O$11:$O$17,'総括表（土工事）'!$I13,'試掘（舗400'!$P$11:$P$17)"
    ' M15    厚さ 10  As   
    A cells_, fmls, n, "M15", _
      "=SUMIF('試掘（舗400'!$O$11:$O$17,'総括表（土工事）'!$I15,'試掘（舗400'!$P$11:$P$17)"
    ' M19    厚さ 15  As   
    A cells_, fmls, n, "M19", _
      "=SUMIF('試掘（舗400'!$O$11:$O$17,'総括表（土工事）'!$I19,'試掘（舗400'!$P$11:$P$17)"
    ' M24    厚さ 15  Co   
    A cells_, fmls, n, "M24", _
      "=SUMIF('試掘（舗400'!$R$11:$R$17,'総括表（土工事）'!$I24,'試掘（舗400'!$S$11:$S$17)"

    ' --- 試掘（舗50 --------------------------------------
    ' J13    厚さ  4  As   
    A cells_, fmls, n, "J13", _
      "=SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I13,'試掘（舗50'!$P$11:$P$17)"
    ' J14    厚さ  5  As   
    A cells_, fmls, n, "J14", _
      "=SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,'試掘（舗50'!$P$11:$P$17)"
End Sub

Private Sub A(ByRef cells_() As String, ByRef fmls() As String, ByRef n As Long, _
              ByVal cellAddr As String, ByVal fml As String)
    If n > UBound(cells_) Then
        ReDim Preserve cells_(0 To n + 50)
        ReDim Preserve fmls(0 To n + 50)
    End If
    cells_(n) = cellAddr
    fmls(n) = fml
    n = n + 1
End Sub

'==================================================================
' 補助
'==================================================================
Private Function CellNum(ByVal ws As Worksheet, ByVal a As String) As Variant
    Dim v As Variant
    v = ws.Range(a).Value
    If IsError(v) Then
        CellNum = Empty
    ElseIf IsNumeric(v) And Not IsEmpty(v) And CStr(v) <> "" Then
        CellNum = CDbl(v)
    Else
        CellNum = Empty
    End If
End Function

Private Function Report(ByVal ws As Worksheet, ByRef cells_() As String, _
                        ByRef fmls() As String, ByRef oldV() As Variant, _
                        ByRef newV() As Variant, ByVal n As Long) As Long
    Dim rep As Worksheet, i As Long, r As Long, nd As Long
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("破砕の式 結果").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set rep = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    rep.Name = "破砕の式 結果"
    rep.Range("A1").Value = "舗装版破砕の式　" & Format$(Now, "yyyy/mm/dd hh:nn")
    rep.Range("A1").Font.Bold = True
    rep.Range("A3").Value = "セル"
    rep.Range("B3").Value = "実行前"
    rep.Range("C3").Value = "実行後"
    rep.Range("D3").Value = "差"
    rep.Range("E3").Value = "入れた式"
    With rep.Range("A3:E3")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    r = 4
    For i = 0 To n - 1
        rep.Cells(r, 1).Value = cells_(i)
        If Not IsEmpty(oldV(i)) Then rep.Cells(r, 2).Value = oldV(i)
        If Not IsEmpty(newV(i)) Then rep.Cells(r, 3).Value = newV(i)
        If Differs(oldV(i), newV(i)) Then
            nd = nd + 1
            rep.Cells(r, 1).Interior.Color = RGB(255, 235, 156)
            If Not IsEmpty(oldV(i)) And Not IsEmpty(newV(i)) Then
                rep.Cells(r, 4).Value = CDbl(newV(i)) - CDbl(oldV(i))
            End If
        End If
        rep.Cells(r, 5).Value = "'" & fmls(i)
        r = r + 1
    Next i
    rep.Columns.AutoFit
    If rep.Columns(5).ColumnWidth > 70 Then rep.Columns(5).ColumnWidth = 70
    rep.Activate
    Report = nd
End Function

Private Function Differs(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) And IsEmpty(b) Then Exit Function
    If IsEmpty(a) Or IsEmpty(b) Then Differs = True: Exit Function
    Differs = (Abs(CDbl(a) - CDbl(b)) > 0.00000001)
End Function

Private Sub Backup(ByVal ws As Worksheet)
    Dim nm As String
    nm = "BK_" & Format$(Now, "mmdd_hhnn")
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count).Name = nm
    Application.DisplayAlerts = True
End Sub
