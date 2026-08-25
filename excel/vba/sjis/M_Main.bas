Attribute VB_Name = "M_Main"
'==================================================================
' M_Main - 実行の入口
'
'   照合実行()       … 書き換えずに、総括表の値と計算値を突き合わせる
'   転記実行()       … 差異を確認したうえで総括表に書き込む
'   定義シートを作成() … M_Config にある（最初に一度だけ実行）
'
' 使い方は excel/vba/README.md を参照。
'==================================================================
Option Explicit

Private Type TResult
    No       As Long
    Kind     As String
    ItemName As String
    Spec     As String
    DestCell As String
    OldVal   As Variant
    NewVal   As Variant
    Judge    As String     ' 一致 / 差異 / 転記 / エラー / 手入力 / 対象外
    Source   As String
    Message  As String
End Type

'==================================================================
' 照合（書き換えない）
'==================================================================
Public Sub 照合実行()
    Dim results() As TResult, n As Long
    n = RunAll(results, False)
    If n < 0 Then Exit Sub
    WriteReport results, n, "照合"
    MsgBox Summary(results, n, False), vbInformation, "照合が終わりました"
End Sub

'==================================================================
' 転記（確認のうえ書き換える）
'==================================================================
Public Sub 転記実行()
    Dim results() As TResult, n As Long, diffs As Long, i As Long
    Dim ws As Worksheet, ans As VbMsgBoxResult

    ' まず書き換えずに差分を出す
    n = RunAll(results, False)
    If n < 0 Then Exit Sub

    For i = 0 To n - 1
        If results(i).Judge = "差異" Then diffs = diffs + 1
    Next i

    If diffs = 0 Then
        WriteReport results, n, "照合"
        MsgBox "総括表はすべて計算値と一致しています。" & vbCrLf & _
               "書き換える必要はありませんでした。", vbInformation, "転記は不要です"
        Exit Sub
    End If

    ans = MsgBox(diffs & " 件の差異が見つかりました。" & vbCrLf & vbCrLf & _
                 "総括表に書き込みますか？" & vbCrLf & _
                 "（書き込む前に、総括表のバックアップシートを作ります）", _
                 vbYesNo + vbQuestion, "転記の確認")
    If ans <> vbYes Then
        WriteReport results, n, "照合"
        MsgBox "書き込みは行いませんでした。差分は「" & REPORT_SHEET & "」を確認してください。", _
               vbInformation
        Exit Sub
    End If

    MakeBackup

    ' 実際に書き込む
    n = RunAll(results, True)
    If n < 0 Then Exit Sub
    WriteReport results, n, "転記"
    MsgBox Summary(results, n, True), vbInformation, "転記が終わりました"
End Sub

'==================================================================
' 本体：定義に沿って値を解決し、必要なら書き込む
'   戻り値 … 件数。異常時は -1
'==================================================================
Private Function RunAll(ByRef results() As TResult, ByVal doWrite As Boolean) As Long
    Dim defs() As TDef, cnt As Long, i As Long
    Dim ws As Worksheet, cell_ As Range
    Dim v As Variant, msg As String
    Dim nmActual As String, spActual As String
    Dim scr As Boolean, calc As XlCalculation

    RunAll = -1

    cnt = LoadDefs(defs)
    If cnt = 0 Then Exit Function

    Set ws = FindSheet(ThisWorkbook, DEST_SHEET)
    If ws Is Nothing Then
        MsgBox "転記先シート「" & DEST_SHEET & "」が見つかりません。" & vbCrLf & _
               "M_Config の DEST_SHEET を実際のシート名に合わせてください。", vbExclamation
        Exit Function
    End If

    scr = Application.ScreenUpdating
    calc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    InitSources
    ReDim results(0 To cnt - 1)

    On Error GoTo Cleanup
    For i = 0 To cnt - 1
        With results(i)
            .No = defs(i).No
            .Kind = defs(i).Kind
            .ItemName = defs(i).ItemName
            .Spec = defs(i).Spec
            .DestCell = defs(i).DestCell
            .Source = defs(i).SrcFile & IIf(Len(defs(i).SrcSheet) > 0, " / " & defs(i).SrcSheet, "")
            .OldVal = Empty
            .NewVal = Empty
            .Message = ""
        End With

        ' 転記先セルの妥当性
        Set cell_ = Nothing
        On Error Resume Next
        Set cell_ = ws.Range(results(i).DestCell)
        On Error GoTo Cleanup
        If cell_ Is Nothing Then
            results(i).Judge = "エラー"
            results(i).Message = "転記先セルが不正です: " & results(i).DestCell
            GoTo NextItem
        End If

        results(i).OldVal = NumOrEmpty(cell_)

        ' 総括表側の名称・摘要が定義とずれていないか確認
        DestLabels ws, cell_, nmActual, spActual
        If Len(defs(i).ItemName) > 0 Then
            If Not KeyMatches(nmActual, defs(i).ItemName, mmPartial) Then
                results(i).Message = "※総括表の名称が定義と違います（実際: " & nmActual & "）"
            End If
        End If

        If Not defs(i).Enabled Then
            results(i).Judge = "対象外"
            GoTo NextItem
        End If

        If defs(i).Kind <> "自動転記" Then
            results(i).Judge = "手入力"
            results(i).Message = Trim$(results(i).Message & " " & defs(i).Note)
            GoTo NextItem
        End If

        ' 転記元から値を取る
        If Not ResolveValue(defs(i), v, msg) Then
            results(i).Judge = "エラー"
            results(i).Message = Trim$(results(i).Message & " " & msg)
            GoTo NextItem
        End If

        results(i).NewVal = v

        If Not IsEmpty(results(i).OldVal) Then
            If Abs(CDbl(results(i).OldVal) - CDbl(v)) < 0.00000001 Then
                results(i).Judge = "一致"
                GoTo NextItem
            End If
        End If

        If doWrite Then
            cell_.Value = v
            results(i).Judge = "転記"
        Else
            results(i).Judge = "差異"
        End If

NextItem:
    Next i

    RunAll = cnt

Cleanup:
    If Err.Number <> 0 Then
        MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
               Err.Number & ": " & Err.Description, vbCritical
        RunAll = -1
    End If
    CloseSources
    Application.DisplayAlerts = True
    Application.Calculation = calc
    Application.ScreenUpdating = scr
    Application.CalculateFull
End Function

'------------------------------------------------------------------
' 転記先セルの行から、総括表上の名称・摘要を読む
'   D列 → 名称B 摘要C ／ J列 → 名称F 摘要G
'   名称は結合や省略で空のことがあるので上方向に補完する
'------------------------------------------------------------------
Private Sub DestLabels(ByVal ws As Worksheet, ByVal cell_ As Range, _
                       ByRef outName As String, ByRef outSpec As String)
    Dim nameCol As Long, specCol As Long, r As Long, i As Long

    outName = "": outSpec = ""
    r = cell_.Row

    Select Case cell_.Column
        Case 4:  nameCol = 2: specCol = 3     ' D列
        Case 10: nameCol = 6: specCol = 7     ' J列
        Case Else: Exit Sub
    End Select

    outSpec = Trim$(CStr(ws.Cells(r, specCol).Value))

    For i = r To Application.Max(1, r - 40) Step -1
        If Len(Norm(ws.Cells(i, nameCol).Value)) > 0 Then
            outName = " " & Trim$(CStr(ws.Cells(i, nameCol).Value))
            outName = Trim$(Replace(Replace(outName, vbLf, " "), vbCr, " "))
            Exit For
        End If
    Next i
End Sub

'------------------------------------------------------------------
' 総括表のバックアップシートを作る
'------------------------------------------------------------------
Private Sub MakeBackup()
    Dim ws As Worksheet, bk As Worksheet, nm As String

    Set ws = FindSheet(ThisWorkbook, DEST_SHEET)
    If ws Is Nothing Then Exit Sub

    nm = "BK_" & Format$(Now, "mmdd_hhnn")
    If Len(nm) > 31 Then nm = Left$(nm, 31)

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set bk = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    bk.Name = nm
    bk.Visible = xlSheetVisible
    Application.DisplayAlerts = True
End Sub

'------------------------------------------------------------------
' 結果シートを書き出す
'------------------------------------------------------------------
Private Sub WriteReport(ByRef results() As TResult, ByVal n As Long, ByVal mode As String)
    Dim ws As Worksheet, i As Long, r As Long
    Dim hdr As Variant, c As Long

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(REPORT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = REPORT_SHEET

    ws.Range("A1").Value = mode & "結果　" & Format$(Now, "yyyy/mm/dd hh:nn")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12

    hdr = Array("No", "区分", "名称", "摘要", "転記先", "総括表の値", "計算値", "差", "判定", "転記元", "メッセージ")
    For c = 0 To UBound(hdr)
        ws.Cells(3, c + 1).Value = hdr(c)
    Next c
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, UBound(hdr) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
        .HorizontalAlignment = xlCenter
    End With

    r = 4
    For i = 0 To n - 1
        With results(i)
            ws.Cells(r, 1).Value = .No
            ws.Cells(r, 2).Value = .Kind
            ws.Cells(r, 3).Value = .ItemName
            ws.Cells(r, 4).Value = .Spec
            ws.Cells(r, 5).Value = .DestCell
            If Not IsEmpty(.OldVal) Then ws.Cells(r, 6).Value = .OldVal
            If Not IsEmpty(.NewVal) Then ws.Cells(r, 7).Value = .NewVal
            If Not IsEmpty(.OldVal) And Not IsEmpty(.NewVal) Then
                ws.Cells(r, 8).Value = CDbl(.NewVal) - CDbl(.OldVal)
            End If
            ws.Cells(r, 9).Value = .Judge
            ws.Cells(r, 10).Value = .Source
            ws.Cells(r, 11).Value = .Message

            Select Case .Judge
                Case "差異":   ws.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
                Case "転記":   ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
                Case "エラー": ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
                Case "手入力": ws.Cells(r, 9).Interior.Color = RGB(226, 226, 226)
            End Select
        End With
        r = r + 1
    Next i

    ws.Rows(3).AutoFilter
    ws.Columns.AutoFit
    If ws.Columns(11).ColumnWidth > 60 Then ws.Columns(11).ColumnWidth = 60
    ws.Activate
    ws.Range("A4").Select
End Sub

Private Function Summary(ByRef results() As TResult, ByVal n As Long, _
                         ByVal wrote As Boolean) As String
    Dim i As Long
    Dim nOK As Long, nDiff As Long, nWr As Long, nErr As Long, nMan As Long

    For i = 0 To n - 1
        Select Case results(i).Judge
            Case "一致":   nOK = nOK + 1
            Case "差異":   nDiff = nDiff + 1
            Case "転記":   nWr = nWr + 1
            Case "エラー": nErr = nErr + 1
            Case "手入力": nMan = nMan + 1
        End Select
    Next i

    Summary = "一致  : " & nOK & " 件" & vbCrLf
    If wrote Then
        Summary = Summary & "転記  : " & nWr & " 件" & vbCrLf
    Else
        Summary = Summary & "差異  : " & nDiff & " 件" & vbCrLf
    End If
    Summary = Summary & "手入力: " & nMan & " 件（照合対象外）" & vbCrLf & _
                        "エラー: " & nErr & " 件" & vbCrLf & vbCrLf & _
                        "詳しくは「" & REPORT_SHEET & "」シートを見てください。"
    If nErr > 0 Then
        Summary = Summary & vbCrLf & vbCrLf & _
                  "※エラーの多くは、転記元ファイルが見つからないか" & vbCrLf & _
                  "　シート名・キーが変わったことが原因です。"
    End If
End Function
