Attribute VB_Name = "CompareExcelP"
Option Explicit

'==========================================================
' 2つのExcelブックを比較するマクロ
'  ・各ブックで比較するシートを選択（シート名が違ってもよい）
'  ・D/E/F/O列がすべて一致する行（行番号は違ってよい）を探し、
'    その行どうしでP列が一致しているかを判定
'  ・結果は新しいブックに一覧出力
'==========================================================

Private Const START_ROW As Long = 2        ' データ開始行（1行目が見出しの場合は2）
Private Const TRIM_VALUES As Boolean = True ' 前後の空白を無視して比較する

Public Sub CompareTwoWorkbooks()
    Dim path1 As Variant, path2 As Variant
    Dim wb1 As Workbook, wb2 As Workbook
    Dim wbOut As Workbook, wsOut As Worksheet
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim outRow As Long
    Dim cntOK As Long, cntNG As Long, cntNo1 As Long, cntNo2 As Long
    Dim info1 As String, info2 As String

    path1 = Application.GetOpenFilename("Excelファイル,*.xls;*.xlsx;*.xlsm", , "1つ目のExcelを選択")
    If path1 = False Then Exit Sub
    path2 = Application.GetOpenFilename("Excelファイル,*.xls;*.xlsx;*.xlsm", , "2つ目のExcelを選択")
    If path2 = False Then Exit Sub

    Application.DisplayAlerts = False
    Set wb1 = Workbooks.Open(path1, ReadOnly:=True)
    Set wb2 = Workbooks.Open(path2, ReadOnly:=True)
    Application.DisplayAlerts = True

    ' 比較するシートをそれぞれ選択
    Set ws1 = SelectSheet(wb1, "ブック1")
    If ws1 Is Nothing Then GoTo Cleanup
    Set ws2 = SelectSheet(wb2, "ブック2")
    If ws2 Is Nothing Then GoTo Cleanup

    info1 = wb1.Name & " [" & ws1.Name & "]"
    info2 = wb2.Name & " [" & ws2.Name & "]"

    Application.ScreenUpdating = False

    Set wbOut = Workbooks.Add
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = "比較結果"
    wsOut.Range("A1:K1").Value = Array("シート名", "判定", "ブック1 行", "ブック2 行", _
        "D列", "E列", "F列", "O列", "ブック1 P列", "ブック2 P列", "備考")
    outRow = 2

    CompareSheets ws1, ws2, wsOut, outRow, cntOK, cntNG, cntNo1, cntNo2

    With wsOut
        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Interior.Color = RGB(221, 235, 247)
        .Columns("A:K").AutoFit
        .Range("A1").AutoFilter
    End With

Cleanup:
    If Not wb2 Is wb1 Then wb2.Close SaveChanges:=False   ' 同じファイルを2回選んだ場合に備える
    wb1.Close SaveChanges:=False
    Application.ScreenUpdating = True
    If wsOut Is Nothing Then Exit Sub

    wbOut.Activate
    wsOut.Range("A2").Select
    ActiveWindow.FreezePanes = True

    MsgBox "比較が完了しました。" & vbCrLf & _
           "ブック1: " & info1 & vbCrLf & _
           "ブック2: " & info2 & vbCrLf & vbCrLf & _
           "P列一致: " & cntOK & " 件" & vbCrLf & _
           "P列不一致: " & cntNG & " 件" & vbCrLf & _
           "ブック2に該当行なし: " & cntNo2 & " 件" & vbCrLf & _
           "ブック1に該当行なし: " & cntNo1 & " 件", vbInformation
End Sub

'----------------------------------------------------------
' 比較するシートを選択させる
' （シートタブをクリックしてそのシートのセルを選び、OKで決定）
' シートが1枚だけのときは自動選択、キャンセル時は Nothing
'----------------------------------------------------------
Private Function SelectSheet(wb As Workbook, label As String) As Worksheet
    Dim rng As Range

    If wb.Worksheets.Count = 1 Then
        Set SelectSheet = wb.Worksheets(1)
        Exit Function
    End If

    wb.Activate
    Do
        Set rng = Nothing
        On Error Resume Next
        Set rng = Application.InputBox( _
            label & "（" & wb.Name & "）で比較するシートのタブをクリックし、" & vbCrLf & _
            "そのシートの任意のセルを選んで OK を押してください。", _
            label & " のシート選択", ActiveCell.Address(External:=True), Type:=8)
        On Error GoTo 0
        If rng Is Nothing Then Exit Function              ' キャンセル

        If rng.Worksheet.Parent Is wb Then
            Set SelectSheet = rng.Worksheet
            Exit Function
        End If
        MsgBox wb.Name & " のシートを選んでください。", vbExclamation
        wb.Activate
    Loop
End Function

'----------------------------------------------------------
' シート同士の比較
'----------------------------------------------------------
Private Sub CompareSheets(ws1 As Worksheet, ws2 As Worksheet, wsOut As Worksheet, _
                          ByRef outRow As Long, ByRef cntOK As Long, ByRef cntNG As Long, _
                          ByRef cntNo1 As Long, ByRef cntNo2 As Long)
    Dim last1 As Long, last2 As Long
    Dim d1 As Variant, d2 As Variant
    Dim dict As Object, used2 As Object
    Dim r As Long, i As Long, key As String
    Dim rows2 As Collection, r2 As Variant
    Dim p1 As String, p2 As String
    Dim sheetLabel As String

    sheetLabel = ws1.Name & " / " & ws2.Name

    last1 = LastRow(ws1)
    last2 = LastRow(ws2)
    If last1 < START_ROW And last2 < START_ROW Then Exit Sub

    ' D～P列をまとめて配列に読み込む（D=1, E=2, F=3, O=12, P=13）
    If last1 >= START_ROW Then d1 = ws1.Range(ws1.Cells(START_ROW, "D"), ws1.Cells(last1, "P")).Value
    If last2 >= START_ROW Then d2 = ws2.Range(ws2.Cells(START_ROW, "D"), ws2.Cells(last2, "P")).Value

    ' ブック2側：キー(D|E|F|O) → 行番号のコレクション
    Set dict = CreateObject("Scripting.Dictionary")
    Set used2 = CreateObject("Scripting.Dictionary")
    If last2 >= START_ROW Then
        For i = 1 To UBound(d2, 1)
            key = MakeKey(d2, i)
            If key <> "" Then
                If Not dict.Exists(key) Then dict.Add key, New Collection
                dict(key).Add i
            End If
        Next i
    End If

    ' ブック1側を1行ずつ照合
    If last1 >= START_ROW Then
        For r = 1 To UBound(d1, 1)
            key = MakeKey(d1, r)
            If key <> "" Then
                If dict.Exists(key) Then
                    used2(key) = True
                    Set rows2 = dict(key)
                    p1 = ToText(d1(r, 13))
                    For Each r2 In rows2
                        p2 = ToText(d2(r2, 13))
                        WriteRow wsOut, outRow, sheetLabel, IIf(p1 = p2, "一致", "不一致"), _
                                 r + START_ROW - 1, r2 + START_ROW - 1, d1, r, d1(r, 13), d2(r2, 13), _
                                 IIf(rows2.Count > 1, "ブック2に同キーが" & rows2.Count & "行あり", "")
                        If p1 = p2 Then
                            cntOK = cntOK + 1
                        Else
                            cntNG = cntNG + 1
                            wsOut.Range(wsOut.Cells(outRow - 1, 1), wsOut.Cells(outRow - 1, 11)).Interior.Color = RGB(255, 199, 206)
                        End If
                    Next r2
                Else
                    WriteRow wsOut, outRow, sheetLabel, "ブック2になし", r + START_ROW - 1, "", _
                             d1, r, d1(r, 13), "", "D/E/F/O列が一致する行がブック2にありません"
                    wsOut.Range(wsOut.Cells(outRow - 1, 1), wsOut.Cells(outRow - 1, 11)).Interior.Color = RGB(255, 235, 156)
                    cntNo2 = cntNo2 + 1
                End If
            End If
        Next r
    End If

    ' ブック2にだけある行
    If last2 >= START_ROW Then
        For i = 1 To UBound(d2, 1)
            key = MakeKey(d2, i)
            If key <> "" Then
                If Not used2.Exists(key) Then
                    WriteRow wsOut, outRow, sheetLabel, "ブック1になし", "", i + START_ROW - 1, _
                             d2, i, "", d2(i, 13), "D/E/F/O列が一致する行がブック1にありません"
                    wsOut.Range(wsOut.Cells(outRow - 1, 1), wsOut.Cells(outRow - 1, 11)).Interior.Color = RGB(255, 235, 156)
                    cntNo1 = cntNo1 + 1
                End If
            End If
        Next i
    End If
End Sub

'----------------------------------------------------------
' 結果1行を書き込み
'----------------------------------------------------------
Private Sub WriteRow(wsOut As Worksheet, ByRef outRow As Long, sheetName As String, judge As String, _
                     row1 As Variant, row2 As Variant, data As Variant, idx As Long, _
                     pVal1 As Variant, pVal2 As Variant, note As String)
    wsOut.Cells(outRow, 1).Value = sheetName
    wsOut.Cells(outRow, 2).Value = judge
    wsOut.Cells(outRow, 3).Value = row1
    wsOut.Cells(outRow, 4).Value = row2
    wsOut.Cells(outRow, 5).Value = "'" & ToText(data(idx, 1))
    wsOut.Cells(outRow, 6).Value = "'" & ToText(data(idx, 2))
    wsOut.Cells(outRow, 7).Value = "'" & ToText(data(idx, 3))
    wsOut.Cells(outRow, 8).Value = "'" & ToText(data(idx, 12))
    wsOut.Cells(outRow, 9).Value = "'" & ToText(pVal1)
    wsOut.Cells(outRow, 10).Value = "'" & ToText(pVal2)
    wsOut.Cells(outRow, 11).Value = note
    outRow = outRow + 1
End Sub

'----------------------------------------------------------
' D/E/F/O列からキーを作る（4列とも空の行は対象外として "" を返す）
'----------------------------------------------------------
Private Function MakeKey(data As Variant, idx As Long) As String
    Dim d As String, e As String, f As String, o As String
    d = ToText(data(idx, 1))
    e = ToText(data(idx, 2))
    f = ToText(data(idx, 3))
    o = ToText(data(idx, 12))
    If d = "" And e = "" And f = "" And o = "" Then
        MakeKey = ""
    Else
        MakeKey = d & vbTab & e & vbTab & f & vbTab & o
    End If
End Function

'----------------------------------------------------------
' セル値を比較用の文字列に変換
'----------------------------------------------------------
Private Function ToText(v As Variant) As String
    If IsError(v) Then
        ToText = "#ERROR"
    ElseIf IsEmpty(v) Then
        ToText = ""
    Else
        ToText = CStr(v)
        If TRIM_VALUES Then ToText = Trim$(ToText)
    End If
End Function

'----------------------------------------------------------
' D/E/F/O/P列のうち一番下のデータ行
'----------------------------------------------------------
Private Function LastRow(ws As Worksheet) As Long
    Dim c As Variant, r As Long
    For Each c In Array("D", "E", "F", "O", "P")
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > LastRow Then LastRow = r
    Next c
End Function
