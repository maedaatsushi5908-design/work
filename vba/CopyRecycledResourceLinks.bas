Attribute VB_Name = "CopyRecycledResourceLinks"
Option Explicit

' 「ファイル名」シートの２行目で「施工単価名称」列を探し、
' そのセルの文字列に指定キーワードを含む行を丸ごと、
' セル参照の数式（=ファイル名!$A$5 形式）として
' 「再生資源」シート（無ければ新規作成）へコピーする。
' 数式で参照するため、元データを変更すると自動的に反映される。
Sub CopyRecycledResourceLinks()

    Const SRC_SHEET_NAME As String = "ファイル名"
    Const DEST_SHEET_NAME As String = "再生資源"
    Const HEADER_ROW As Long = 2
    Const HEADER_TEXT As String = "施工単価名称"

    Dim keywords As Variant
    keywords = Array("街渠工", "舗装復旧工", "先行路盤工", "殻運搬処理", _
                      "構造物とりこわし", "床掘", "土砂等運搬")

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim wsSrc As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Trim(ws.Name) = SRC_SHEET_NAME Then
            Set wsSrc = ws
            Exit For
        End If
    Next ws

    If wsSrc Is Nothing Then
        Dim sheetList As String
        For Each ws In wb.Worksheets
            sheetList = sheetList & "・" & ws.Name & vbCrLf
        Next ws
        MsgBox "シート「" & SRC_SHEET_NAME & "」が見つかりません。処理を中止します。" & vbCrLf & vbCrLf & _
               "「" & wb.Name & "」ブック内の実際のシート名:" & vbCrLf & sheetList, vbExclamation
        Exit Sub
    End If

    Dim wsDest As Worksheet
    For Each ws In wb.Worksheets
        If Trim(ws.Name) = DEST_SHEET_NAME Then
            Set wsDest = ws
            Exit For
        End If
    Next ws

    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDest.Name = DEST_SHEET_NAME
    End If

    ' １行目・２行目のうち、より右まで使われている方に合わせて列数を決定
    Dim lastColSrc As Long
    Dim lastColRow1 As Long
    lastColRow1 = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    lastColSrc = wsSrc.Cells(HEADER_ROW, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastColRow1 > lastColSrc Then lastColSrc = lastColRow1

    Dim targetCol As Long
    targetCol = 0

    Dim c As Long
    For c = 1 To lastColSrc
        If InStr(1, CStr(wsSrc.Cells(HEADER_ROW, c).Value), HEADER_TEXT, vbTextCompare) > 0 Then
            targetCol = c
            Exit For
        End If
    Next c

    If targetCol = 0 Then
        MsgBox "シート「" & SRC_SHEET_NAME & "」の" & HEADER_ROW & "行目に「" & _
               HEADER_TEXT & "」列が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 出力先を初期化（元データを貼り直す範囲のみクリアし、
    ' 右側に追記した集計用見出し・手入力データは消さない）
    wsDest.Range(wsDest.Columns(1), wsDest.Columns(lastColSrc)).Clear

    ' １行目・２行目（見出し）を数式リンクでコピー
    For c = 1 To lastColSrc
        wsDest.Cells(1, c).Formula = "='" & SRC_SHEET_NAME & "'!" & wsSrc.Cells(1, c).Address
        wsDest.Cells(2, c).Formula = "='" & SRC_SHEET_NAME & "'!" & wsSrc.Cells(HEADER_ROW, c).Address
    Next c

    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, targetCol).End(xlUp).Row

    Dim outRow As Long
    outRow = 3

    Dim r As Long, i As Long
    Dim cellText As String
    Dim matched As Boolean

    For r = HEADER_ROW + 1 To lastRowSrc
        cellText = CStr(wsSrc.Cells(r, targetCol).Value)
        If Len(cellText) > 0 Then

            matched = False
            For i = LBound(keywords) To UBound(keywords)
                If InStr(1, cellText, CStr(keywords(i)), vbTextCompare) > 0 Then
                    matched = True
                    Exit For
                End If
            Next i

            If matched Then
                ' 行全体（全列）をセル参照の数式でコピー
                For c = 1 To lastColSrc
                    wsDest.Cells(outRow, c).Formula = "='" & SRC_SHEET_NAME & "'!" & wsSrc.Cells(r, c).Address
                Next c
                outRow = outRow + 1
            End If
        End If
    Next r

    ' データ列の右端の次から、材料数量集計用の見出しを追記
    ' （既に追記済みの場合は、手入力データを消さないよう追記しない）
    Dim extraCol As Long
    extraCol = lastColSrc + 1

    If Len(Trim(CStr(wsDest.Cells(1, extraCol).Value))) = 0 And _
       Len(Trim(CStr(wsDest.Cells(2, extraCol).Value))) = 0 Then

        Dim extraHeaders(1 To 18, 1 To 2) As String
        extraHeaders(1, 1) = "単位Co量(m3/施工単位)": extraHeaders(1, 2) = ""
        extraHeaders(2, 1) = "Co量(m3)": extraHeaders(2, 2) = ""
        extraHeaders(3, 1) = "粗粒度": extraHeaders(3, 2) = "単位As量(t/m2)"
        extraHeaders(4, 1) = "": extraHeaders(4, 2) = "As量(t)"
        extraHeaders(5, 1) = "密粒度": extraHeaders(5, 2) = "単位As量(t/m2)"
        extraHeaders(6, 1) = "": extraHeaders(6, 2) = "As量(t)"
        extraHeaders(7, 1) = "細粒度": extraHeaders(7, 2) = "単位As量(t/m2)"
        extraHeaders(8, 1) = "": extraHeaders(8, 2) = "As量(t)"
        extraHeaders(9, 1) = "開粒度": extraHeaders(9, 2) = "単位As量(t/m2)"
        extraHeaders(10, 1) = "": extraHeaders(10, 2) = "As量(t)"
        extraHeaders(11, 1) = "改質アスコン": extraHeaders(11, 2) = "単位As量(t/m2)"
        extraHeaders(12, 1) = "": extraHeaders(12, 2) = "As量(t)"
        extraHeaders(13, 1) = "砕石": extraHeaders(13, 2) = "単位砕石量(m3/m2)"
        extraHeaders(14, 1) = "": extraHeaders(14, 2) = "砕石量(m3/m2)"
        extraHeaders(15, 1) = "粒調砕石": extraHeaders(15, 2) = "単位粒調砕石量(m3/m2)"
        extraHeaders(16, 1) = "": extraHeaders(16, 2) = "粒調砕石量(m3/m2)"
        extraHeaders(17, 1) = "掘削土量(m3)": extraHeaders(17, 2) = ""
        extraHeaders(18, 1) = "処分土量(m3)": extraHeaders(18, 2) = ""

        For i = 1 To 18
            If Len(extraHeaders(i, 1)) > 0 Then wsDest.Cells(1, extraCol + i - 1).Value = extraHeaders(i, 1)
            If Len(extraHeaders(i, 2)) > 0 Then wsDest.Cells(2, extraCol + i - 1).Value = extraHeaders(i, 2)
        Next i
    End If

    wsDest.Columns.AutoFit

    MsgBox (outRow - 3) & " 件の行を「" & DEST_SHEET_NAME & "」シートにリンク（数式）でコピーしました。", vbInformation

End Sub
