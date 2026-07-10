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

    ' 出力先を初期化
    wsDest.Cells.Clear

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

    wsDest.Columns.AutoFit

    MsgBox (outRow - 3) & " 件の行を「" & DEST_SHEET_NAME & "」シートにリンク（数式）でコピーしました。", vbInformation

End Sub
