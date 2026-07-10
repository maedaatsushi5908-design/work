Attribute VB_Name = "CopyRecycledResourceLinks"
Option Explicit

' 「ファイル名」シートの２行目で「施工単価名称」列を探し、
' そのセルの文字列に指定キーワードを含む行のリンクを
' 「再生資源」シート（無ければ新規作成）へコピーする。
Sub CopyRecycledResourceLinks()

    Const SRC_SHEET_NAME As String = "ファイル名"
    Const DEST_SHEET_NAME As String = "再生資源"
    Const HEADER_ROW As Long = 2
    Const HEADER_TEXT As String = "施工単価名称"

    Dim keywords As Variant
    keywords = Array("街渠工", "舗装復旧工", "先行路盤工", "殻運搬処理", _
                      "構造物とりこわし", "床掘", "土砂等運搬")

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = wb.Worksheets(SRC_SHEET_NAME)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "シート「" & SRC_SHEET_NAME & "」が見つかりません。処理を中止します。", vbExclamation
        Exit Sub
    End If

    Dim wsDest As Worksheet
    On Error Resume Next
    Set wsDest = wb.Worksheets(DEST_SHEET_NAME)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDest.Name = DEST_SHEET_NAME
    End If

    ' ２行目から「施工単価名称」列を検索
    Dim lastColSrc As Long
    lastColSrc = wsSrc.Cells(HEADER_ROW, wsSrc.Columns.Count).End(xlToLeft).Column

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
    wsDest.Cells(1, 1).Value = HEADER_TEXT
    wsDest.Cells(1, 2).Value = "リンク"

    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, targetCol).End(xlUp).Row

    Dim outRow As Long
    outRow = 2

    Dim r As Long, i As Long
    Dim cellText As String
    Dim matched As Boolean
    Dim hl As Hyperlink

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
                wsDest.Cells(outRow, 1).Value = cellText

                If wsSrc.Cells(r, targetCol).Hyperlinks.Count > 0 Then
                    Set hl = wsSrc.Cells(r, targetCol).Hyperlinks(1)
                    wsDest.Hyperlinks.Add Anchor:=wsDest.Cells(outRow, 2), _
                                           Address:=hl.Address, _
                                           SubAddress:=hl.SubAddress, _
                                           TextToDisplay:=cellText
                Else
                    wsDest.Cells(outRow, 2).Value = "(リンクなし)"
                End If

                outRow = outRow + 1
            End If
        End If
    Next r

    wsDest.Columns("A:B").AutoFit

    MsgBox (outRow - 2) & " 件のリンクを「" & DEST_SHEET_NAME & "」シートにコピーしました。", vbInformation

End Sub
