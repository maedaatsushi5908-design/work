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
        extraHeaders(1, 1) = "": extraHeaders(1, 2) = "単位Co量(m3/施工単位)"
        extraHeaders(2, 1) = "": extraHeaders(2, 2) = "Co量(m3)"
        extraHeaders(3, 1) = "": extraHeaders(3, 2) = "処分Co量(m3)"
        extraHeaders(4, 1) = "粗粒度": extraHeaders(4, 2) = "単位As量(t/m2)"
        extraHeaders(5, 1) = "": extraHeaders(5, 2) = "As量(t)"
        extraHeaders(6, 1) = "密粒度": extraHeaders(6, 2) = "単位As量(t/m2)"
        extraHeaders(7, 1) = "": extraHeaders(7, 2) = "As量(t)"
        extraHeaders(8, 1) = "細粒度": extraHeaders(8, 2) = "単位As量(t/m2)"
        extraHeaders(9, 1) = "": extraHeaders(9, 2) = "As量(t)"
        extraHeaders(10, 1) = "開粒度": extraHeaders(10, 2) = "単位As量(t/m2)"
        extraHeaders(11, 1) = "": extraHeaders(11, 2) = "As量(t)"
        extraHeaders(12, 1) = "改質アスコン": extraHeaders(12, 2) = "単位As量(t/m2)"
        extraHeaders(13, 1) = "": extraHeaders(13, 2) = "As量(t)"
        extraHeaders(14, 1) = "": extraHeaders(14, 2) = "処分As量(t)"
        extraHeaders(15, 1) = "": extraHeaders(15, 2) = "砕石量(m3/m2)"
        extraHeaders(16, 1) = "": extraHeaders(16, 2) = "粒調砕石量(m3/m2)"
        extraHeaders(17, 1) = "": extraHeaders(17, 2) = "掘削土量(m3)"
        extraHeaders(18, 1) = "": extraHeaders(18, 2) = "処分土量(m3)"

        For i = 1 To 18
            If Len(extraHeaders(i, 1)) > 0 Then wsDest.Cells(1, extraCol + i - 1).Value = extraHeaders(i, 1)
            If Len(extraHeaders(i, 2)) > 0 Then wsDest.Cells(2, extraCol + i - 1).Value = extraHeaders(i, 2)
        Next i
    Else
        ' 既に見出しがある古いレイアウトのシートには、不足している見出し列
        ' だけを実際に挿入して自動的に補う（列を挿入するので、３行目以降に
        ' 手入力された数量データも見出しと一緒に正しくずれる）
        InsertMissingHeaderColumn wsDest, extraCol, "処分Co量(m3)", "単位As量(t/m2)"
        InsertMissingHeaderColumn wsDest, extraCol, "処分As量(t)", "単位砕石量(m3/m2)"

        ' 不要になった「砕石」「粒調砕石」の見出し列は自動的に削除する
        ' （列を削除するので、右側の見出し・データも自動的に詰まる）
        DeleteHeaderColumnIfPresent wsDest, extraCol, "砕石", "単位砕石量(m3/m2)"
        DeleteHeaderColumnIfPresent wsDest, extraCol, "粒調砕石", "単位粒調砕石量(m3/m2)"
    End If

    ' 材料数量集計用見出し（18列、extraCol～extraCol+17）のすぐ右の
    ' 固定位置に、単位数量の見出しを縦に記載する。列位置を毎回スキャンで
    ' 探すのではなく固定にすることで、既存データの内容に影響されず、
    ' 再実行しても列がずれたり増えたりしない
    Dim unitLabels As Variant
    unitLabels = Array("単位Co量(m3/施工単位)", "粗粒度単位As量(t/m2)", "密粒度単位As量(t/m2)", _
                        "細粒度単位As量(t/m2)", "開粒度単位As量(t/m2)", "改質アスコン単位As量(t/m2)")

    Dim unitLabelCol As Long
    unitLabelCol = extraCol + 18

    Dim labelRow As Long
    labelRow = 2
    For i = LBound(unitLabels) To UBound(unitLabels)
        wsDest.Cells(labelRow + i, unitLabelCol).Value = unitLabels(i)
    Next i

    ' 一番端の「単位Co量」セルの横に、数値・単位・参照先を記載
    wsDest.Cells(labelRow, unitLabelCol + 1).Value = "0.138"
    wsDest.Cells(labelRow, unitLabelCol + 2).Value = "m3/m"
    wsDest.Cells(labelRow, unitLabelCol + 3).Value = "標準図 NG-L-FA参照"

    ' D列に「街渠工」を含む行の「単位Co量」列（extraCol）に、
    ' 上で記載した0.138セルを絶対参照する数式を入れる
    Dim unitCoRefAddr As String
    unitCoRefAddr = wsDest.Cells(labelRow, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

    For r = 3 To outRow - 1
        If InStr(1, CStr(wsDest.Cells(r, 4).Value), "街渠工", vbTextCompare) > 0 Then
            wsDest.Cells(r, extraCol).Formula = "=" & unitCoRefAddr
        End If
    Next r

    ' 粗粒度/密粒度/細粒度/開粒度/改質アスコン、各単位As量セルの横に、数値・単位・注記を記載
    Dim asValues As Variant
    asValues = Array("0.115", "0.118", "0.115", "0.097", "0.115")

    For i = 1 To 5
        wsDest.Cells(labelRow + i, unitLabelCol + 1).Value = asValues(i - 1)
        wsDest.Cells(labelRow + i, unitLabelCol + 2).Value = "t/m2"
        wsDest.Cells(labelRow + i, unitLabelCol + 3).Value = "t=5cm"
    Next i

    wsDest.Columns.AutoFit

    MsgBox (outRow - 3) & " 件の行を「" & DEST_SHEET_NAME & "」シートにリンク（数式）でコピーしました。", vbInformation

End Sub

' ２行目に newHeaderText が既に無ければ、beforeHeaderText と書かれた列の
' 手前に新しい列を挿入して newHeaderText を書き込む。
' 列挿入なので、その右側の見出し・データは自動的に１列分ずれる。
Private Sub InsertMissingHeaderColumn(ws As Worksheet, searchFromCol As Long, _
                                       newHeaderText As String, beforeHeaderText As String)
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(2, c).Value) = newHeaderText Then
            Exit Sub
        End If
    Next c

    Dim targetCol As Long
    targetCol = 0
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(2, c).Value) = beforeHeaderText Then
            targetCol = c
            Exit For
        End If
    Next c
    If targetCol = 0 Then Exit Sub

    ws.Columns(targetCol).Insert Shift:=xlToRight
    ws.Cells(2, targetCol).Value = newHeaderText
End Sub

' １行目が headerRow1Text、２行目が headerRow2Text の列が見つかれば、
' その列ごと削除する（右側の見出し・データは自動的に詰まる）。
Private Sub DeleteHeaderColumnIfPresent(ws As Worksheet, searchFromCol As Long, _
                                         headerRow1Text As String, headerRow2Text As String)
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(1, c).Value) = headerRow1Text And CStr(ws.Cells(2, c).Value) = headerRow2Text Then
            ws.Columns(c).Delete Shift:=xlToLeft
            Exit Sub
        End If
    Next c
End Sub
