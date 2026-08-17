Option Explicit

' コードの版数。貼り替え忘れの確認用に、更新のたびに増やす。
' 実行後のメッセージボックスにこの番号が表示される。
Const MACRO_VERSION As String = "v46"

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

    ' 実行するたびに「再生資源」シートの内容を全て消してから作り直す
    wsDest.Cells.Clear

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
    ' （シート全体を実行冒頭で消しているので、毎回まっさらな状態に書く）
    Dim extraCol As Long
    extraCol = lastColSrc + 1

    Dim extraHeaders(1 To 19, 1 To 2) As String
    extraHeaders(1, 1) = "": extraHeaders(1, 2) = "単位Co量(m3/施工単位)"
    extraHeaders(2, 1) = "": extraHeaders(2, 2) = "Co量(m3)"
    extraHeaders(3, 1) = "": extraHeaders(3, 2) = "処分無筋Co量(m3)"
    extraHeaders(4, 1) = "": extraHeaders(4, 2) = "処分鉄筋Co量(m3)"
    extraHeaders(5, 1) = "粗粒度": extraHeaders(5, 2) = "単位As量(t/m2)"
    extraHeaders(6, 1) = "": extraHeaders(6, 2) = "As量(t)"
    extraHeaders(7, 1) = "密粒度": extraHeaders(7, 2) = "単位As量(t/m2)"
    extraHeaders(8, 1) = "": extraHeaders(8, 2) = "As量(t)"
    extraHeaders(9, 1) = "細粒度": extraHeaders(9, 2) = "単位As量(t/m2)"
    extraHeaders(10, 1) = "": extraHeaders(10, 2) = "As量(t)"
    extraHeaders(11, 1) = "開粒度": extraHeaders(11, 2) = "単位As量(t/m2)"
    extraHeaders(12, 1) = "": extraHeaders(12, 2) = "As量(t)"
    extraHeaders(13, 1) = "改質アスコン": extraHeaders(13, 2) = "単位As量(t/m2)"
    extraHeaders(14, 1) = "": extraHeaders(14, 2) = "As量(t)"
    extraHeaders(15, 1) = "": extraHeaders(15, 2) = "処分As量(t)"
    extraHeaders(16, 1) = "": extraHeaders(16, 2) = "砕石量(m3/m2)"
    extraHeaders(17, 1) = "": extraHeaders(17, 2) = "粒調砕石量(m3/m2)"
    extraHeaders(18, 1) = "": extraHeaders(18, 2) = "掘削土量(m3)"
    extraHeaders(19, 1) = "": extraHeaders(19, 2) = "処分土量(m3)"

    For i = 1 To 19
        If Len(extraHeaders(i, 1)) > 0 Then wsDest.Cells(1, extraCol + i - 1).Value = extraHeaders(i, 1)
        If Len(extraHeaders(i, 2)) > 0 Then wsDest.Cells(2, extraCol + i - 1).Value = extraHeaders(i, 2)
    Next i

    ' 材料数量集計用見出し（19列、extraCol～extraCol+18）のすぐ右の
    ' 固定位置に、単位数量の見出しを縦に記載する。列位置を毎回スキャンで
    ' 探すのではなく固定にすることで、既存データの内容に影響されず、
    ' 再実行しても列がずれたり増えたりしない
    Dim unitLabels As Variant
    unitLabels = Array("単位Co量(m3/施工単位)", "粗粒度単位As量(t/m2)", "密粒度単位As量(t/m2)", _
                        "細粒度単位As量(t/m2)", "開粒度単位As量(t/m2)", "改質アスコン単位As量(t/m2)", _
                        "処分As単位量(t/m3)")

    Dim unitLabelCol As Long
    unitLabelCol = extraCol + 19

    Dim labelRow As Long
    labelRow = 2
    For i = LBound(unitLabels) To UBound(unitLabels)
        wsDest.Cells(labelRow + i, unitLabelCol).Value = unitLabels(i)
    Next i

    ' 一番端の「単位Co量」セルの横に、数値・単位・参照先を記載
    wsDest.Cells(labelRow, unitLabelCol + 1).Value = "0.138"
    wsDest.Cells(labelRow, unitLabelCol + 2).Value = "m3/m"
    wsDest.Cells(labelRow, unitLabelCol + 3).Value = "標準図 NG-L-FA参照"

    ' 「改質アスコン単位As量」の下の「処分As単位量(t/m3)」セルの横に、
    ' 数値・単位・注記を記載
    wsDest.Cells(labelRow + 6, unitLabelCol + 1).Value = "2.30"
    wsDest.Cells(labelRow + 6, unitLabelCol + 2).Value = "t/m3"
    wsDest.Cells(labelRow + 6, unitLabelCol + 3).Value = "(歩車道密粒細粒などの平均として)"

    ' D列に「街渠工」を含む行の「単位Co量」列（extraCol）に、
    ' 上で記載した0.138セルを絶対参照する数式を入れる
    Dim unitCoRefAddr As String
    unitCoRefAddr = wsDest.Cells(labelRow, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

    ' 同じ行の「単位Co量」列（extraCol）の右隣（Co量(m3)列）に、
    ' 単位Co量×L列 の数式を入れる
    For r = 3 To outRow - 1
        If InStr(1, CStr(wsDest.Cells(r, 4).Value), "街渠工", vbTextCompare) > 0 Then
            wsDest.Cells(r, extraCol).Formula = "=" & unitCoRefAddr
            wsDest.Cells(r, extraCol + 1).Formula = "=" & wsDest.Cells(r, extraCol).Address(False, False) & _
                                                     "*" & wsDest.Cells(r, 12).Address(False, False)
        End If
    Next r

    ' D列に「殻運搬処理」を含み、かつE列またはI列に「Co」を含む行の
    ' 「処分無筋Co量(m3)」列に、その行のL列を絶対参照する数式を入れる
    Dim disposalConcreteCol As Long
    disposalConcreteCol = FindHeaderColumn(wsDest, extraCol, "処分無筋Co量(m3)")

    If disposalConcreteCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "殻運搬処理", vbTextCompare) > 0 And _
               (InStr(1, StrConv(CStr(wsDest.Cells(r, 5).Value), vbNarrow), "Co", vbTextCompare) > 0 Or _
                InStr(1, StrConv(CStr(wsDest.Cells(r, 9).Value), vbNarrow), "Co", vbTextCompare) > 0) Then
                wsDest.Cells(r, disposalConcreteCol).Formula = "=$L$" & r
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつ（G列に「２－２号工」もしくは「３号工」、
    ' またはE列に「10-1号工(乗入部)」を含む）行の「粗粒度」列
    ' （１行目=粗粒度、２行目=単位As量(t/m2)）に、粗粒度単位As量の
    ' 0.115セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim roughAsCol As Long
    roughAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "粗粒度", "単位As量(t/m2)")

    If roughAsCol > 0 Then
        Dim roughAsRefAddr As String
        roughAsRefAddr = wsDest.Cells(labelRow + 1, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "2-2号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "3号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "10-1号工(乗入部)", vbTextCompare) > 0) Then
                wsDest.Cells(r, roughAsCol).Formula = "=" & roughAsRefAddr
                wsDest.Cells(r, roughAsCol + 1).Formula = "=" & wsDest.Cells(r, roughAsCol).Address(False, False) & _
                                                           "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「３号工」を含む行の「密粒度」列
    ' （１行目=密粒度、２行目=単位As量(t/m2)）に、密粒度単位As量の
    ' 0.118セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim fineAsCol As Long
    fineAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "密粒度", "単位As量(t/m2)")

    If fineAsCol > 0 Then
        Dim fineAsRefAddr As String
        fineAsRefAddr = wsDest.Cells(labelRow + 2, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "3号工", vbTextCompare) > 0 Then
                wsDest.Cells(r, fineAsCol).Formula = "=" & fineAsRefAddr
                wsDest.Cells(r, fineAsCol + 1).Formula = "=" & wsDest.Cells(r, fineAsCol).Address(False, False) & _
                                                          "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「５号工」もしくは「１０号工」を含む
    ' 行の「細粒度」列（１行目=細粒度、２行目=単位As量(t/m2)）に、細粒度単位As量の
    ' 0.115セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim fineFineAsCol As Long
    fineFineAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "細粒度", "単位As量(t/m2)")

    If fineFineAsCol > 0 Then
        Dim fineFineAsRefAddr As String
        fineFineAsRefAddr = wsDest.Cells(labelRow + 3, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "5号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "10号工", vbTextCompare) > 0) Then
                wsDest.Cells(r, fineFineAsCol).Formula = "=" & fineFineAsRefAddr

                ' 右隣（As量(t)列）に、G列の記載内容に応じた係数付きの数式を入れる
                Dim fineFineG As String
                fineFineG = NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value))

                Dim fineFineAjAddr As String, fineFineLAddr As String
                fineFineAjAddr = wsDest.Cells(r, fineFineAsCol).Address(False, False)
                fineFineLAddr = wsDest.Cells(r, 12).Address(False, False)

                If InStr(1, fineFineG, "4cm", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr & "*4/5"
                ElseIf InStr(1, fineFineG, "5cm", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr
                ElseIf InStr(1, fineFineG, "10cm", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr & "*2"
                ElseIf InStr(1, fineFineG, "10号工", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr & "*4/5"
                End If
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「２－２号工」もしくは「１０－１号工」を
    ' 含む行の「開粒度」列（１行目=開粒度、２行目=単位As量(t/m2)）に、開粒度単位As量の
    ' 0.097セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim coarseAsCol As Long
    coarseAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "開粒度", "単位As量(t/m2)")

    If coarseAsCol > 0 Then
        Dim coarseAsRefAddr As String
        coarseAsRefAddr = wsDest.Cells(labelRow + 4, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "2-2号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "10-1号工", vbTextCompare) > 0) Then
                wsDest.Cells(r, coarseAsCol).Formula = "=" & coarseAsRefAddr

                ' 右隣（As量(t)列）に、G列の記載内容に応じた係数付きの数式を入れる
                Dim coarseG As String
                coarseG = NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value))

                Dim coarseAlAddr As String, coarseLAddr As String
                coarseAlAddr = wsDest.Cells(r, coarseAsCol).Address(False, False)
                coarseLAddr = wsDest.Cells(r, 12).Address(False, False)

                If InStr(1, coarseG, "2-2号工", vbTextCompare) > 0 Then
                    wsDest.Cells(r, coarseAsCol + 1).Formula = "=" & coarseAlAddr & "*" & coarseLAddr & "*2"
                ElseIf InStr(1, coarseG, "10-1号工", vbTextCompare) > 0 Then
                    wsDest.Cells(r, coarseAsCol + 1).Formula = "=" & coarseAlAddr & "*" & coarseLAddr & "*4/5"
                End If
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「７号工」を含む行の
    ' 「Co量(m3)」列に、＝L列×0.15 の数式を入れる（全角/半角の表記ゆれを吸収）
    Dim coQuantityCol As Long
    coQuantityCol = FindHeaderColumn(wsDest, extraCol, "Co量(m3)")

    If coQuantityCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "7号工", vbTextCompare) > 0 Then
                wsDest.Cells(r, coQuantityCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False) & "*0.15"
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつE列に「１０－１号工」と「乗入」の両方を
    ' 含む行の「改質アスコン」列（１行目=改質アスコン、２行目=単位As量(t/m2)）に、
    ' 改質アスコン単位As量の0.115セルを絶対参照する数式を入れる
    ' （全角/半角・ダッシュの表記ゆれを吸収）
    Dim modifiedAsCol As Long
    modifiedAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "改質アスコン", "単位As量(t/m2)")

    If modifiedAsCol > 0 Then
        Dim modifiedAsRefAddr As String
        modifiedAsRefAddr = wsDest.Cells(labelRow + 5, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "10-1号工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "乗入", vbTextCompare) > 0 Then
                wsDest.Cells(r, modifiedAsCol).Formula = "=" & modifiedAsRefAddr
                wsDest.Cells(r, modifiedAsCol + 1).Formula = "=" & wsDest.Cells(r, modifiedAsCol).Address(False, False) & _
                                                              "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「殻運搬処理」を含み、かつE列・F列・G列・H列のいずれかに「As」を
    ' 含む行の「処分As量(t)」列に、＝L列 の数式を入れる（全角/半角の表記ゆれを吸収）
    Dim disposalAsCol As Long
    disposalAsCol = FindHeaderColumn(wsDest, extraCol, "処分As量(t)")

    If disposalAsCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "殻運搬処理", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "As", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 6).Value)), "As", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "As", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 8).Value)), "As", vbTextCompare) > 0) Then
                wsDest.Cells(r, disposalAsCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' 粗粒度/密粒度/細粒度/開粒度/改質アスコン、各単位As量セルの横に、数値・単位・注記を記載
    Dim asValues As Variant
    asValues = Array("0.115", "0.118", "0.115", "0.097", "0.115")

    For i = 1 To 5
        wsDest.Cells(labelRow + i, unitLabelCol + 1).Value = asValues(i - 1)
        wsDest.Cells(labelRow + i, unitLabelCol + 2).Value = "t/m2"
        wsDest.Cells(labelRow + i, unitLabelCol + 3).Value = "t=5cm"
    Next i

    wsDest.Columns.AutoFit

    MsgBox (outRow - 3) & " 件の行を「" & DEST_SHEET_NAME & "」シートにリンク（数式）でコピーしました。" & vbCrLf & _
           "(" & MACRO_VERSION & ")", vbInformation

End Sub

' ２行目が headerText と一致する列番号を返す。見つからなければ0を返す。
Private Function FindHeaderColumn(ws As Worksheet, searchFromCol As Long, headerText As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(2, c).Value) = headerText Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

' １行目が row1Text、２行目が row2Text と一致する列番号を返す。
' （"単位As量(t/m2)"のように２行目だけでは複数該当する見出しを区別するため）
' 見つからなければ0を返す。
Private Function FindHeaderColumnByRow1Row2(ws As Worksheet, searchFromCol As Long, _
                                             row1Text As String, row2Text As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(1, c).Value) = row1Text And CStr(ws.Cells(2, c).Value) = row2Text Then
            FindHeaderColumnByRow1Row2 = c
            Exit Function
        End If
    Next c
    FindHeaderColumnByRow1Row2 = 0
End Function

' 全角英数字・記号を半角に揃え、長音記号やダッシュ類も半角ハイフンに
' 統一してから比較できるようにする（表記ゆれ対策）。
' StrConv(vbNarrow)は環境によって全角英数字を変換しないことがあるため、
' 文字コード（Unicode）を直接シフトして半角化する。
Private Function NormalizeForMatch(ByVal s As String) As String
    Dim result As String
    result = ""

    Dim i As Long, code As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch)
        If code >= &HFF01 And code <= &HFF5E Then
            ' 全角の！～～（英数字・記号を含む）を半角へ
            ch = ChrW(code - &HFEE0)
        ElseIf code = &H3000 Then
            ' 全角スペース -> 半角スペース
            ch = " "
        End If
        result = result & ch
    Next i

    result = Replace(result, "‐", "-")
    result = Replace(result, "‑", "-")
    result = Replace(result, "–", "-")
    result = Replace(result, "—", "-")
    result = Replace(result, "ー", "-")
    result = Replace(result, "ｰ", "-")

    ' cm/m2/m3などの単位記号（CJK互換文字の１文字表記）を通常の英字表記に展開
    result = Replace(result, "㎝", "cm")
    result = Replace(result, "㎜", "mm")
    result = Replace(result, "㎡", "m2")
    result = Replace(result, "㎥", "m3")

    NormalizeForMatch = result
End Function
