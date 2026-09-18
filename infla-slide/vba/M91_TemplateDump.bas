'==============================================================
' M91_TemplateDump  ―  テンプレートの体裁をテキストに書き出す
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M91_TemplateDump" に。
'
' ファイルをアップロードできない環境向けのモジュールです。
' インフレスライド設計書のテンプレートを開いて対象シートを表示した状態で
' ［テンプレ構造書き出し］を実行すると、"_構造書き出し" シートに
' セルの値・結合・罫線・列幅・印刷設定がテキストで出力されます。
' そのA列をコピーして、チャットかGitHubに貼り付けてください。
'==============================================================
Option Explicit

Private mLines() As String
Private mCount As Long

' 書き出す最大範囲（これを超える分は切り捨て、末尾に警告を出す）
Private Const MAX_ROW As Long = 300
Private Const MAX_COL As Long = 60


Public Sub テンプレ構造書き出し()
    Dim src As Worksheet, out As Worksheet
    Dim ur As Range
    Dim lastR As Long, lastC As Long
    Dim r As Long, c As Long, i As Long
    Dim cel As Range
    Dim v As String, bc As String
    Dim truncated As Boolean

    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "ワークシートを表示した状態で実行してください。", vbExclamation
        Exit Sub
    End If
    Set src = ActiveSheet

    Application.ScreenUpdating = False
    mCount = 0
    ReDim mLines(0 To 999)

    Set ur = src.UsedRange
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    If lastR > MAX_ROW Then lastR = MAX_ROW: truncated = True
    If lastC > MAX_COL Then lastC = MAX_COL: truncated = True

    AddLine "### SHEET: " & src.Name
    AddLine "### 使用範囲: " & ur.Address(False, False) & _
            "  （書き出し範囲 R1:R" & lastR & " / C1:C" & lastC & "）"
    AddLine ""

    '--- 印刷設定 ---
    AddLine "### 印刷設定"
    On Error Resume Next
    With src.PageSetup
        AddLine "用紙" & vbTab & PaperName(.PaperSize)
        AddLine "向き" & vbTab & IIf(.Orientation = 2, "横", "縦")
        AddLine "余白cm" & vbTab & "上=" & Fmt(.TopMargin / 28.3465) & _
                " 下=" & Fmt(.BottomMargin / 28.3465) & _
                " 左=" & Fmt(.LeftMargin / 28.3465) & _
                " 右=" & Fmt(.RightMargin / 28.3465)
        AddLine "拡大縮小" & vbTab & IIf(.Zoom = False, _
                "横" & .FitToPagesWide & "ページ×縦" & .FitToPagesTall & "ページ", _
                CStr(.Zoom) & "%")
        AddLine "印刷範囲" & vbTab & .PrintArea
        AddLine "タイトル行" & vbTab & .PrintTitleRows
        AddLine "タイトル列" & vbTab & .PrintTitleColumns
        AddLine "ヘッダー" & vbTab & "左=" & .LeftHeader & " 中=" & .CenterHeader & " 右=" & .RightHeader
        AddLine "フッター" & vbTab & "左=" & .LeftFooter & " 中=" & .CenterFooter & " 右=" & .RightFooter
        AddLine "中央揃え" & vbTab & "水平=" & .CenterHorizontally & " 垂直=" & .CenterVertically
    End With
    On Error GoTo 0
    AddLine ""

    '--- 列幅 ---
    v = ""
    For c = 1 To lastC
        v = v & ColLetter(c) & "=" & Fmt(src.Columns(c).ColumnWidth) & " "
    Next c
    AddLine "### 列幅"
    AddLine v
    AddLine ""

    '--- 行高 ---
    v = ""
    For r = 1 To lastR
        v = v & r & "=" & Fmt(src.Rows(r).RowHeight) & " "
    Next r
    AddLine "### 行高"
    AddLine v
    AddLine ""

    '--- 結合セル ---
    AddLine "### 結合セル"
    v = ""
    For r = 1 To lastR
        For c = 1 To lastC
            Set cel = src.Cells(r, c)
            If cel.MergeCells Then
                If cel.MergeArea.Cells(1, 1).Address = cel.Address Then
                    v = v & cel.MergeArea.Address(False, False) & " "
                End If
            End If
        Next c
    Next r
    AddLine v
    AddLine ""

    '--- セル ---
    AddLine "### セル（セル / 値・数式 / 表示形式 / 横位置 / 縦位置 / フォント / サイズ / 太字 / 罫線(上右下左) / 背景色）"
    AddLine "### 罫線コード:  - なし / 1 細 / 2 中 / 3 太 / = 二重 / . 点線"

    For r = 1 To lastR
        For c = 1 To lastC
            Set cel = src.Cells(r, c)
            bc = BorderCode(cel)
            If CellHasContent(cel, bc) Then
                AddLine cel.Address(False, False) & vbTab & _
                        CStr(cel.Formula) & vbTab & _
                        cel.NumberFormatLocal & vbTab & _
                        HAlignName(cel.HorizontalAlignment) & vbTab & _
                        VAlignName(cel.VerticalAlignment) & vbTab & _
                        cel.Font.Name & vbTab & _
                        Fmt(cel.Font.Size) & vbTab & _
                        IIf(cel.Font.Bold, "太字", "") & vbTab & _
                        bc & vbTab & _
                        FillName(cel)
            End If
        Next c
    Next r

    If truncated Then
        AddLine ""
        AddLine "### ※ 範囲が大きいため R" & MAX_ROW & " / C" & MAX_COL & " で打ち切りました"
    End If

    '--- 出力シートへ ---
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("_構造書き出し").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set out = ThisWorkbook.Worksheets.Add( _
                  After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    out.Name = "_構造書き出し"

    out.Columns(1).NumberFormat = "@"
    For i = 0 To mCount - 1
        out.Cells(i + 1, 1).Value = mLines(i)
    Next i

    out.Columns(1).ColumnWidth = 120
    out.Activate
    out.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox "「_構造書き出し」シートに " & mCount & " 行を書き出しました。" & vbCrLf & vbCrLf & _
           "A列を選択（列見出しAをクリック）→ Ctrl+C でコピーして、" & vbCrLf & _
           "チャットまたはGitHubに貼り付けてください。", vbInformation
End Sub


'==============================================================
' 補助
'==============================================================
Private Sub AddLine(ByVal s As String)
    If mCount > UBound(mLines) Then ReDim Preserve mLines(0 To mCount + 999)
    mLines(mCount) = s
    mCount = mCount + 1
End Sub


' 値・罫線・背景色のいずれかがあるセルだけ書き出す
Private Function CellHasContent(ByVal cel As Range, ByVal borderCd As String) As Boolean
    If Len(CStr(cel.Formula)) > 0 Then CellHasContent = True: Exit Function
    If borderCd <> "----" Then CellHasContent = True: Exit Function
    If cel.Interior.ColorIndex <> xlColorIndexNone Then CellHasContent = True
End Function


Private Function BorderCode(ByVal cel As Range) As String
    BorderCode = EdgeCode(cel, xlEdgeTop) & EdgeCode(cel, xlEdgeRight) & _
                 EdgeCode(cel, xlEdgeBottom) & EdgeCode(cel, xlEdgeLeft)
End Function


Private Function EdgeCode(ByVal cel As Range, ByVal edge As Long) As String
    Dim ls As Variant, wt As Variant

    On Error Resume Next
    ls = cel.Borders(edge).LineStyle
    wt = cel.Borders(edge).Weight
    On Error GoTo 0

    If IsEmpty(ls) Then EdgeCode = "-": Exit Function

    Select Case ls
        Case xlLineStyleNone:  EdgeCode = "-"
        Case xlDouble:         EdgeCode = "="
        Case xlDot, xlDash, xlDashDot, xlDashDotDot: EdgeCode = "."
        Case Else
            Select Case wt
                Case xlHairline, xlThin: EdgeCode = "1"
                Case xlMedium:           EdgeCode = "2"
                Case xlThick:            EdgeCode = "3"
                Case Else:               EdgeCode = "1"
            End Select
    End Select
End Function


Private Function HAlignName(ByVal a As Variant) As String
    Select Case a
        Case xlLeft: HAlignName = "左"
        Case xlCenter: HAlignName = "中"
        Case xlRight: HAlignName = "右"
        Case xlCenterAcrossSelection: HAlignName = "選択範囲内で中央"
        Case xlDistributed: HAlignName = "均等"
        Case Else: HAlignName = ""
    End Select
End Function


Private Function VAlignName(ByVal a As Variant) As String
    Select Case a
        Case xlTop: VAlignName = "上"
        Case xlCenter: VAlignName = "中"
        Case xlBottom: VAlignName = "下"
        Case Else: VAlignName = ""
    End Select
End Function


Private Function FillName(ByVal cel As Range) As String
    If cel.Interior.ColorIndex = xlColorIndexNone Then Exit Function
    FillName = "RGB(" & (cel.Interior.Color Mod 256) & "," & _
               ((cel.Interior.Color \ 256) Mod 256) & "," & _
               ((cel.Interior.Color \ 65536) Mod 256) & ")"
End Function


Private Function PaperName(ByVal p As Variant) As String
    Select Case p
        Case 8: PaperName = "A3"
        Case 9: PaperName = "A4"
        Case 11: PaperName = "A5"
        Case 12: PaperName = "B4"
        Case 13: PaperName = "B5"
        Case Else: PaperName = "その他(" & p & ")"
    End Select
End Function


Private Function ColLetter(ByVal c As Long) As String
    Dim a As String
    a = Split(Cells(1, c).Address(True, False, xlA1), "$")(0)
    ColLetter = a
End Function


Private Function Fmt(ByVal v As Variant) As String
    On Error Resume Next
    Fmt = Format$(v, "0.##")
End Function
