Attribute VB_Name = "M_Engine"
'==================================================================
' M_Engine - 転記元セルの解決エンジン
'
' 定義シートの「キー」からセル番地を毎回探し直す。番地を直接
' 持たないので、行が増減しても、ファイル名が元の日本語名でも動く。
'==================================================================
Option Explicit

Private mOpened As Collection      ' このマクロが開いたブック（最後に閉じる）
Private mFolder As String          ' 転記元フォルダ（一度選んだら記憶）

'==================================================================
' ブックの管理
'==================================================================
Public Sub InitSources()
    Set mOpened = New Collection
    mFolder = ""
End Sub

Public Sub CloseSources()
    Dim i As Long, wb As Workbook
    If mOpened Is Nothing Then Exit Sub
    For i = mOpened.Count To 1 Step -1
        On Error Resume Next
        Set wb = mOpened(i)
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        On Error GoTo 0
    Next i
    Set mOpened = Nothing
End Sub

'------------------------------------------------------------------
' 転記元ブックを取得する
'   1. 既に開いていればそれを使う
'   2. ThisWorkbook と同じフォルダを探す
'   3. 見つからなければ、先頭の番号（02 / 05-1 など）が一致する
'      ファイルを探す ＝ 元の日本語ファイル名のままでも拾える
'   4. それでも駄目ならフォルダを尋ねる（一度だけ）
'------------------------------------------------------------------
Public Function GetSourceBook(ByVal fileName As String, ByRef errMsg As String) As Workbook
    Dim wb As Workbook, path_ As String, actual As String

    errMsg = ""
    If Len(fileName) = 0 Then errMsg = "元ファイルが指定されていません": Exit Function

    ' 1. 開いているブックから
    For Each wb In Application.Workbooks
        If Norm(wb.Name) = Norm(fileName) Then
            Set GetSourceBook = wb
            Exit Function
        End If
    Next wb

    ' 2～3. フォルダから探す
    actual = LocateFile(ThisWorkbook.path, fileName)
    If Len(actual) = 0 And Len(mFolder) > 0 Then
        actual = LocateFile(mFolder, fileName)
    End If

    ' 4. フォルダを尋ねる
    If Len(actual) = 0 Then
        If AskFolder(fileName) Then
            actual = LocateFile(mFolder, fileName)
        End If
    End If

    If Len(actual) = 0 Then
        errMsg = "ファイルが見つかりません: " & fileName
        Exit Function
    End If

    ' 既に開いていないか、実ファイル名で再確認
    For Each wb In Application.Workbooks
        If Norm(wb.Name) = Norm(Dir$(actual)) Then
            Set GetSourceBook = wb
            Exit Function
        End If
    Next wb

    On Error GoTo OpenFailed
    Set wb = Application.Workbooks.Open(fileName:=actual, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0

    mOpened.Add wb
    Set GetSourceBook = wb
    Exit Function

OpenFailed:
    errMsg = "ファイルを開けません: " & actual & " (" & Err.Description & ")"
End Function

'------------------------------------------------------------------
' フォルダ内から目的のファイルを探して、フルパスを返す
'------------------------------------------------------------------
Private Function LocateFile(ByVal folder_ As String, ByVal wantName As String) As String
    Dim p As String, token As String, f As String, ext As String

    LocateFile = ""
    If Len(folder_) = 0 Then Exit Function
    If Right$(folder_, 1) <> "\" Then folder_ = folder_ & "\"

    ' 完全一致
    p = folder_ & wantName
    If Len(Dir$(p)) > 0 Then LocateFile = p: Exit Function

    ' 先頭の番号が一致するファイル（例 "02_..." → "02　東白川　鋳鉄管製造..."）
    token = LeadingToken(wantName)
    If Len(token) = 0 Then Exit Function

    f = Dir$(folder_ & "*.xls*")
    Do While Len(f) > 0
        If LeadingToken(f) = token Then
            LocateFile = folder_ & f
            Exit Function
        End If
        f = Dir$
    Loop
End Function

'------------------------------------------------------------------
' ファイル名の先頭にある番号を取り出す（"05-1_x.xlsx" → "05-1"）
'------------------------------------------------------------------
Private Function LeadingToken(ByVal s As String) As String
    Dim i As Long, ch As String, buf As String
    s = StrConv(s, vbNarrow)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Then
            buf = buf & ch
        Else
            Exit For
        End If
    Next i
    Do While Right$(buf, 1) = "-"
        buf = Left$(buf, Len(buf) - 1)
    Loop
    LeadingToken = buf
End Function

Private Function AskFolder(ByVal fileName As String) As Boolean
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "「" & fileName & "」がある フォルダを選んでください"
    If fd.Show = -1 Then
        mFolder = fd.SelectedItems(1)
        AskFolder = True
    End If
End Function

'==================================================================
' 値の解決
'==================================================================
Public Function ResolveValue(ByRef d As TDef, ByRef outVal As Variant, _
                             ByRef errMsg As String) As Boolean
    Dim wb As Workbook, ws As Worksheet
    Dim r As Long, c As Long, v As Variant

    outVal = Empty
    errMsg = ""

    Set wb = GetSourceBook(d.SrcFile, errMsg)
    If wb Is Nothing Then Exit Function

    Set ws = FindSheet(wb, d.SrcSheet)
    If ws Is Nothing Then
        errMsg = "シートがありません: [" & wb.Name & "] " & d.SrcSheet
        Exit Function
    End If

    r = FindTargetRow(ws, d)
    If r <= 0 Then
        errMsg = "行が見つかりません: キー1=" & d.Key1 & _
                 IIf(Len(d.Key2) > 0, " / キー2=" & d.Key2, "")
        Exit Function
    End If

    c = FindValueCol(ws, d)
    If c <= 0 Then
        errMsg = "列が特定できません: 値列=" & d.ValueCol & " 見出しキー=" & d.HeaderKey
        Exit Function
    End If

    v = NumOrEmpty(ws.Cells(r, c))
    If IsEmpty(v) Then
        errMsg = "数値ではありません: " & ws.Name & "!" & ws.Cells(r, c).Address(False, False)
        Exit Function
    End If

    If d.Decimals >= 0 Then v = Application.WorksheetFunction.Round(v, d.Decimals)

    outVal = v
    ResolveValue = True
End Function

'------------------------------------------------------------------
' キーに合う行を探す
'   キー列が結合や省略で空になっている場合に備え、上方向に値を補完する
'------------------------------------------------------------------
Private Function FindTargetRow(ByVal ws As Worksheet, ByRef d As TDef) As Long
    Dim k1c As Long, k2c As Long, lastRow As Long, r As Long
    Dim cur1 As String, v As Variant
    Dim ok1 As Boolean, ok2 As Boolean

    k1c = ColToNum(d.Key1Col)
    k2c = ColToNum(d.Key2Col)
    lastRow = LastUsedRow(ws)

    If k1c = 0 And k2c = 0 Then FindTargetRow = 0: Exit Function

    For r = 1 To lastRow
        ' キー1（上方向に補完）
        If k1c > 0 Then
            v = ws.Cells(r, k1c).Value
            If Len(Norm(v)) > 0 Then cur1 = CStr(v)
        End If

        If k1c = 0 Or Len(d.Key1) = 0 Then
            ok1 = True
        Else
            ok1 = KeyMatches(cur1, d.Key1, d.Key1Mode)
        End If

        If ok1 Then
            If k2c = 0 Or Len(d.Key2) = 0 Then
                ok2 = True
            Else
                ok2 = KeyMatches(ws.Cells(r, k2c).Value, d.Key2, d.Key2Mode)
            End If

            If ok2 Then
                FindTargetRow = r + d.RowOffset
                Exit Function
            End If
        End If
    Next r

    FindTargetRow = 0
End Function

'------------------------------------------------------------------
' 値のある列を決める
'   値列が指定されていればそれ。無ければ見出しから探す。
'   指定の見出し行に無い場合は、上から40行を走査して探し直す。
'------------------------------------------------------------------
Private Function FindValueCol(ByVal ws As Worksheet, ByRef d As TDef) As Long
    Dim c As Long, r As Long, lastCol As Long, mode As MatchMode

    If Len(d.ValueCol) > 0 Then
        FindValueCol = ColToNum(d.ValueCol)
        Exit Function
    End If

    If Len(d.HeaderKey) = 0 Then FindValueCol = 0: Exit Function

    ' 見出しが数字（口径）なら数値比較、そうでなければ部分一致
    If IsNumeric(StrConv(d.HeaderKey, vbNarrow)) Then
        mode = mmDia
    Else
        mode = mmPartial
    End If

    lastCol = LastUsedCol(ws)

    If d.HeaderRow > 0 Then
        For c = 1 To lastCol
            If KeyMatches(ws.Cells(d.HeaderRow, c).Value, d.HeaderKey, mode) Then
                FindValueCol = c
                Exit Function
            End If
        Next c
    End If

    ' 見つからないので行がずれたとみなして探し直す
    For r = 1 To Application.Min(40, LastUsedRow(ws))
        For c = 1 To lastCol
            If KeyMatches(ws.Cells(r, c).Value, d.HeaderKey, mode) Then
                FindValueCol = c
                Exit Function
            End If
        Next c
    Next r

    FindValueCol = 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If Err.Number <> 0 Or LastUsedRow < 1 Then LastUsedRow = 1
    On Error GoTo 0
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If Err.Number <> 0 Or LastUsedCol < 1 Then LastUsedCol = 1
    On Error GoTo 0
End Function
