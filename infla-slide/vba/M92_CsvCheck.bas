'==============================================================
' M92_CsvCheck  ―  CSVの構造を金額抜きで書き出す（診断用）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M92_CsvCheck" に。
'
' コード・名称・単位・規格と、マクロがどう判定したかだけを書き出します。
' 単価・数量・金額は一切出しません。
' 階層や諸経費の判定が合わないときは、この結果を送ってもらえれば直せます。
'==============================================================
Option Explicit


Public Sub CSV構造チェック()
    Dim v As Variant
    Dim rows As Collection
    Dim arr() As String
    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim cfg As Object
    Dim code As String, lvl As String, kubun As String
    Dim inZ As Boolean, inNaiyaku As Boolean
    Dim charSetName As String

    v = Application.GetOpenFilename("CSVファイル (*.csv),*.csv,すべてのファイル (*.*),*.*", , _
                                    "構造を確認するCSVを選択してください")
    If VarType(v) = vbBoolean Then Exit Sub

    charSetName = "Shift_JIS"
    If SheetExists(SH_CONFIG) Then
        Set cfg = GetConfig()
        charSetName = CStr(CfgVal(cfg, "文字コード", "Shift_JIS"))
    End If

    Application.ScreenUpdating = False
    Set rows = ParseCsvText(ReadTextFile(CStr(v), charSetName), ",")

    Set ws = FreshSheet("_CSV構造")
    ws.Range("A1").Value = "CSV構造チェック（金額・数量は出していません）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "ファイル：" & Mid$(CStr(v), InStrRev(CStr(v), Application.PathSeparator) + 1)

    ws.Range("A4").Value = "CSV行"
    ws.Range("B4").Value = "コード"
    ws.Range("C4").Value = "レベル判定"
    ws.Range("D4").Value = "行の区分"
    ws.Range("E4").Value = "施工単価名称"
    ws.Range("F4").Value = "単位"
    ws.Range("G4").Value = "規格1"
    ws.Range("H4").Value = "規格2"
    ws.Range("I4").Value = "摘要"
    With ws.Range("A4:I4")
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With

    r = 4
    inZ = False
    inNaiyaku = False

    For i = 1 To rows.Count
        arr = rows(i)
        code = Trim$(ColVal(arr, CSV_CODE))
        lvl = LevelOf(code)

        If lvl = "費目" Then inNaiyaku = True
        If lvl = "Z" Then inZ = True

        If Not inNaiyaku Then
            ' X1000より前はGコードの単価表。内訳には入れない
            kubun = "単価表部（内訳に入れません）"
            GoTo WriteLine
        End If

        Select Case lvl
            Case "G":    kubun = "内訳の明細（Gコードの代価）"
            Case "費目": kubun = "費目（本工事費）"
            Case "L1":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・工事区分"
            Case "L2":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・工種"
            Case "L3":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・種別"
            Case "L4":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・細別"
            Case "YZ":   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・YZ（工種あつかい）"
            Case "Z":    kubun = "★Zコード項目　区分＝" & _
                                 IIf(ZKubunPub(code, Trim$(ColVal(arr, CSV_NAME))) = "", _
                                     "（使わない）", ZKubunPub(code, Trim$(ColVal(arr, CSV_NAME))))
            Case Else:   kubun = IIf(inZ, "諸経費部", "直接工事費部") & "・明細"
        End Select

WriteLine:

        r = r + 1
        ws.Cells(r, 1).Value = i
        ws.Cells(r, 2).Value = "'" & code
        ws.Cells(r, 3).Value = IIf(lvl = "", "（明細）", lvl)
        ws.Cells(r, 4).Value = kubun
        ws.Cells(r, 5).Value = Trim$(ColVal(arr, CSV_NAME))
        ws.Cells(r, 6).Value = Trim$(ColVal(arr, CSV_TANI))
        ws.Cells(r, 7).Value = Trim$(ColVal(arr, CSV_KIKAKU1))
        ws.Cells(r, 8).Value = Trim$(ColVal(arr, CSV_KIKAKU2))
        ws.Cells(r, 9).Value = Trim$(ColVal(arr, CSV_TEKIYO))

        If lvl = "Z" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(255, 242, 204)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font.Bold = True
        ElseIf lvl = "費目" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = RGB(221, 235, 247)
        End If
    Next i

    ws.Range(ws.Cells(4, 1), ws.Cells(r, 9)).Borders.LineStyle = xlContinuous
    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 10
    ws.Columns(4).ColumnWidth = 30
    ws.Columns(5).ColumnWidth = 28
    ws.Columns(6).ColumnWidth = 7
    ws.Columns(7).ColumnWidth = 20
    ws.Columns(8).ColumnWidth = 20
    ws.Columns(9).ColumnWidth = 20

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("A4").Select

    MsgBox "「_CSV構造」シートに " & (r - 4) & " 行を書き出しました。" & vbCrLf & vbCrLf & _
           "A列からI列を選択してコピーすれば、そのまま送れます。" & vbCrLf & _
           "単価・数量・金額は含まれていません。", vbInformation
End Sub


' M20_Slide の ZKubun は Private なので、ここから使えるように同じ判定を置く
Private Function ZKubunPub(ByVal code As String, ByVal nm As String) As String
    Dim t As String, i As Long, ch As String, digits As String
    Dim n As Long

    If IsScrapName(nm) Then ZKubunPub = "ス（スクラップ）": Exit Function
    If InStr(1, NormText(nm), NormText("支給品費")) > 0 Then Exit Function

    t = NormText(code)
    If Len(t) = 0 Then Exit Function
    If UCase$(Left$(t, 1)) <> "Z" Then Exit Function

    For i = 2 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit For
    Next i
    If digits = "" Then Exit Function

    n = CLng(digits)
    If n < 40 Then
        ZKubunPub = "積（共通仮設費の積上分）"
    ElseIf n < 45 Then
        ZKubunPub = "原（工事原価に加算）"
    Else
        ZKubunPub = "価（工事価格に加算）"
    End If
End Function
