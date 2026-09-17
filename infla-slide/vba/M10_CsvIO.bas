'==============================================================
' M10_CsvIO  ―  CSV読み込み・文字列／数値ユーティリティ
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M10_CsvIO" にしてください。
' （プロパティウィンドウの (オブジェクト名) を書き換え）
' 参照設定は不要です（すべて CreateObject の遅延バインディング）。
'==============================================================
Option Explicit

' 明細1行の配列インデックス
Public Const F_KOSHU    As Long = 0   ' 工種
Public Const F_SHUBETSU As Long = 1   ' 種別
Public Const F_SAIBETSU As Long = 2   ' 細別
Public Const F_KIKAKU   As Long = 3   ' 規格
Public Const F_TANI     As Long = 4   ' 単位
Public Const F_SURYO    As Long = 5   ' 数量
Public Const F_TANKA    As Long = 6   ' 単価
Public Const F_KINGAKU  As Long = 7   ' 金額
Public Const F_TEKIYO   As Long = 8   ' 摘要
Public Const F_COUNT    As Long = 9


'--------------------------------------------------------------
' テキストファイルを指定文字コードで読み込む
'--------------------------------------------------------------
Public Function ReadTextFile(ByVal filePath As String, ByVal charSetName As String) As String
    Dim st As Object, s As String

    If charSetName = "" Then charSetName = "Shift_JIS"

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                      ' adTypeText
    st.Charset = charSetName
    st.Open
    st.LoadFromFile filePath
    s = st.ReadText(-1)              ' adReadAll
    st.Close

    ' UTF-8 BOM が残る場合があるので除去
    If Len(s) > 0 Then
        If Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    End If

    ReadTextFile = s
End Function


'--------------------------------------------------------------
' テキストファイルを指定文字コードで書き出す
'--------------------------------------------------------------
Public Sub WriteTextFile(ByVal filePath As String, ByVal s As String, ByVal charSetName As String)
    Dim st As Object

    If charSetName = "" Then charSetName = "Shift_JIS"

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                      ' adTypeText
    st.Charset = charSetName
    st.Open
    st.WriteText s
    st.SaveToFile filePath, 2        ' adSaveCreateOverWrite
    st.Close
End Sub

'--------------------------------------------------------------
' CSVテキストを解析して Collection(String配列) を返す
'   ・ダブルクォート囲み、"" によるエスケープ、囲み内の改行に対応
'--------------------------------------------------------------
Public Function ParseCsvText(ByVal s As String, ByVal delim As String) As Collection
    Dim rows As Collection
    Dim fields() As String
    Dim nField As Long
    Dim buf As String
    Dim inQuote As Boolean
    Dim i As Long, n As Long
    Dim ch As String

    If delim = "" Then delim = ","
    Set rows = New Collection

    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    n = Len(s)

    ReDim fields(0 To 63)
    nField = 0
    i = 1

    Do While i <= n
        ch = Mid$(s, i, 1)

        If inQuote Then
            If ch = """" Then
                If i < n And Mid$(s, i + 1, 1) = """" Then
                    buf = buf & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                buf = buf & ch
            End If

        ElseIf ch = """" Then
            inQuote = True

        ElseIf ch = delim Then
            If nField > UBound(fields) Then ReDim Preserve fields(0 To nField + 63)
            fields(nField) = buf
            buf = ""
            nField = nField + 1

        ElseIf ch = vbLf Then
            If nField > UBound(fields) Then ReDim Preserve fields(0 To nField + 63)
            fields(nField) = buf
            rows.Add TrimFieldArray(fields, nField)
            buf = ""
            nField = 0
            ReDim fields(0 To 63)

        Else
            buf = buf & ch
        End If

        i = i + 1
    Loop

    ' 最終行（末尾に改行が無い場合）
    If nField > 0 Or Len(buf) > 0 Then
        If nField > UBound(fields) Then ReDim Preserve fields(0 To nField + 63)
        fields(nField) = buf
        rows.Add TrimFieldArray(fields, nField)
    End If

    Set ParseCsvText = rows
End Function


Private Function TrimFieldArray(ByRef fields() As String, ByVal lastIdx As Long) As String()
    Dim r() As String, k As Long
    ReDim r(0 To lastIdx)
    For k = 0 To lastIdx
        r(k) = fields(k)
    Next k
    TrimFieldArray = r
End Function


'--------------------------------------------------------------
' 配列から列番号(1始まり)の値を安全に取り出す。範囲外は空文字。
'--------------------------------------------------------------
Public Function ColVal(ByRef arr() As String, ByVal colNo As Long) As String
    If colNo <= 0 Then Exit Function
    If colNo - 1 > UBound(arr) Then Exit Function
    ColVal = arr(colNo - 1)
End Function


'--------------------------------------------------------------
' 文字列を数値へ（カンマ・円記号・全角数字・空白を吸収）
'--------------------------------------------------------------
Public Function ToNum(ByVal v As String) As Double
    Dim t As String

    t = Trim$(v)
    If t = "" Then Exit Function

    On Error Resume Next
    t = StrConv(t, vbNarrow)          ' 全角 → 半角
    On Error GoTo 0

    t = Replace(t, ",", "")
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")
    t = Replace(t, "\", "")
    t = Replace(t, "¥", "")
    t = Replace(t, "円", "")

    If t = "" Or t = "-" Then Exit Function
    If Not IsNumeric(t) Then Exit Function

    ToNum = CDbl(t)
End Function


'--------------------------------------------------------------
' 突合キー用の正規化（前後空白・空白文字を除去し半角化）
'--------------------------------------------------------------
Public Function NormText(ByVal s As String) As String
    Dim t As String

    t = Trim$(s)
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")

    On Error Resume Next
    t = StrConv(t, vbNarrow)
    On Error GoTo 0

    NormText = t
End Function


'--------------------------------------------------------------
' 金額の端数処理
'   mode : "切り捨て" / "切り上げ" / "四捨五入"
'   unit : 丸め単位（1 / 10 / 100 / 1000 …）
'--------------------------------------------------------------
Public Function RoundAmt(ByVal v As Double, ByVal mode As String, ByVal unit As Double) As Double
    If unit <= 0 Then unit = 1

    Select Case mode
        Case "切り上げ"
            RoundAmt = -Int(-v / unit) * unit
        Case "四捨五入"
            RoundAmt = Int(v / unit + 0.5) * unit
        Case Else                       ' 切り捨て
            RoundAmt = Int(v / unit) * unit
    End Select
End Function
