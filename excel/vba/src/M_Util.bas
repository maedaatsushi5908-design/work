Attribute VB_Name = "M_Util"
'==================================================================
' M_Util - 文字列正規化・セル操作の共通ヘルパー
'
' 数量計算書は作成者や年度によって全角/半角、スペースの入り方が
' まちまちなので、比較する前に必ず Norm を通して表記ゆれを吸収する。
'==================================================================
Option Explicit

' 照合方式
Public Enum MatchMode
    mmPartial = 0   ' 部分一致（既定）
    mmExact = 1     ' 完全一致
    mmDia = 2       ' 口径一致（数字だけ取り出して数値比較）
End Enum

'------------------------------------------------------------------
' 表記ゆれを吸収した比較用文字列を返す
'   ・全角英数記号 → 半角
'   ・空白（半角/全角/タブ/改行）を除去
'   ・φ Φ ｆ などの径記号を除去
'   ・英字は大文字に統一
'------------------------------------------------------------------
Public Function Norm(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Or IsEmpty(s) Then Norm = "": Exit Function
    t = CStr(s)
    If Len(t) = 0 Then Norm = "": Exit Function

    t = StrConv(t, vbNarrow)                 ' 全角→半角
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, vbTab, "")
    t = Replace(t, " ", "")                  ' 半角スペース
    t = Replace(t, ChrW(&H3000), "")         ' 全角スペース
    t = Replace(t, ChrW(&H3C6), "")          ' φ
    t = Replace(t, ChrW(&H3A6), "")          ' Φ
    t = Replace(t, "f", "")
    t = Replace(t, "F", "")
    Norm = UCase$(t)
End Function

'------------------------------------------------------------------
' 文字列から最初の数字の並びを取り出して数値で返す（見つからなければ -1）
'   "SP400A" → 400   "φ400 1種管" → 400   "300A" → 300
'------------------------------------------------------------------
Public Function ExtractNum(ByVal s As Variant) As Double
    Dim t As String, i As Long, ch As String, buf As String
    Dim started As Boolean

    ExtractNum = -1
    If IsError(s) Or IsEmpty(s) Then Exit Function
    t = StrConv(CStr(s), vbNarrow)

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then
            buf = buf & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i

    If Len(buf) > 0 Then ExtractNum = CDbl(buf)
End Function

'------------------------------------------------------------------
' 指定の照合方式でキーと対象を比較する
'------------------------------------------------------------------
Public Function KeyMatches(ByVal target As Variant, ByVal key As String, _
                           ByVal mode As MatchMode) As Boolean
    Dim nt As String, nk As String

    If Len(key) = 0 Then KeyMatches = False: Exit Function

    Select Case mode
        Case mmDia
            Dim a As Double, b As Double
            a = ExtractNum(target)
            b = ExtractNum(key)
            KeyMatches = (a >= 0 And b >= 0 And a = b)
        Case mmExact
            KeyMatches = (Norm(target) = Norm(key))
        Case Else
            nt = Norm(target)
            nk = Norm(key)
            KeyMatches = (Len(nt) > 0 And InStr(1, nt, nk, vbTextCompare) > 0)
    End Select
End Function

'------------------------------------------------------------------
' 列文字("A" / "AO") → 列番号。数字が渡された場合はそのまま数値化。
' 不正な値は 0 を返す。
'------------------------------------------------------------------
Public Function ColToNum(ByVal col As Variant) As Long
    Dim s As String
    s = Trim$(StrConv(CStr(col), vbNarrow))
    If Len(s) = 0 Then ColToNum = 0: Exit Function

    If IsNumeric(s) Then
        ColToNum = CLng(s)
    Else
        On Error Resume Next
        ColToNum = Range(s & "1").Column
        If Err.Number <> 0 Then Err.Clear: ColToNum = 0
        On Error GoTo 0
    End If
End Function

'------------------------------------------------------------------
' シート名をゆるく探す（全角/半角スペースの違いを無視）
' 見つからなければ Nothing
'------------------------------------------------------------------
Public Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet, target As String
    target = Norm(sheetName)

    For Each ws In wb.Worksheets
        If Norm(ws.Name) = target Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Set FindSheet = Nothing
End Function

'------------------------------------------------------------------
' セルが数値ならその値、それ以外は Empty を返す
'------------------------------------------------------------------
Public Function NumOrEmpty(ByVal c As Range) As Variant
    If IsError(c.Value) Then
        NumOrEmpty = Empty
    ElseIf IsNumeric(c.Value) And Not IsEmpty(c.Value) And CStr(c.Value) <> "" Then
        NumOrEmpty = CDbl(c.Value)
    Else
        NumOrEmpty = Empty
    End If
End Function
