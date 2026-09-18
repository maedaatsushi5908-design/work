'==============================================================
' M30_Shokeihi  ―  諸経費部（直接工事費計〜工事費）とスライド額の算出
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M30_Shokeihi" にしてください。
'
' 率式は神戸市・道路改良工事等を前提にしています（設定シートで変更可）。
' 工種区分が違う場合や運用が違う場合は、設定シートの係数を直してください。
'==============================================================
Option Explicit

Private mWs As Worksheet
Private mCfg As Object

' 明細行（＝黄色の手入力欄を置く範囲）の最終行。体裁処理が参照する
Public gLastDetailRow As Long


'==============================================================
' Zコード（諸経費）セクションがある場合
'   戻り値：最終行
'==============================================================
Public Function BuildShokeihi(ByVal ws As Worksheet, ByVal cfg As Object, _
                              ByVal recs As Collection, ByRef newTanka() As Double, _
                              ByVal firstZ As Long, ByVal startRow As Long, _
                              ByVal himokuRow As Long) As Long
    Dim i As Long, r As Long, cnt As Long
    Dim rec As Variant
    Dim lvl As String, lvlNum As Long
    Dim zRows() As Long, zLvls() As Long
    Dim rDirect As Long
    Dim tsumiStart As Long, tsumiEnd As Long
    Dim rGijutsu As Long
    Dim topRows As String

    Set mWs = ws
    Set mCfg = cfg

    r = startRow

    '--- 直接工事費計 ---
    rDirect = r
    CalcRow r, "直接工事費計"
    SetFour r, EqCell(himokuRow, SC_A_ALL), EqCell(himokuRow, SC_A_DONE), _
               EqCell(himokuRow, SC_B_ALL), EqCell(himokuRow, SC_B_REST)
    Emphasize r
    r = r + 1

    '--- 共通仮設費の積上げ分（Zセクションをそのまま出す）---
    tsumiStart = r
    ReDim zRows(1 To recs.Count)
    ReDim zLvls(1 To recs.Count)
    cnt = 0
    topRows = ""

    For i = firstZ To recs.Count
        rec = recs(i)
        lvl = CStr(rec(R_LEVEL))
        If lvl = "G" Then GoTo ContinueLoop

        lvlNum = LevelNum(lvl)
        WriteZRow ws, r, rec, newTanka(i), lvl, lvlNum

        cnt = cnt + 1
        zRows(cnt) = r
        zLvls(cnt) = lvlNum

        If lvl = "Z" Then
            If topRows <> "" Then topRows = topRows & ","
            topRows = topRows & "#" & r
            If InStr(1, NormText(CStr(rec(R_NAME))), "技術管理費") > 0 Then rGijutsu = r
        End If

        r = r + 1
ContinueLoop:
    Next i
    tsumiEnd = r - 1

    WriteSumFormulas ws, zRows, zLvls, cnt

    If tsumiEnd < tsumiStart Then
        tsumiStart = 0: tsumiEnd = 0
    End If

    BuildShokeihi = WriteCalcRows(r, rDirect, tsumiStart, tsumiEnd, rGijutsu, topRows)
End Function


'==============================================================
' Zコードが無いCSVの場合（諸経費行だけを作る）
'==============================================================
Public Function BuildShokeihiNoZ(ByVal ws As Worksheet, ByVal cfg As Object, _
                                 ByVal startRow As Long, ByVal himokuRow As Long) As Long
    Dim r As Long, rDirect As Long

    Set mWs = ws
    Set mCfg = cfg

    r = startRow
    rDirect = r
    CalcRow r, "直接工事費計"
    SetFour r, EqCell(himokuRow, SC_A_ALL), EqCell(himokuRow, SC_A_DONE), _
               EqCell(himokuRow, SC_B_ALL), EqCell(himokuRow, SC_B_REST)
    Emphasize r
    r = r + 1

    BuildShokeihiNoZ = WriteCalcRows(r, rDirect, 0, 0, 0, "")
End Function


'--------------------------------------------------------------
' Zセクションの1行
'--------------------------------------------------------------
Private Sub WriteZRow(ByVal ws As Worksheet, ByVal r As Long, ByVal rec As Variant, _
                      ByVal tankaNew As Double, ByVal lvl As String, ByVal lvlNum As Long)
    ws.Cells(r, SC_CODE).Value = rec(R_CODE)
    ws.Cells(r, SC_TEKIYO).Value = rec(R_TEKIYO)

    If lvl = "Z" Then
        ws.Cells(r, SC_KUBUN).Value = rec(R_NAME)          ' 名称を直接C列へ
    ElseIf lvl = "YZ" Then
        ws.Cells(r, SC_KOSHU).Value = "②"
        ws.Cells(r, SC_SHUBETSU).Value = rec(R_NAME)
    Else
        ws.Cells(r, SC_NAME).Value = rec(R_NAME)
        ws.Cells(r, SC_KIKAKU).Value = rec(R_KIKAKU)
    End If

    If lvlNum = 9 Then
        ws.Cells(r, SC_TANKA_O).Value = rec(R_TANKA)
        ws.Cells(r, SC_TANKA_N).Value = tankaNew
        ws.Cells(r, SC_Q_ALL).Value = rec(R_SURYO)
        ws.Cells(r, SC_TANI).Value = rec(R_TANI)
        WriteDetailFormulas ws, r
    Else
        ws.Cells(r, SC_Q_ALL).Value = 1
        ws.Cells(r, SC_Q_DONE).Value = 1
        ws.Cells(r, SC_Q_REST).Value = 1
        ws.Cells(r, SC_TANI).Value = "式"
    End If
End Sub


'==============================================================
' 処分費〜工事費、スライド額
'==============================================================
Private Function WriteCalcRows(ByVal startRow As Long, ByVal rDirect As Long, _
                               ByVal tsumiStart As Long, ByVal tsumiEnd As Long, _
                               ByVal rGijutsu As Long, ByVal topRows As String) As Long
    Dim r As Long
    Dim rShobun As Long, rTaisho As Long, rGengaku As Long, rTaishoAdj As Long
    Dim rRateK As Long, rRateKAdj As Long, rKasetsuRitsu As Long
    Dim rTsumiage As Long, rKasetsuKei As Long, rJun As Long
    Dim rGenbaTaisho As Long, rGenbaRate As Long, rGenbaRateAdj As Long, rGenba As Long
    Dim rGenka As Long, rIppanTaisho As Long, rIppanRate As Long, rHosho As Long, rIppan As Long
    Dim rScrap As Long, rKakaku As Long, rZei As Long, rHi As Long
    Dim kA As String, kB As String, kChiiki As String, kShukyu As String
    Dim gA As String, gB As String, gChiiki As String, gShukyu As String
    Dim iK As String, iTei As String, hosho As String, kojo As String, zei As String
    Dim lastData As Long

    kA = NumStr(CfgVal(mCfg, "共通仮設費率A", 1228.3))
    kB = NumStr(CfgVal(mCfg, "共通仮設費率B", -0.2614))
    kChiiki = NumStr(CfgVal(mCfg, "共通仮設費 地域補正", 1.2))
    kShukyu = NumStr(CfgVal(mCfg, "共通仮設費 週休補正", 1.04))
    gA = NumStr(CfgVal(mCfg, "現場管理費率A", 458.2))
    gB = NumStr(CfgVal(mCfg, "現場管理費率B", -0.1508))
    gChiiki = NumStr(CfgVal(mCfg, "現場管理費 地域補正", 1.1))
    gShukyu = NumStr(CfgVal(mCfg, "現場管理費 週休補正", 1.06))
    iK = NumStr(CfgVal(mCfg, "一般管理費率係数", -5.48972))
    iTei = NumStr(CfgVal(mCfg, "一般管理費率定数", 59.4977))
    hosho = NumStr(CfgVal(mCfg, "契約保障費率", 0.0004))
    kojo = NumStr(CfgVal(mCfg, "処分費控除率", 0.03))
    zei = NumStr(CfgVal(mCfg, "消費税率", 0.1))

    r = startRow
    lastData = r - 1
    gLastDetailRow = lastData

    '--- 処分費（明細行の処分費額を合計）---
    rShobun = r
    CalcRow r, "処分費"
    SetFour r, _
        "SUM(" & Rng(FIRST_ROW, lastData, SC_SA_ALL) & ")", _
        "SUM(" & Rng(FIRST_ROW, lastData, SC_SA_DONE) & ")", _
        "SUM(" & Rng(FIRST_ROW, lastData, SC_SB_ALL) & ")", _
        "SUM(" & Rng(FIRST_ROW, lastData, SC_SB_REST) & ")"
    Note r, "O列に〇を付けた明細の合計。処分費単価（U/V列）を入れた行はその単価で計算"
    r = r + 1

    '--- 共通仮設費率分対象額 ---
    rTaisho = r
    CalcRow r, "共通仮設費率分対象額"
    SetFour r, _
        Cel(rDirect, SC_A_ALL) & Tsumi(tsumiStart, tsumiEnd, SC_SA_ALL) & GijRef(rGijutsu, SC_A_ALL), _
        Cel(rDirect, SC_A_DONE) & Tsumi(tsumiStart, tsumiEnd, SC_SA_DONE) & GijRef(rGijutsu, SC_A_DONE), _
        Cel(rDirect, SC_B_ALL) & Tsumi(tsumiStart, tsumiEnd, SC_SB_ALL) & GijRef(rGijutsu, SC_B_ALL), _
        Cel(rDirect, SC_B_REST) & Tsumi(tsumiStart, tsumiEnd, SC_SB_REST) & GijRef(rGijutsu, SC_B_REST)
    If rGijutsu = 0 Then
        Warn r, "技術管理費の行が見つかりませんでした。控除が必要ならこの行の式を直してください"
    Else
        Note r, "直接工事費計 ＋ 積上げ分の処分費 － 技術管理費"
    End If
    r = r + 1

    '--- 減額分 ---
    rGengaku = r
    CalcRow r, "共通仮設費率分対象減額分"
    SetFour r, _
        Gengaku(rShobun, rTaisho, SC_A_ALL, kojo), _
        Gengaku(rShobun, rTaisho, SC_A_DONE, kojo), _
        Gengaku(rShobun, rTaisho, SC_B_ALL, kojo), _
        Cel(rGengaku, SC_B_ALL)
    Note r, "処分費が対象額の" & CStr(CDbl(CfgVal(mCfg, "処分費控除率", 0.03)) * 100) & "%を超えた分"
    r = r + 1

    '--- 対象額（補正後）---
    rTaishoAdj = r
    CalcRow r, "共通仮設費率分対象額（補正後）"
    SetFourCols r, rTaisho, rGengaku, "-"
    r = r + 1

    '--- 共通仮設費率 ---
    rRateK = r
    CalcRow r, "共通仮設費率"
    SetFour r, _
        "ROUND(" & kA & "*(" & Cel(rTaisho, SC_A_ALL) & "-" & Cel(rGengaku, SC_A_ALL) & ")^(" & kB & "),2)", _
        Cel(rRateK, SC_A_ALL), _
        "ROUND(" & kA & "*(" & Cel(rTaisho, SC_B_ALL) & "-" & Cel(rGengaku, SC_B_ALL) & ")^(" & kB & "),2)", _
        Cel(rRateK, SC_B_ALL)
    RateFormat r
    Note r, "率 ＝ " & kA & " × 対象額 ^ (" & kB & ")　※工種区分で係数が変わります"
    r = r + 1

    rRateKAdj = r
    CalcRow r, "共通仮設費率（補正後）"
    SetFour r, _
        "ROUND(ROUND(" & Cel(rRateK, SC_A_ALL) & "*" & kChiiki & ",2)*" & kShukyu & ",2)", _
        Cel(rRateKAdj, SC_A_ALL), _
        "ROUND(ROUND(" & Cel(rRateK, SC_B_ALL) & "*" & kChiiki & ",2)*" & kShukyu & ",2)", _
        Cel(rRateKAdj, SC_B_ALL)
    RateFormat r
    Note r, "地域補正 ×" & kChiiki & "　週休補正 ×" & kShukyu
    r = r + 1

    rKasetsuRitsu = r
    CalcRow r, "共通仮設費率分"
    SetFourEach r, "ROUNDDOWN(#" & rTaishoAdj & "*(#" & rRateKAdj & "/100),-3)"
    r = r + 1

    rTsumiage = r
    CalcRow r, "共通仮設費（積上げ分）"
    If topRows <> "" Then
        SetFourEach r, "SUM(" & topRows & ")"
    Else
        SetFour r, "0", "0", "0", "0"
        Warn r, "積上げ分の行が見つかりませんでした"
    End If
    r = r + 1

    rKasetsuKei = r
    CalcRow r, "共通仮設費計"
    SetFourEach r, "#" & rKasetsuRitsu & "+#" & rTsumiage
    r = r + 1

    rJun = r
    CalcRow r, "純工事費計"
    SetFourEach r, "#" & rDirect & "+#" & rKasetsuKei
    Emphasize r
    r = r + 1

    '--- 現場管理費 ---
    rGenbaTaisho = r
    CalcRow r, "現場管理費率対象額"
    SetFourEach r, "ROUNDDOWN(#" & rJun & "-#" & rGengaku & GijPlus(rGijutsu) & ",0)"
    r = r + 1

    rGenbaRate = r
    CalcRow r, "現場管理費率"
    SetFour r, _
        "ROUND(" & gA & "*(" & Cel(rJun, SC_A_ALL) & "-" & Cel(rGengaku, SC_A_ALL) & GijRef(rGijutsu, SC_A_ALL) & ")^(" & gB & "),2)", _
        Cel(rGenbaRate, SC_A_ALL), _
        "ROUND(" & gA & "*(" & Cel(rJun, SC_B_ALL) & "-" & Cel(rGengaku, SC_B_ALL) & GijRef(rGijutsu, SC_B_ALL) & ")^(" & gB & "),2)", _
        Cel(rGenbaRate, SC_B_ALL)
    RateFormat r
    r = r + 1

    rGenbaRateAdj = r
    CalcRow r, "現場管理費率（補正後）"
    SetFour r, _
        "ROUND(ROUND(" & Cel(rGenbaRate, SC_A_ALL) & "*" & gChiiki & ",2)*" & gShukyu & ",2)", _
        Cel(rGenbaRateAdj, SC_A_ALL), _
        "ROUND(ROUND(" & Cel(rGenbaRate, SC_B_ALL) & "*" & gChiiki & ",2)*" & gShukyu & ",2)", _
        Cel(rGenbaRateAdj, SC_B_ALL)
    RateFormat r
    r = r + 1

    rGenba = r
    CalcRow r, "現場管理費"
    SetFourEach r, "ROUNDDOWN(#" & rGenbaTaisho & "*(#" & rGenbaRateAdj & "/100),-3)"
    r = r + 1

    rGenka = r
    CalcRow r, "工事原価計"
    SetFourEach r, "#" & rJun & "+#" & rGenba
    Emphasize r
    r = r + 1

    '--- 一般管理費 ---
    rIppanTaisho = r
    CalcRow r, "一般管理費率対象額"
    SetFourEach r, "ROUNDDOWN(#" & rGenka & "-#" & rGengaku & GijPlus(rGijutsu) & ",0)"
    r = r + 1

    rIppanRate = r
    CalcRow r, "一般管理費率"
    SetFour r, _
        "ROUND(" & iK & "*LOG10(" & Cel(rGenka, SC_A_ALL) & "-" & Cel(rGengaku, SC_A_ALL) & GijRef(rGijutsu, SC_A_ALL) & ")+" & iTei & ",2)", _
        Cel(rIppanRate, SC_A_ALL), _
        "ROUND(" & iK & "*LOG10(" & Cel(rGenka, SC_B_ALL) & "-" & Cel(rGengaku, SC_B_ALL) & GijRef(rGijutsu, SC_B_ALL) & ")+" & iTei & ",2)", _
        Cel(rIppanRate, SC_B_ALL)
    RateFormat r
    r = r + 1

    rHosho = r
    CalcRow r, "契約保障費"
    SetFour r, _
        "ROUNDDOWN(" & Cel(rIppanTaisho, SC_A_ALL) & "*" & hosho & ",0)", _
        Cel(rHosho, SC_A_ALL), _
        "ROUNDDOWN(" & Cel(rIppanTaisho, SC_B_ALL) & "*" & hosho & ",0)", _
        Cel(rHosho, SC_B_ALL)
    r = r + 1

    rIppan = r
    CalcRow r, "一般管理費"
    SetFourEach r, "#" & rIppanTaisho & "*(#" & rIppanRate & "/100)+#" & rHosho
    r = r + 1

    '--- ｽｸﾗｯﾌﾟ・工事価格・工事費 ---
    rScrap = r
    CalcRow r, "ｽｸﾗｯﾌﾟ"
    mWs.Range(mWs.Cells(r, SC_A_ALL), mWs.Cells(r, SC_B_REST)).Value = 0
    mWs.Range(mWs.Cells(r, SC_A_ALL), mWs.Cells(r, SC_B_REST)).Interior.Color = CLR_INPUT
    Note r, "該当があれば手入力（マイナス計上）"
    r = r + 1

    rKakaku = r
    CalcRow r, "工事価格"
    SetFourEach r, "ROUNDDOWN(#" & rGenka & "+#" & rIppan & "+#" & rScrap & ",-3)"
    Emphasize r
    r = r + 1

    rZei = r
    CalcRow r, "消費税相当額"
    SetFourEach r, "ROUNDDOWN(#" & rKakaku & "*" & zei & ",-1)"
    r = r + 1

    rHi = r
    CalcRow r, "工事費"
    SetFourEach r, "ROUNDDOWN(#" & rKakaku & "+#" & rZei & ",-1)"
    Emphasize r
    r = r + 1

    '--- スライド額の算出 ---
    r = r + 1
    WriteCalcResult r, rHi

    WriteCalcRows = r + 7
End Function


'--------------------------------------------------------------
' スライド額の算出ブロック
'--------------------------------------------------------------
Private Sub WriteCalcResult(ByVal r As Long, ByVal rHi As Long)
    Dim contract As Double, rate As Double
    Dim c As Long

    contract = CDbl(CfgVal(mCfg, "請負代金額", 0))
    rate = 0.01

    c = SC_A_ALL

    mWs.Cells(r, SC_KUBUN).Value = "【スライド額の算出】　※運用の確認が必要です"
    mWs.Cells(r, SC_KUBUN).Font.Bold = True

    Lab r + 1, "スライド前 工事費（全体）", "A", "=" & Cel(rHi, SC_A_ALL)
    Lab r + 2, "スライド後 工事費（全体）", "B", "=" & Cel(rHi, SC_B_ALL)
    Lab r + 3, "スライド差額", "C = B - A", "=" & Cel(r + 2, c) & "-" & Cel(r + 1, c)
    Lab r + 4, "現請負代金額", "D", CStr(contract)
    Lab r + 5, "受注者負担額（1%）", "E = D × 1%", "=ROUNDDOWN(" & Cel(r + 4, c) & "*" & CStr(rate) & ",0)"
    Lab r + 6, "変更額", "F = C - E", "=IF(" & Cel(r + 3, c) & "-" & Cel(r + 5, c) & ">0," & _
                                      Cel(r + 3, c) & "-" & Cel(r + 5, c) & ",0)"
    Lab r + 7, "変更後請負代金額", "D + F", "=" & Cel(r + 4, c) & "+" & Cel(r + 6, c)

    mWs.Cells(r + 4, c).Interior.Color = CLR_INPUT
    With mWs.Range(mWs.Cells(r + 6, SC_KUBUN), mWs.Cells(r + 7, SC_A_ALL))
        .Font.Bold = True
        .Interior.Color = CLR_TOTAL
    End With
End Sub


Private Sub Lab(ByVal r As Long, ByVal label As String, ByVal expr As String, ByVal f As String)
    mWs.Cells(r, SC_KUBUN).Value = label
    mWs.Cells(r, SC_NAME).Value = expr
    mWs.Cells(r, SC_NAME).Font.Color = RGB(120, 120, 120)
    If Left$(f, 1) = "=" Then
        mWs.Cells(r, SC_A_ALL).Formula = f
    Else
        mWs.Cells(r, SC_A_ALL).Value = CDbl(f)
    End If
    mWs.Cells(r, SC_A_ALL).NumberFormatLocal = "#,##0"
    mWs.Range(mWs.Cells(r, SC_KUBUN), mWs.Cells(r, SC_A_ALL)).Borders.LineStyle = xlContinuous
End Sub


'==============================================================
' 補助
'==============================================================
Private Sub CalcRow(ByVal r As Long, ByVal label As String)
    mWs.Cells(r, SC_KUBUN).Value = label
    mWs.Cells(r, SC_Q_ALL).Value = 1
    mWs.Cells(r, SC_Q_DONE).Value = 1
    mWs.Cells(r, SC_Q_REST).Value = 1
    mWs.Cells(r, SC_TANI).Value = "式"
End Sub


Private Sub SetFour(ByVal r As Long, ByVal fP As String, ByVal fQ As String, _
                    ByVal fR As String, ByVal fS As String)
    mWs.Cells(r, SC_A_ALL).Formula = "=" & fP
    mWs.Cells(r, SC_A_DONE).Formula = "=" & fQ
    mWs.Cells(r, SC_B_ALL).Formula = "=" & fR
    mWs.Cells(r, SC_B_REST).Formula = "=" & fS
End Sub


' 4列とも同じ形。テンプレ内の "#123" は「123行目の同じ列」に置き換える
Private Sub SetFourEach(ByVal r As Long, ByVal tpl As String)
    Dim cols As Variant, c As Variant
    cols = Array(SC_A_ALL, SC_A_DONE, SC_B_ALL, SC_B_REST)
    For Each c In cols
        mWs.Cells(r, c).Formula = "=" & ExpandTpl(tpl, CLng(c))
    Next c
End Sub


Private Sub SetFourCols(ByVal r As Long, ByVal r1 As Long, _
                        ByVal r2 As Long, ByVal op As String)
    Dim cols As Variant, c As Variant
    cols = Array(SC_A_ALL, SC_A_DONE, SC_B_ALL, SC_B_REST)
    For Each c In cols
        mWs.Cells(r, c).Formula = "=" & Cel(r1, CLng(c)) & op & Cel(r2, CLng(c))
    Next c
End Sub


Private Function ExpandTpl(ByVal tpl As String, ByVal col As Long) As String
    Dim s As String, i As Long, j As Long, num As String
    s = tpl
    Do
        i = InStr(s, "#")
        If i = 0 Then Exit Do
        j = i + 1
        num = ""
        Do While j <= Len(s)
            If Mid$(s, j, 1) Like "#" Then
                num = num & Mid$(s, j, 1)
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        s = Left$(s, i - 1) & Cel(CLng(num), col) & Mid$(s, j)
    Loop
    ExpandTpl = s
End Function


Private Function Cel(ByVal r As Long, ByVal c As Long) As String
    Cel = mWs.Cells(r, c).Address(False, False)
End Function


Private Function EqCell(ByVal r As Long, ByVal c As Long) As String
    EqCell = mWs.Cells(r, c).Address(False, False)
End Function


Private Function Rng(ByVal r1 As Long, ByVal r2 As Long, ByVal c As Long) As String
    If r2 < r1 Then r2 = r1
    Rng = mWs.Cells(r1, c).Address(False, False) & ":" & mWs.Cells(r2, c).Address(False, False)
End Function


Private Function Tsumi(ByVal r1 As Long, ByVal r2 As Long, ByVal c As Long) As String
    If r1 = 0 Then Exit Function
    Tsumi = "+SUM(" & Rng(r1, r2, c) & ")"
End Function


Private Function GijRef(ByVal rGijutsu As Long, ByVal c As Long) As String
    If rGijutsu = 0 Then Exit Function
    GijRef = "-" & Cel(rGijutsu, c)
End Function


Private Function GijPlus(ByVal rGijutsu As Long) As String
    If rGijutsu = 0 Then Exit Function
    GijPlus = "-#" & rGijutsu
End Function


Private Function Gengaku(ByVal rShobun As Long, ByVal rTaisho As Long, _
                         ByVal c As Long, ByVal kojo As String) As String
    Dim a As String
    a = Cel(rShobun, c) & "-" & Cel(rTaisho, c) & "*" & kojo
    Gengaku = "IF(" & a & ">0," & a & ",0)"
End Function


Private Function NumStr(ByVal v As Variant) As String
    NumStr = Format$(CDbl(v), "0.##########")
End Function


Private Sub RateFormat(ByVal r As Long)
    mWs.Range(mWs.Cells(r, SC_A_ALL), mWs.Cells(r, SC_B_REST)).NumberFormatLocal = "0.00"
End Sub


Private Sub Emphasize(ByVal r As Long)
    With mWs.Range(mWs.Cells(r, SC_HIMOKU), mWs.Cells(r, SC_TEKIYO))
        .Font.Bold = True
        .Interior.Color = CLR_TOTAL
    End With
End Sub


Private Sub Note(ByVal r As Long, ByVal s As String)
    mWs.Cells(r, SC_TEKIYO).Value = s
    mWs.Cells(r, SC_TEKIYO).Font.Color = RGB(120, 120, 120)
End Sub


Private Sub Warn(ByVal r As Long, ByVal s As String)
    mWs.Cells(r, SC_TEKIYO).Value = "★ " & s
    mWs.Cells(r, SC_TEKIYO).Font.Color = RGB(192, 0, 0)
    mWs.Cells(r, SC_TEKIYO).Interior.Color = CLR_INPUT
End Sub
