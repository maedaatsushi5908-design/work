'==============================================================
' M40_Keihi  ―  経費計算シート と スライド調書（様式4-2号）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M40_Keihi" にしてください。
'
' 「203_経費計算シート」と同じ並びで、スライド計算表の明細から
' 直接工事費〜工事費を9系列で計算します。
'==============================================================
Option Explicit

Public Const SH_KEIHI As String = "経費計算シート"
Public Const SH_CHOSHO As String = "スライド調書（様式4-2号）"

' 経費計算シートの行
Private Const K_DCHOKU  As Long = 5     ' ①直接工事費
Private Const K_KANZAI  As Long = 6     ' ①'水道工事における管材費
Private Const K_SHOBUN  As Long = 7     ' ②直接工事費内の処分費等
Private Const K_TAIGAI  As Long = 8     ' ③率計算の対象外費
Private Const K_UNPAN   As Long = 10    ' ④共通仮設費積上分-運搬費
Private Const K_TSUMI2  As Long = 11    ' ④'その他の積上分
Private Const K_GIJUTSU As Long = 12    ' ⑤共通仮設費積上分-技術管理費
Private Const K_KTAISHO As Long = 13    ' ⑥共通仮設費対象額
Private Const K_KRITSU  As Long = 14    ' ⑦共通仮設費率
Private Const K_KBUN    As Long = 15    ' ⑦'共通仮設費率分
Private Const K_KKEI    As Long = 16    ' ⑧共通仮設費【合計】
Private Const K_JUN     As Long = 17    ' ⑨純工事費
Private Const K_GTAISHO As Long = 19    ' ⑩現場管理費対象額
Private Const K_GRITSU  As Long = 20    ' ⑪現場管理費率
Private Const K_GKEI    As Long = 21    ' ⑫現場管理費【合計】
Private Const K_GENKA   As Long = 22    ' ⑬工事原価
Private Const K_ITAISHO As Long = 24    ' ⑭一般管理費対象額
Private Const K_IRITSU  As Long = 25    ' ⑮一般管理費率
Private Const K_IBUN    As Long = 26    ' ⑮'一般管理費率分
Private Const K_HOSHO   As Long = 27    ' ⑯契約保証費
Private Const K_IKEI0   As Long = 28    ' ⑰一般管理費【合計】（端数整理前）
Private Const K_KAKAKU0 As Long = 29    ' ⑱工事価格（端数整理前）
Private Const K_HASU    As Long = 30    ' ⑲端数整理
Private Const K_IKEI    As Long = 31    ' ⑳一般管理費【合計】（端数整理後）
Private Const K_KAKAKU  As Long = 32    ' ㉑工事価格
Private Const K_SCRAP   As Long = 34    ' ㉒スクラップ
Private Const K_KAKAKU2 As Long = 35    ' ㉑'工事価格（スクラップ込み）
Private Const K_ZEI     As Long = 36    ' ㉓消費税相当額
Private Const K_KOUJIHI As Long = 37    ' ㉔工事費

Private Const K_FIRSTCOL As Long = 3    ' C列＝系列A
Private Const K_NSERIES  As Long = 9

Private mWs As Worksheet
Private mSlide As Worksheet
Private mCfg As Object


'==============================================================
' 経費計算シートを作る
'==============================================================
Public Sub BuildKeihiSheet(ByVal cfg As Object)
    Dim i As Long, c As Long
    Dim srcAmt As Variant, srcSho As Variant, exShin As Variant
    Dim rateSrc As Variant, sym As Variant, nm As Variant
    Dim unpan As String, gijutsu As String
    Dim kojo As String, zei As String

    Set mCfg = cfg
    Set mSlide = ThisWorkbook.Worksheets(SH_SLIDE)
    Set mWs = FreshSheet(SH_KEIHI)

    kojo = NumStr(CfgVal(cfg, "処分費控除率", 0.03))
    zei = NumStr(CfgVal(cfg, "消費税率", 0.1))

    '            A            B            C            D            E            F            G            H            I
    srcAmt = Array(SC_A_ALL, SC_A_REST, SC_B_REST, SC_A_DONE, SC_B_ALL, SC_B_REST, SC_A_DONE, SC_A_ALL, SC_A_DONE)
    srcSho = Array(SC_SA_ALL, SC_SA_REST, SC_SB_REST, SC_SA_DONE, SC_SB_ALL, SC_SB_REST, SC_SA_DONE, SC_SA_ALL, SC_SA_DONE)
    exShin = Array(False, False, False, False, False, False, False, True, True)
    ' 0＝自前の率（手入力可）／-1＝率なし（素材のみ）／1以上＝その系列番号の率を参照
    rateSrc = Array(0, -1, -1, -1, 0, 5, 1, 0, 8)
    sym = Array("A", "B", "C", "D", "E", "F", "G", "H", "I")
    nm = Array("変更設計書" & vbLf & "（全体）", _
               "変更設計書" & vbLf & "残工事", _
               "単価更新設計書" & vbLf & "残工事", _
               "変更設計書" & vbLf & "出来高", _
               "スライド設計書" & vbLf & "（全体・新単価）", _
               "単価更新設計書" & vbLf & "残工事" & vbLf & "【経費計算】", _
               "変更設計書" & vbLf & "出来高" & vbLf & "【経費計算】", _
               "変更設計書" & vbLf & "新工種抜き（全体）", _
               "変更設計書" & vbLf & "新工種抜き 出来高")

    WriteKeihiLabels

    unpan = FindZItem("運搬費")
    gijutsu = FindZItem("技術管理費")

    For i = 0 To K_NSERIES - 1
        c = K_FIRSTCOL + i
        mWs.Cells(2, c).Value = sym(i)
        mWs.Cells(3, c).Value = nm(i)
        WriteSeries c, CLng(srcAmt(i)), CLng(srcSho(i)), CBool(exShin(i)), _
                    CLng(rateSrc(i)), unpan, gijutsu, kojo, zei
    Next i

    FinishKeihi
End Sub


'--------------------------------------------------------------
' 項目名の列
'--------------------------------------------------------------
Private Sub WriteKeihiLabels()
    mWs.Range("A1").Value = "経費計算シート"
    mWs.Range("A1").Font.Size = 14
    mWs.Range("A1").Font.Bold = True
    mWs.Range("B1").Value = "　※黄色セルは手入力。経費率は積算システムの計算結果があれば上書きしてください"
    mWs.Range("B1").Font.Color = RGB(192, 0, 0)

    Sec 4, "直接工事費"
    Lb K_DCHOKU, "①直接工事費"
    Lb K_KANZAI, "①'水道工事における管材費"
    Lb K_SHOBUN, "②直接工事費内の処分費等"
    Lb K_TAIGAI, "③②のうち率計算の対象外費　※②－(①－①'/2)×3%"
    Sec 9, "共通仮設費"
    Lb K_UNPAN, "④共通仮設費積上分－運搬費"
    Lb K_TSUMI2, "④'共通仮設費積上分－その他"
    Lb K_GIJUTSU, "⑤共通仮設費積上分－技術管理費"
    Lb K_KTAISHO, "⑥共通仮設費対象額　※①－③－①'/2"
    Lb K_KRITSU, "⑦共通仮設費率"
    Lb K_KBUN, "⑦'共通仮設費率分　※⑥×⑦"
    Lb K_KKEI, "⑧共通仮設費【合計】　※④＋④'＋⑤＋⑦'"
    Lb K_JUN, "⑨純工事費　※①＋⑧"
    Sec 18, "現場管理費"
    Lb K_GTAISHO, "⑩現場管理費対象額　※⑨－③－①'/2"
    Lb K_GRITSU, "⑪現場管理費率"
    Lb K_GKEI, "⑫現場管理費【合計】　※⑩×⑪"
    Lb K_GENKA, "⑬工事原価　※⑨＋⑫"
    Sec 23, "一般管理費"
    Lb K_ITAISHO, "⑭一般管理費対象額　※⑬－③"
    Lb K_IRITSU, "⑮一般管理費率"
    Lb K_IBUN, "⑮'一般管理費率分　※⑭×⑮"
    Lb K_HOSHO, "⑯契約保証費　※当初設計時額で固定"
    Lb K_IKEI0, "⑰一般管理費【合計】（端数整理前）　※⑮'＋⑯"
    Lb K_KAKAKU0, "⑱工事価格（端数整理前）　※⑬＋⑰"
    Lb K_HASU, "⑲端数整理"
    Lb K_IKEI, "⑳一般管理費【合計】（端数整理後）　※⑰－⑲"
    Lb K_KAKAKU, "㉑工事価格　※⑬＋⑳"
    Sec 33, "工事費"
    Lb K_SCRAP, "㉒スクラップ　※スライド計算表のスクラップ〇から集計"
    Lb K_KAKAKU2, "㉑'工事価格（スクラップ込み）　※㉑＋㉒"
    Lb K_ZEI, "㉓消費税相当額　※(㉑＋㉒)×消費税率"
    Lb K_KOUJIHI, "㉔工事費　※㉑＋㉒＋㉓"

    mWs.Cells(39, 2).Value = "※ B/C/D列は素材（直接工事費・積上分）のみ。経費はF/G列で計算します"
    mWs.Cells(40, 2).Value = "※ 契約保証費は当初積算時から変更しません（手入力）"
    mWs.Cells(41, 2).Value = "※ ㉑'工事価格（スクラップ込み）を スライド調書（様式4-2号）へ転記します"
    mWs.Range("B39:B41").Font.Color = RGB(120, 120, 120)
End Sub


Private Sub Sec(ByVal r As Long, ByVal s As String)
    mWs.Cells(r, 1).Value = s
    With mWs.Range(mWs.Cells(r, 1), mWs.Cells(r, K_FIRSTCOL + K_NSERIES - 1))
        .Interior.Color = CLR_HEAD
        .Font.Bold = True
    End With
End Sub


Private Sub Lb(ByVal r As Long, ByVal s As String)
    mWs.Cells(r, 2).Value = s
End Sub


'==============================================================
' 1系列ぶんの数式
'==============================================================
Private Sub WriteSeries(ByVal c As Long, ByVal amtCol As Long, ByVal shoCol As Long, _
                        ByVal exShin As Boolean, ByVal rateSrc As Long, _
                        ByVal unpan As String, ByVal gijutsu As String, _
                        ByVal kojo As String, ByVal zei As String)
    Dim d1 As Long, d2 As Long, a1 As Long, a2 As Long
    Dim shinCond As String
    Dim rc As Long

    d1 = gDirectFirst: d2 = gDirectLast
    ' 新工種を除く条件（SUMIFSの "<>" は環境差が出るため SUMPRODUCT で書く）
    shinCond = ""
    If exShin Then shinCond = "*(" & SR(SC_MK_SHIN, d1, d2) & "<>""〇"")"
    ' スクラップは㉒として工事価格の後に足すので、直接工事費側からは外す
    shinCond = shinCond & "*(" & SR(SC_MK_SCRAP, d1, d2) & "<>""〇"")"

    '--- 素材 ---
    ' ①直接工事費（明細行のみ。階層行は単価が空なので除外される）
    mWs.Cells(K_DCHOKU, c).Formula = "=SUMPRODUCT((" & SR(SC_TANKA_O, d1, d2) & "<>"""")" & _
        shinCond & "*(" & SR(amtCol, d1, d2) & "))"

    ' ①'管材費
    mWs.Cells(K_KANZAI, c).Formula = "=SUMPRODUCT((" & SR(SC_MK_KANZA, d1, d2) & "=""〇"")" & _
        shinCond & "*(" & SR(amtCol, d1, d2) & "))"

    ' ②処分費等
    mWs.Cells(K_SHOBUN, c).Formula = "=SUMPRODUCT((" & SR(SC_TANKA_O, d1, d2) & "<>"""")" & _
        shinCond & "*(" & SR(shoCol, d1, d2) & "))"

    ' ③率計算の対象外費
    mWs.Cells(K_TAIGAI, c).Formula = "=MAX(0," & Cel(K_SHOBUN, c) & "-ROUNDDOWN((" & _
        Cel(K_DCHOKU, c) & "-" & Cel(K_KANZAI, c) & "/2)*" & kojo & ",0))"

    ' ④運搬費／⑤技術管理費／④'その他
    mWs.Cells(K_UNPAN, c).Formula = ZSum(unpan, amtCol, exShin)
    mWs.Cells(K_GIJUTSU, c).Formula = ZSum(gijutsu, amtCol, exShin)
    If gZFirst > 0 Then
        mWs.Cells(K_TSUMI2, c).Formula = "=SUMPRODUCT((" & SR(SC_TANKA_O, gZFirst, gZLast) & "<>"""")" & _
            IIf(exShin, "*(" & SR(SC_MK_SHIN, gZFirst, gZLast) & "<>""〇"")", "") & _
            "*(" & SR(SC_MK_SCRAP, gZFirst, gZLast) & "<>""〇"")" & _
            "*(" & SR(amtCol, gZFirst, gZLast) & "))-" & Cel(K_UNPAN, c) & "-" & Cel(K_GIJUTSU, c)
    Else
        mWs.Cells(K_TSUMI2, c).Value = 0
    End If

    '--- 素材だけの系列はここまで ---
    If rateSrc < 0 Then
        mWs.Cells(K_KTAISHO, c).Value = "直接工事費・積上分のみ使用"
        mWs.Range(mWs.Cells(K_KTAISHO, c), mWs.Cells(K_KOUJIHI, c)).Merge
        With mWs.Range(mWs.Cells(K_KTAISHO, c), mWs.Cells(K_KOUJIHI, c))
            .Orientation = xlVertical
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Color = RGB(120, 120, 120)
        End With
        Exit Sub
    End If

    '--- 共通仮設費 ---
    mWs.Cells(K_KTAISHO, c).Formula = "=" & Cel(K_DCHOKU, c) & "-" & Cel(K_TAIGAI, c) & _
                                      "-" & Cel(K_KANZAI, c) & "/2"

    If rateSrc = 0 Then
        mWs.Cells(K_KRITSU, c).Formula = RateKasetsu(Cel(K_KTAISHO, c))
        MarkInput K_KRITSU, c
    Else
        rc = K_FIRSTCOL + rateSrc - 1
        mWs.Cells(K_KRITSU, c).Formula = "=" & Cel(K_KRITSU, rc)
    End If

    mWs.Cells(K_KBUN, c).Formula = "=ROUNDDOWN(" & Cel(K_KTAISHO, c) & "*" & Cel(K_KRITSU, c) & ",-3)"
    mWs.Cells(K_KKEI, c).Formula = "=" & Cel(K_UNPAN, c) & "+" & Cel(K_TSUMI2, c) & _
                                   "+" & Cel(K_GIJUTSU, c) & "+" & Cel(K_KBUN, c)
    mWs.Cells(K_JUN, c).Formula = "=" & Cel(K_DCHOKU, c) & "+" & Cel(K_KKEI, c)

    '--- 現場管理費 ---
    mWs.Cells(K_GTAISHO, c).Formula = "=" & Cel(K_JUN, c) & "-" & Cel(K_TAIGAI, c) & _
                                      "-" & Cel(K_KANZAI, c) & "/2"
    If rateSrc = 0 Then
        mWs.Cells(K_GRITSU, c).Formula = RateGenba(Cel(K_GTAISHO, c))
        MarkInput K_GRITSU, c
    Else
        rc = K_FIRSTCOL + rateSrc - 1
        mWs.Cells(K_GRITSU, c).Formula = "=" & Cel(K_GRITSU, rc)
    End If
    mWs.Cells(K_GKEI, c).Formula = "=ROUNDDOWN(" & Cel(K_GTAISHO, c) & "*" & Cel(K_GRITSU, c) & ",-3)"
    mWs.Cells(K_GENKA, c).Formula = "=" & Cel(K_JUN, c) & "+" & Cel(K_GKEI, c)

    '--- 一般管理費 ---
    mWs.Cells(K_ITAISHO, c).Formula = "=" & Cel(K_GENKA, c) & "-" & Cel(K_TAIGAI, c)
    If rateSrc = 0 Then
        mWs.Cells(K_IRITSU, c).Formula = RateIppan(Cel(K_ITAISHO, c))
        MarkInput K_IRITSU, c
    Else
        rc = K_FIRSTCOL + rateSrc - 1
        mWs.Cells(K_IRITSU, c).Formula = "=" & Cel(K_IRITSU, rc)
    End If
    mWs.Cells(K_IBUN, c).Formula = "=" & Cel(K_ITAISHO, c) & "*" & Cel(K_IRITSU, c)
    mWs.Cells(K_HOSHO, c).Value = CDbl(CfgVal(mCfg, "契約保証費", 0))
    MarkInput K_HOSHO, c
    mWs.Cells(K_IKEI0, c).Formula = "=" & Cel(K_IBUN, c) & "+" & Cel(K_HOSHO, c)
    mWs.Cells(K_KAKAKU0, c).Formula = "=" & Cel(K_GENKA, c) & "+" & Cel(K_IKEI0, c)
    mWs.Cells(K_HASU, c).Formula = "=" & Cel(K_KAKAKU0, c) & "-ROUNDDOWN(" & Cel(K_KAKAKU0, c) & ",-3)"
    mWs.Cells(K_IKEI, c).Formula = "=" & Cel(K_IKEI0, c) & "-" & Cel(K_HASU, c)
    mWs.Cells(K_KAKAKU, c).Formula = "=" & Cel(K_GENKA, c) & "+" & Cel(K_IKEI, c)

    '--- 工事費 ---
    ' ㉒スクラップ：スクラップ〇の行をCSVから集計する
    a1 = gDirectFirst
    a2 = IIf(gZLast > 0, gZLast, gDirectLast)
    mWs.Cells(K_SCRAP, c).Formula = "=SUMPRODUCT((" & SR(SC_TANKA_O, a1, a2) & "<>"""")*(" & _
        SR(SC_MK_SCRAP, a1, a2) & "=""〇"")" & _
        IIf(exShin, "*(" & SR(SC_MK_SHIN, a1, a2) & "<>""〇"")", "") & _
        "*(" & SR(amtCol, a1, a2) & "))"
    mWs.Cells(K_KAKAKU2, c).Formula = "=" & Cel(K_KAKAKU, c) & "+" & Cel(K_SCRAP, c)
    mWs.Cells(K_ZEI, c).Formula = "=(" & Cel(K_KAKAKU, c) & "+" & Cel(K_SCRAP, c) & ")*" & zei
    mWs.Cells(K_KOUJIHI, c).Formula = "=" & Cel(K_KAKAKU, c) & "+" & Cel(K_SCRAP, c) & "+" & Cel(K_ZEI, c)
End Sub


'==============================================================
' 経費率の自動計算（積算システムの値があれば上書きしてください）
'==============================================================
Private Function RateKasetsu(ByVal taisho As String) As String
    Dim a As String, b As String, ch As String, sh As String
    a = NumStr(CfgVal(mCfg, "共通仮設費率A", 1228.3))
    b = NumStr(CfgVal(mCfg, "共通仮設費率B", -0.2614))
    ch = NumStr(CfgVal(mCfg, "共通仮設費 地域補正", 1.2))
    sh = NumStr(CfgVal(mCfg, "共通仮設費 週休補正", 1.04))
    RateKasetsu = "=IF(" & taisho & "<=0,0,ROUND(ROUND(ROUND(" & a & "*" & taisho & _
                  "^(" & b & "),2)*" & ch & ",2)*" & sh & ",2)/100)"
End Function


Private Function RateGenba(ByVal taisho As String) As String
    Dim a As String, b As String, ch As String, sh As String
    a = NumStr(CfgVal(mCfg, "現場管理費率A", 458.2))
    b = NumStr(CfgVal(mCfg, "現場管理費率B", -0.1508))
    ch = NumStr(CfgVal(mCfg, "現場管理費 地域補正", 1.1))
    sh = NumStr(CfgVal(mCfg, "現場管理費 週休補正", 1.06))
    RateGenba = "=IF(" & taisho & "<=0,0,ROUND(ROUND(ROUND(" & a & "*" & taisho & _
                "^(" & b & "),2)*" & ch & ",2)*" & sh & ",2)/100)"
End Function


Private Function RateIppan(ByVal taisho As String) As String
    Dim k As String, t As String
    k = NumStr(CfgVal(mCfg, "一般管理費率係数", -5.48972))
    t = NumStr(CfgVal(mCfg, "一般管理費率定数", 59.4977))
    RateIppan = "=IF(" & taisho & "<=0,0,ROUND(" & k & "*LOG10(" & taisho & ")+" & t & ",2)/100)"
End Function


'==============================================================
' 補助
'==============================================================
Private Function SR(ByVal c As Long, ByVal r1 As Long, ByVal r2 As Long) As String
    If r2 < r1 Then r2 = r1
    SR = "'" & SH_SLIDE & "'!" & mSlide.Cells(r1, c).Address(True, True) & ":" & _
         mSlide.Cells(r2, c).Address(True, True)
End Function


Private Function Cel(ByVal r As Long, ByVal c As Long) As String
    Cel = mWs.Cells(r, c).Address(False, False)
End Function


' gZItems から名称を含む項目の行範囲を探す（"r1|r2" を返す）
Private Function FindZItem(ByVal keyword As String) As String
    Dim k As Variant
    If gZItems Is Nothing Then Exit Function
    For Each k In gZItems.Keys
        If InStr(1, NormText(CStr(k)), NormText(keyword)) > 0 Then
            FindZItem = gZItems(k)
            Exit Function
        End If
    Next k
End Function


Private Function ZSum(ByVal rangeSpec As String, ByVal amtCol As Long, ByVal exShin As Boolean) As String
    Dim p() As String, r1 As Long, r2 As Long

    If rangeSpec = "" Then
        ZSum = "=0"
        Exit Function
    End If

    p = Split(rangeSpec, "|")
    r1 = CLng(p(0)): r2 = CLng(p(1))

    ZSum = "=SUMPRODUCT((" & SR(SC_TANKA_O, r1, r2) & "<>"""")" & _
           IIf(exShin, "*(" & SR(SC_MK_SHIN, r1, r2) & "<>""〇"")", "") & _
           "*(" & SR(SC_MK_SCRAP, r1, r2) & "<>""〇"")" & _
           "*(" & SR(amtCol, r1, r2) & "))"
End Function


Private Sub MarkInput(ByVal r As Long, ByVal c As Long)
    mWs.Cells(r, c).Interior.Color = CLR_INPUT
End Sub


Private Function NumStr(ByVal v As Variant) As String
    NumStr = Format$(CDbl(v), "0.##########")
End Function


Private Sub FinishKeihi()
    Dim lastC As Long
    lastC = K_FIRSTCOL + K_NSERIES - 1

    With mWs.Range(mWs.Cells(2, 1), mWs.Cells(K_KOUJIHI, lastC))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    With mWs.Range(mWs.Cells(2, K_FIRSTCOL), mWs.Cells(3, lastC))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = CLR_HEAD
    End With

    mWs.Range(mWs.Cells(5, K_FIRSTCOL), mWs.Cells(K_KOUJIHI, lastC)).NumberFormatLocal = "#,##0"
    mWs.Range(mWs.Cells(K_KRITSU, K_FIRSTCOL), mWs.Cells(K_KRITSU, lastC)).NumberFormatLocal = "0.0000"
    mWs.Range(mWs.Cells(K_GRITSU, K_FIRSTCOL), mWs.Cells(K_GRITSU, lastC)).NumberFormatLocal = "0.0000"
    mWs.Range(mWs.Cells(K_IRITSU, K_FIRSTCOL), mWs.Cells(K_IRITSU, lastC)).NumberFormatLocal = "0.0000"

    HiLite K_KAKAKU2, lastC
    HiLite K_KOUJIHI, lastC

    mWs.Columns(1).ColumnWidth = 12
    mWs.Columns(2).ColumnWidth = 42
    mWs.Range(mWs.Columns(K_FIRSTCOL), mWs.Columns(lastC)).ColumnWidth = 15
    mWs.Rows(3).RowHeight = 54

    With mWs.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
End Sub


Private Sub HiLite(ByVal r As Long, ByVal lastC As Long)
    With mWs.Range(mWs.Cells(r, 2), mWs.Cells(r, lastC))
        .Font.Bold = True
        .Interior.Color = CLR_TOTAL
    End With
End Sub


'==============================================================
' スライド調書（様式4-2号）
'==============================================================
Public Sub BuildChosho(ByVal cfg As Object)
    Dim ws As Worksheet
    Dim K As String
    Dim contract As Double

    Set mCfg = cfg
    Set ws = FreshSheet(SH_CHOSHO)
    K = "'" & SH_KEIHI & "'!"
    contract = CDbl(CfgVal(cfg, "請負代金額", 0))

    ws.Range("G2").Value = "様式4-2号"
    ws.Range("B3").Value = "スライド調書"
    ws.Range("B3").Font.Size = 14
    ws.Range("B3").Font.Bold = True
    ws.Range("E3").Formula = "=IF(E16>F16,""減額スライド"",IF(E16<F16,""増額スライド"",""""))"
    ws.Range("B4").Value = CfgVal(cfg, "工事名", "")

    ws.Range("C5").Value = "元設計"
    ws.Range("D5").Value = "出来高"
    ws.Range("E5").Value = "残工事"
    ws.Range("G5").Value = "スライド額"
    ws.Range("E6").Value = "変動前"
    ws.Range("F6").Value = "変動後"

    ws.Range("C7").Value = "①"
    ws.Range("B8").Value = "設計額（税込）"
    ws.Range("C8").Formula = "=" & K & "C" & K_KOUJIHI
    ws.Range("H8").Value = "請負率(%)"
    If contract > 0 Then
        ws.Range("I8").Formula = "=" & contract & "/C8*100"
    Else
        ws.Range("I8").Value = 100
    End If
    ws.Range("I8").Interior.Color = CLR_INPUT

    ws.Range("C9").Value = "②": ws.Range("D9").Value = "⑤"
    ws.Range("E9").Value = "⑦=②-⑤": ws.Range("F9").Value = "⑧"
    ws.Range("B10").Value = "　工事価格（税抜）"
    ws.Range("C10").Formula = "=" & K & "C" & K_KAKAKU2
    ws.Range("D10").Formula = "=" & K & "I" & K_KAKAKU2      ' G列＝変更設計書 出来高
    ws.Range("E10").Formula = "=C10-D10"
    ws.Range("F10").Formula = "=" & K & "H" & K_KAKAKU2      ' F列＝単価更新設計書 残工事

    ws.Range("B12").Value = "　消費税相当額"
    ws.Range("C12").Formula = "=" & K & "C" & K_ZEI

    ws.Range("C13").Value = "③"
    ws.Range("B14").Value = "請負代金額（税込）"
    ws.Range("C14").Formula = "=ROUNDDOWN(C8*I8/100,0)"

    ws.Range("B15").Value = "　請負工事価格"
    ws.Range("C15").Value = "④"
    ws.Range("D15").Value = "⑥=⑤×（③/①）"
    ws.Range("E15").Value = "P1'=④-⑥"
    ws.Range("F15").Value = "P2'=⑧×（③/①）"
    ws.Range("G15").Formula = "=IF(E16>F16,""S'=P2'-P1'+（P1''×1/100）"",""S'=P2'-P1'-（P1''×1/100）"")"

    ws.Range("B16").Value = "　（請負代金額（税抜））"
    ws.Range("C16").Formula = "=ROUNDDOWN(C14*10/11,0)"
    ws.Range("D16").Formula = "=ROUNDDOWN(D10*C14/C8,0)"
    ws.Range("E16").Formula = "=C16-D16"
    ws.Range("F16").Formula = "=ROUNDDOWN(F10*C14/C8,0)"
    ' 増額スライドは1%を控除、減額スライドは1%を戻す
    ws.Range("G16").Formula = "=IF(E16>F16,F16-E16+(E20*1/100),F16-E16-(E20*1/100))"

    ws.Range("B18").Value = "　消費税相当額"
    ws.Range("C18").Formula = "=C14-C16"

    ws.Range("B19").Value = "　請負工事価格" & vbLf & "（新工種抜き）"
    ws.Range("C19").Value = "⑨" & vbLf & "（請負率考慮）"
    ws.Range("D19").Value = "⑩" & vbLf & "（請負率考慮）"
    ws.Range("E19").Value = "P1''=⑨-⑩"
    ws.Range("B20").Value = "　（請負代金額（税抜））"
    ws.Range("C20").Formula = "=ROUNDDOWN(" & K & "J" & K_KAKAKU2 & "*C14/C8,0)"   ' H列＝新工種抜き 全体
    ws.Range("D20").Formula = "=ROUNDDOWN(" & K & "K" & K_KAKAKU2 & "*C14/C8,0)"   ' I列＝新工種抜き 出来高
    ws.Range("E20").Formula = "=C20-D20"

    ws.Range("B22").Value = "※ P1''（新工種を抜いた請負ベースの残工事）が受注者負担1%の母数です"
    ws.Range("B22").Font.Color = RGB(120, 120, 120)

    '--- 体裁 ---
    ws.Range("C5:C6").Merge
    ws.Range("D5:D6").Merge
    ws.Range("E5:F5").Merge
    ws.Range("G5:G6").Merge
    ws.Range("C5:G6").HorizontalAlignment = xlCenter
    ws.Range("C5:G6").Interior.Color = CLR_HEAD
    ws.Range("C5:G6").Font.Bold = True

    ws.Range("B8:G20").Borders.LineStyle = xlContinuous
    ws.Range("C5:G6").Borders.LineStyle = xlContinuous
    ws.Range("C8:G20").NumberFormatLocal = "#,##0"
    ws.Range("C7:G7").NumberFormatLocal = "@"
    ws.Range("C9:G9").NumberFormatLocal = "@"
    ws.Range("C13:G13").NumberFormatLocal = "@"
    ws.Range("C15:G15").NumberFormatLocal = "@"
    ws.Range("C19:G19").NumberFormatLocal = "@"
    ws.Range("I8").NumberFormatLocal = "0.0000"

    With ws.Range("G16")
        .Font.Bold = True
        .Interior.Color = CLR_TOTAL
    End With
    ws.Range("B19").WrapText = True
    ws.Range("C19:D19").WrapText = True

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Range(ws.Columns("C"), ws.Columns("G")).ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I").ColumnWidth = 14
    ws.Rows(19).RowHeight = 34

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
End Sub
