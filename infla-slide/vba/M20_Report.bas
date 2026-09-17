'==============================================================
' M20_Report  ―  設定シート生成・設計書シートの出力と体裁
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M20_Report" にしてください。
'==============================================================
Option Explicit

' 突合結果の列インデックス
Private Const C_KOSHU    As Long = 1
Private Const C_SHUBETSU As Long = 2
Private Const C_SAIBETSU As Long = 3
Private Const C_KIKAKU   As Long = 4
Private Const C_TANI     As Long = 5
Private Const C_OSURYO   As Long = 6    ' 変動前 数量
Private Const C_OTANKA   As Long = 7    ' 変動前 単価
Private Const C_OKINGAKU As Long = 8    ' 変動前 金額
Private Const C_NSURYO   As Long = 9    ' 変動後 数量
Private Const C_NTANKA   As Long = 10   ' 変動後 単価
Private Const C_NKINGAKU As Long = 11   ' 変動後 金額
Private Const C_SAGAKU   As Long = 12   ' 差額
Private Const C_LAST     As Long = 12


'==============================================================
' 設定シートを作成する（既存があれば作り直し）
'==============================================================
Public Sub 設定シート作成()
    Dim ws As Worksheet
    Dim r As Long

    If SheetExists(SH_CONFIG) Then
        If MsgBox("「" & SH_CONFIG & "」シートを作り直します。" & vbCrLf & _
                  "入力済みの内容は消えます。よろしいですか？", _
                  vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = FreshSheet(SH_CONFIG)

    With ws.Range("A1")
        .Value = "インフレスライド設計書 作成設定"
        .Font.Size = 14
        .Font.Bold = True
    End With

    r = 3
    PutHead ws, r, "【工事情報】":                                    r = r + 1
    PutItem ws, r, "工事名", "":                                       r = r + 1
    PutItem ws, r, "工事場所", "":                                     r = r + 1
    PutItem ws, r, "工期（自）", "":                                   r = r + 1
    PutItem ws, r, "工期（至）", "":                                   r = r + 1
    PutItem ws, r, "請負代金額", 0:                                    r = r + 1
    PutItem ws, r, "基準日", "":                                       r = r + 1
    PutItem ws, r, "発注者", "":                                       r = r + 1
    PutItem ws, r, "受注者", "":                                       r = r + 2

    PutHead ws, r, "【CSV設定】":                                      r = r + 1
    PutItem ws, r, "変動前CSVパス", "", "空欄ならマクロ実行時にダイアログで選択": r = r + 1
    PutItem ws, r, "変動後CSVパス", "", "空欄ならマクロ実行時にダイアログで選択": r = r + 1
    PutItem ws, r, "文字コード", "Shift_JIS", "UTF-8 の場合は UTF-8 と入力":      r = r + 1
    PutItem ws, r, "見出し行数", 1, "データが始まる前の行数":                     r = r + 1
    PutItem ws, r, "区切り文字", ",", "タブ区切りなら vbTab は不可。カンマ推奨":  r = r + 2

    PutHead ws, r, "【列マッピング】", "CSVの何列目か（1＝A列）。使わない項目は 0": r = r + 1
    PutItem ws, r, "工種列", 1:                                        r = r + 1
    PutItem ws, r, "種別列", 2:                                        r = r + 1
    PutItem ws, r, "細別列", 3:                                        r = r + 1
    PutItem ws, r, "規格列", 4:                                        r = r + 1
    PutItem ws, r, "単位列", 5:                                        r = r + 1
    PutItem ws, r, "数量列", 6:                                        r = r + 1
    PutItem ws, r, "単価列", 7:                                        r = r + 1
    PutItem ws, r, "金額列", 8, "空欄の行は 数量×単価 で補完":         r = r + 1
    PutItem ws, r, "摘要列", 9:                                        r = r + 2

    PutHead ws, r, "【計算設定】":                                     r = r + 1
    PutItem ws, r, "受注者負担率", 0.01, "請負代金額に対する受注者負担（通常 1%）": r = r + 1
    PutItem ws, r, "端数処理", "切り捨て", "切り捨て／切り上げ／四捨五入":          r = r + 1
    PutItem ws, r, "端数単位", 1, "1／10／100／1000":                   r = r + 2

    PutHead ws, r, "除外語", "この語と一致する行は集計から除外（合計・小計行の二重計上防止）"
    ws.Cells(r, 2).Value = "合計":     r = r + 1
    ws.Cells(r, 2).Value = "小計":     r = r + 1
    ws.Cells(r, 2).Value = "総計":     r = r + 1
    ws.Cells(r, 2).Value = "計":       r = r + 1
    ws.Cells(r, 2).Value = "直接工事費"

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 50
    ws.Range("B:B").HorizontalAlignment = xlLeft
    ws.Range("C:C").Font.Color = RGB(120, 120, 120)

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("B4").Select

    MsgBox "「" & SH_CONFIG & "」シートを作成しました。" & vbCrLf & _
           "工事情報と列マッピングを入力してから" & vbCrLf & _
           "［インフレスライド設計書作成］を実行してください。", vbInformation
End Sub


Private Sub PutHead(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    Optional ByVal note As String = "")
    With ws.Cells(r, 1)
        .Value = label
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    If note <> "" Then ws.Cells(r, 3).Value = note
End Sub


Private Sub PutItem(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByVal defaultValue As Variant, Optional ByVal note As String = "")
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = defaultValue
    ws.Cells(r, 2).Interior.Color = RGB(255, 255, 204)
    ws.Cells(r, 2).Borders.LineStyle = xlContinuous
    If note <> "" Then ws.Cells(r, 3).Value = note
End Sub


'==============================================================
' 設計書シート一式を作成する
'==============================================================
Public Sub BuildReport(ByVal cfg As Object, ByVal dOld As Object, ByVal dNew As Object)
    Dim m As Variant
    Dim p1 As Double, p2 As Double

    m = MergeMeisai(dOld, dNew)

    p1 = SumCol(m, C_OKINGAKU)
    p2 = SumCol(m, C_NKINGAKU)

    MakeHyoushi cfg, p1, p2
    MakeSoukatsu cfg, m, p1, p2
    MakeTaihi cfg, m
    MakeUchiwake cfg, m, True
    MakeUchiwake cfg, m, False

    ThisWorkbook.Worksheets("01_表紙").Activate
End Sub


'--------------------------------------------------------------
' 変動前・変動後を突合して2次元配列にする
'--------------------------------------------------------------
Private Function MergeMeisai(ByVal dOld As Object, ByVal dNew As Object) As Variant
    Dim uk As Object, k As Variant
    Dim m() As Variant
    Dim i As Long
    Dim rec As Variant

    Set uk = CreateObject("Scripting.Dictionary")
    For Each k In dOld.Keys
        If Not uk.Exists(k) Then uk.Add k, True
    Next k
    For Each k In dNew.Keys
        If Not uk.Exists(k) Then uk.Add k, True
    Next k

    If uk.Count = 0 Then
        MergeMeisai = Empty
        Exit Function
    End If

    ReDim m(1 To uk.Count, 1 To C_LAST)
    i = 0

    For Each k In uk.Keys
        i = i + 1

        If dOld.Exists(k) Then
            rec = dOld(k)
        ElseIf dNew.Exists(k) Then
            rec = dNew(k)
        End If

        m(i, C_KOSHU) = rec(F_KOSHU)
        m(i, C_SHUBETSU) = rec(F_SHUBETSU)
        m(i, C_SAIBETSU) = rec(F_SAIBETSU)
        m(i, C_KIKAKU) = rec(F_KIKAKU)
        m(i, C_TANI) = rec(F_TANI)

        If dOld.Exists(k) Then
            rec = dOld(k)
            m(i, C_OSURYO) = rec(F_SURYO)
            m(i, C_OTANKA) = rec(F_TANKA)
            m(i, C_OKINGAKU) = rec(F_KINGAKU)
        Else
            m(i, C_OSURYO) = 0: m(i, C_OTANKA) = 0: m(i, C_OKINGAKU) = 0
        End If

        If dNew.Exists(k) Then
            rec = dNew(k)
            m(i, C_NSURYO) = rec(F_SURYO)
            m(i, C_NTANKA) = rec(F_TANKA)
            m(i, C_NKINGAKU) = rec(F_KINGAKU)
        Else
            m(i, C_NSURYO) = 0: m(i, C_NTANKA) = 0: m(i, C_NKINGAKU) = 0
        End If

        m(i, C_SAGAKU) = CDbl(m(i, C_NKINGAKU)) - CDbl(m(i, C_OKINGAKU))
    Next k

    MergeMeisai = m
End Function


Private Function SumCol(ByVal m As Variant, ByVal col As Long) As Double
    Dim i As Long, t As Double
    If IsEmpty(m) Then Exit Function
    For i = LBound(m, 1) To UBound(m, 1)
        t = t + CDbl(m(i, col))
    Next i
    SumCol = t
End Function


'==============================================================
' 01_表紙
'==============================================================
Private Sub MakeHyoushi(ByVal cfg As Object, ByVal p1 As Double, ByVal p2 As Double)
    Dim ws As Worksheet
    Dim contract As Double, futan As Double, sagaku As Double, henko As Double
    Dim r As Long

    contract = CDbl(CfgVal(cfg, "請負代金額", 0))
    sagaku = p2 - p1
    futan = RoundAmt(contract * CDbl(CfgVal(cfg, "受注者負担率", 0.01)), _
                     CStr(CfgVal(cfg, "端数処理", "切り捨て")), _
                     CDbl(CfgVal(cfg, "端数単位", 1)))
    henko = sagaku - futan
    If henko < 0 Then henko = 0

    Set ws = FreshSheet("01_表紙")

    With ws.Range("B2")
        .Value = "イ ン フ レ ス ラ イ ド 設 計 書"
        .Font.Size = 18
        .Font.Bold = True
    End With
    ws.Range("B2:F2").Merge
    ws.Range("B2").HorizontalAlignment = xlCenter

    r = 5
    PutPair ws, r, "工 事 名", CfgVal(cfg, "工事名", ""):        r = r + 1
    PutPair ws, r, "工 事 場 所", CfgVal(cfg, "工事場所", ""):   r = r + 1
    PutPair ws, r, "工 期", CStr(CfgVal(cfg, "工期（自）", "")) & "  〜  " & _
                            CStr(CfgVal(cfg, "工期（至）", "")): r = r + 1
    PutPair ws, r, "発 注 者", CfgVal(cfg, "発注者", ""):        r = r + 1
    PutPair ws, r, "受 注 者", CfgVal(cfg, "受注者", ""):        r = r + 1
    PutPair ws, r, "基 準 日", CfgVal(cfg, "基準日", ""):        r = r + 2

    PutPairNum ws, r, "現 請 負 代 金 額", contract:          r = r + 1
    PutPairNum ws, r, "変 動 前 残 工 事 代 金 額 (A)", p1:   r = r + 1
    PutPairNum ws, r, "変 動 後 残 工 事 代 金 額 (B)", p2:   r = r + 1
    PutPairNum ws, r, "ス ラ イ ド 差 額 (B-A)", sagaku:      r = r + 1
    PutPairNum ws, r, "受 注 者 負 担 額", futan:             r = r + 1

    PutPairNum ws, r, "変 更 額", henko
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Bold = True
    r = r + 1

    PutPairNum ws, r, "変 更 後 請 負 代 金 額", contract + henko
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With
    r = r + 2

    ws.Cells(r, 2).Value = "作成日：" & Format$(Date, "yyyy年m月d日")

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 26
    ws.Columns("D:F").ColumnWidth = 12

    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
End Sub


Private Sub PutPair(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal v As Variant)
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Value = v
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Borders.LineStyle = xlContinuous
        .RowHeight = 22
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(r, 2).Interior.Color = RGB(242, 242, 242)
End Sub


Private Sub PutPairNum(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal v As Double)
    PutPair ws, r, label, v
    ws.Cells(r, 3).NumberFormatLocal = "#,##0""円"""
    ws.Cells(r, 3).HorizontalAlignment = xlRight
End Sub


'==============================================================
' 02_総括表（工種別の集計＋スライド額計算）
'==============================================================
Private Sub MakeSoukatsu(ByVal cfg As Object, ByVal m As Variant, _
                         ByVal p1 As Double, ByVal p2 As Double)
    Dim ws As Worksheet
    Dim agg As Object, k As Variant
    Dim i As Long, r As Long, headR As Long
    Dim v As Variant
    Dim contract As Double, futan As Double, sagaku As Double, henko As Double

    Set ws = FreshSheet("02_総括表")

    ws.Range("A1").Value = "インフレスライド　総括表"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "工事名"
    ws.Range("B2").Value = CfgVal(cfg, "工事名", "")
    ws.Range("A3").Value = "基準日"
    ws.Range("B3").Value = CfgVal(cfg, "基準日", "")

    ' 工種別に集計
    Set agg = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(m) Then
        For i = LBound(m, 1) To UBound(m, 1)
            k = CStr(m(i, C_KOSHU))
            If k = "" Then k = "（工種未設定）"
            If agg.Exists(k) Then
                v = agg(k)
            Else
                v = Array(0#, 0#)
            End If
            v(0) = CDbl(v(0)) + CDbl(m(i, C_OKINGAKU))
            v(1) = CDbl(v(1)) + CDbl(m(i, C_NKINGAKU))
            agg(k) = v
        Next i
    End If

    headR = 5
    ws.Cells(headR, 1).Value = "工　種"
    ws.Cells(headR, 2).Value = "変動前金額 (A)"
    ws.Cells(headR, 3).Value = "変動後金額 (B)"
    ws.Cells(headR, 4).Value = "差　額 (B-A)"
    StyleHeader ws, headR, 1, headR, 4

    r = headR
    For Each k In agg.Keys
        r = r + 1
        v = agg(k)
        ws.Cells(r, 1).Value = k
        ws.Cells(r, 2).Value = CDbl(v(0))
        ws.Cells(r, 3).Value = CDbl(v(1))
        ws.Cells(r, 4).Value = CDbl(v(1)) - CDbl(v(0))
    Next k

    r = r + 1
    ws.Cells(r, 1).Value = "合　　計"
    ws.Cells(r, 2).Value = p1
    ws.Cells(r, 3).Value = p2
    ws.Cells(r, 4).Value = p2 - p1
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    ws.Range(ws.Cells(headR, 1), ws.Cells(r, 4)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(headR + 1, 2), ws.Cells(r, 4)).NumberFormatLocal = "#,##0"

    ' スライド額の計算
    contract = CDbl(CfgVal(cfg, "請負代金額", 0))
    sagaku = p2 - p1
    futan = RoundAmt(contract * CDbl(CfgVal(cfg, "受注者負担率", 0.01)), _
                     CStr(CfgVal(cfg, "端数処理", "切り捨て")), _
                     CDbl(CfgVal(cfg, "端数単位", 1)))
    henko = sagaku - futan
    If henko < 0 Then henko = 0

    r = r + 2
    ws.Cells(r, 1).Value = "【スライド額の算出】"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    PutCalc ws, r, "変動前残工事代金額", "A", p1:                         r = r + 1
    PutCalc ws, r, "変動後残工事代金額", "B", p2:                         r = r + 1
    PutCalc ws, r, "スライド差額", "C = B - A", sagaku:                   r = r + 1
    PutCalc ws, r, "現請負代金額", "D", contract:                         r = r + 1
    PutCalc ws, r, "受注者負担額", "E = D × " & _
            Format$(CDbl(CfgVal(cfg, "受注者負担率", 0.01)) * 100, "0.#") & "%", futan: r = r + 1
    PutCalc ws, r, "変　更　額", "F = C - E", henko
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
    r = r + 1
    PutCalc ws, r, "変更後請負代金額", "D + F", contract + henko
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With

    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 18

    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
End Sub


Private Sub PutCalc(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByVal expr As String, ByVal v As Double)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = expr
    ws.Cells(r, 3).Value = v
    ws.Cells(r, 3).NumberFormatLocal = "#,##0"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(r, 2).HorizontalAlignment = xlCenter
    ws.Cells(r, 2).Font.Color = RGB(120, 120, 120)
End Sub


'==============================================================
' 03_対比表（変動前・変動後を左右に並べる）
'==============================================================
Private Sub MakeTaihi(ByVal cfg As Object, ByVal m As Variant)
    Dim ws As Worksheet
    Dim i As Long, r As Long, h1 As Long, h2 As Long, lastR As Long

    Set ws = FreshSheet("03_対比表")

    ws.Range("A1").Value = "インフレスライド　変動前・変動後 対比表"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "工事名：" & CStr(CfgVal(cfg, "工事名", "")) & _
                           "　　基準日：" & CStr(CfgVal(cfg, "基準日", ""))

    h1 = 4: h2 = 5

    ws.Cells(h1, C_KOSHU).Value = "工　種"
    ws.Cells(h1, C_SHUBETSU).Value = "種　別"
    ws.Cells(h1, C_SAIBETSU).Value = "細　別"
    ws.Cells(h1, C_KIKAKU).Value = "規　格"
    ws.Cells(h1, C_TANI).Value = "単位"
    ws.Cells(h1, C_OSURYO).Value = "変動前（旧単価）"
    ws.Cells(h1, C_NSURYO).Value = "変動後（新単価）"
    ws.Cells(h1, C_SAGAKU).Value = "差　額"

    ws.Range(ws.Cells(h1, C_KOSHU), ws.Cells(h2, C_KOSHU)).Merge
    ws.Range(ws.Cells(h1, C_SHUBETSU), ws.Cells(h2, C_SHUBETSU)).Merge
    ws.Range(ws.Cells(h1, C_SAIBETSU), ws.Cells(h2, C_SAIBETSU)).Merge
    ws.Range(ws.Cells(h1, C_KIKAKU), ws.Cells(h2, C_KIKAKU)).Merge
    ws.Range(ws.Cells(h1, C_TANI), ws.Cells(h2, C_TANI)).Merge
    ws.Range(ws.Cells(h1, C_OSURYO), ws.Cells(h1, C_OKINGAKU)).Merge
    ws.Range(ws.Cells(h1, C_NSURYO), ws.Cells(h1, C_NKINGAKU)).Merge
    ws.Range(ws.Cells(h1, C_SAGAKU), ws.Cells(h2, C_SAGAKU)).Merge

    ws.Cells(h2, C_OSURYO).Value = "数量"
    ws.Cells(h2, C_OTANKA).Value = "単価"
    ws.Cells(h2, C_OKINGAKU).Value = "金額"
    ws.Cells(h2, C_NSURYO).Value = "数量"
    ws.Cells(h2, C_NTANKA).Value = "単価"
    ws.Cells(h2, C_NKINGAKU).Value = "金額"

    StyleHeader ws, h1, 1, h2, C_LAST

    r = h2
    If Not IsEmpty(m) Then
        r = h2 + 1
        ws.Cells(r, 1).Resize(UBound(m, 1), C_LAST).Value = m
        r = r + UBound(m, 1) - 1
    End If
    lastR = r

    ' 合計行
    r = r + 1
    ws.Cells(r, C_KOSHU).Value = "合　　計"
    If lastR >= h2 + 1 Then
        ws.Cells(r, C_OKINGAKU).Formula = "=SUM(" & ws.Cells(h2 + 1, C_OKINGAKU).Address(False, False) & ":" & ws.Cells(lastR, C_OKINGAKU).Address(False, False) & ")"
        ws.Cells(r, C_NKINGAKU).Formula = "=SUM(" & ws.Cells(h2 + 1, C_NKINGAKU).Address(False, False) & ":" & ws.Cells(lastR, C_NKINGAKU).Address(False, False) & ")"
        ws.Cells(r, C_SAGAKU).Formula = "=SUM(" & ws.Cells(h2 + 1, C_SAGAKU).Address(False, False) & ":" & ws.Cells(lastR, C_SAGAKU).Address(False, False) & ")"
    End If
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    FinishTable ws, h1, r, h2
End Sub


'==============================================================
' 04_変動前内訳書 / 05_変動後内訳書
'==============================================================
Private Sub MakeUchiwake(ByVal cfg As Object, ByVal m As Variant, ByVal isOld As Boolean)
    Dim ws As Worksheet
    Dim i As Long, r As Long, h As Long, firstR As Long
    Dim cS As Long, cT As Long, cK As Long
    Dim sheetName As String, title As String

    If isOld Then
        sheetName = "04_変動前内訳書": title = "変動前（旧単価）　内訳書"
        cS = C_OSURYO: cT = C_OTANKA: cK = C_OKINGAKU
    Else
        sheetName = "05_変動後内訳書": title = "変動後（新単価）　内訳書"
        cS = C_NSURYO: cT = C_NTANKA: cK = C_NKINGAKU
    End If

    Set ws = FreshSheet(sheetName)

    ws.Range("A1").Value = title
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "工事名：" & CStr(CfgVal(cfg, "工事名", "")) & _
                           "　　基準日：" & CStr(CfgVal(cfg, "基準日", ""))

    h = 4
    ws.Cells(h, 1).Value = "工　種"
    ws.Cells(h, 2).Value = "種　別"
    ws.Cells(h, 3).Value = "細　別"
    ws.Cells(h, 4).Value = "規　格"
    ws.Cells(h, 5).Value = "単位"
    ws.Cells(h, 6).Value = "数量"
    ws.Cells(h, 7).Value = "単価"
    ws.Cells(h, 8).Value = "金額"
    StyleHeader ws, h, 1, h, 8

    r = h
    firstR = h + 1
    If Not IsEmpty(m) Then
        For i = LBound(m, 1) To UBound(m, 1)
            If CDbl(m(i, cK)) <> 0 Or CDbl(m(i, cS)) <> 0 Then
                r = r + 1
                ws.Cells(r, 1).Value = m(i, C_KOSHU)
                ws.Cells(r, 2).Value = m(i, C_SHUBETSU)
                ws.Cells(r, 3).Value = m(i, C_SAIBETSU)
                ws.Cells(r, 4).Value = m(i, C_KIKAKU)
                ws.Cells(r, 5).Value = m(i, C_TANI)
                ws.Cells(r, 6).Value = m(i, cS)
                ws.Cells(r, 7).Value = m(i, cT)
                ws.Cells(r, 8).Value = m(i, cK)
            End If
        Next i
    End If

    r = r + 1
    ws.Cells(r, 1).Value = "合　　計"
    If r > firstR Then
        ws.Cells(r, 8).Formula = "=SUM(" & ws.Cells(firstR, 8).Address(False, False) & ":" & ws.Cells(r - 1, 8).Address(False, False) & ")"
    End If
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With

    ws.Range(ws.Cells(h, 1), ws.Cells(r, 8)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(h + 1, 6), ws.Cells(r, 6)).NumberFormatLocal = "#,##0.00"
    ws.Range(ws.Cells(h + 1, 7), ws.Cells(r, 8)).NumberFormatLocal = "#,##0"

    ws.Columns("A:D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 6
    ws.Columns("F:H").ColumnWidth = 12
    ws.rows(h).WrapText = True

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$" & h & ":$" & h
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
    End With
    ws.Activate
    ws.Range("A" & (h + 1)).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub


'==============================================================
' 体裁ヘルパ
'==============================================================
Private Sub StyleHeader(ByVal ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, _
                        ByVal r2 As Long, ByVal c2 As Long)
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(221, 235, 247)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
    End With
End Sub


Private Sub FinishTable(ByVal ws As Worksheet, ByVal headR As Long, ByVal lastR As Long, _
                        ByVal headLastR As Long)
    ws.Range(ws.Cells(headR, 1), ws.Cells(lastR, C_LAST)).Borders.LineStyle = xlContinuous

    ws.Range(ws.Cells(headLastR + 1, C_OSURYO), ws.Cells(lastR, C_OSURYO)).NumberFormatLocal = "#,##0.00"
    ws.Range(ws.Cells(headLastR + 1, C_NSURYO), ws.Cells(lastR, C_NSURYO)).NumberFormatLocal = "#,##0.00"
    ws.Range(ws.Cells(headLastR + 1, C_OTANKA), ws.Cells(lastR, C_OKINGAKU)).NumberFormatLocal = "#,##0"
    ws.Range(ws.Cells(headLastR + 1, C_NTANKA), ws.Cells(lastR, C_NKINGAKU)).NumberFormatLocal = "#,##0"
    ws.Range(ws.Cells(headLastR + 1, C_SAGAKU), ws.Cells(lastR, C_SAGAKU)).NumberFormatLocal = "#,##0;[赤]-#,##0"

    ws.Columns("A:D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 6
    ws.Columns("F:L").ColumnWidth = 11

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$" & headR & ":$" & headLastR
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
    End With
End Sub
