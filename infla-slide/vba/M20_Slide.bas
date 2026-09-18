'==============================================================
' M20_Slide  ―  設定シート生成／スライド計算表（明細）の生成
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M20_Slide" にしてください。
'
' このシートは「明細」だけを持ちます。
' 諸経費（直接工事費〜工事費）は M40_Keihi が別シート「経費計算シート」に作ります。
'==============================================================
Option Explicit

' ---- スライド計算表の列 ----
Public Const SC_CODE     As Long = 1    ' A コード（作業用・非表示可）
Public Const SC_HIMOKU   As Long = 2    ' B 費目
Public Const SC_KUBUN    As Long = 3    ' C 工事区分
Public Const SC_KOSHU    As Long = 4    ' D 工種
Public Const SC_SHUBETSU As Long = 5    ' E 種別
Public Const SC_SAIBETSU As Long = 6    ' F 細別
Public Const SC_NAME     As Long = 7    ' G 施工単価名称
Public Const SC_KIKAKU   As Long = 8    ' H 規格
Public Const SC_TANKA_O  As Long = 9    ' I 単価 スライド前
Public Const SC_TANKA_N  As Long = 10   ' J 単価 スライド後
Public Const SC_Q_ALL    As Long = 11   ' K 数量 全体
Public Const SC_Q_DONE   As Long = 12   ' L 数量 出来形   ★手入力
Public Const SC_Q_REST   As Long = 13   ' M 数量 残工事   =K-L
Public Const SC_TANI     As Long = 14   ' N 単位
Public Const SC_MK_SHOBU As Long = 15   ' O 処分費〇      ★手入力
Public Const SC_MK_KANZA As Long = 16   ' P 管材費〇      ★手入力
Public Const SC_MK_SHIN  As Long = 17   ' Q 新工種〇      ★手入力
Public Const SC_A_ALL    As Long = 18   ' R スライド前・全体   =K*I
Public Const SC_A_DONE   As Long = 19   ' S スライド前・出来形 =L*I
Public Const SC_A_REST   As Long = 20   ' T スライド前・残工事 =M*I
Public Const SC_B_ALL    As Long = 21   ' U スライド後・全体   =K*J
Public Const SC_B_REST   As Long = 22   ' V スライド後・残工事 =M*J
Public Const SC_TEKIYO   As Long = 23   ' W 摘要
Public Const SC_SP_O     As Long = 24   ' X 処分費単価 スライド前 ★手入力
Public Const SC_SP_N     As Long = 25   ' Y 処分費単価 スライド後 ★手入力
Public Const SC_SA_ALL   As Long = 26   ' Z  処分費額 前・全体
Public Const SC_SA_DONE  As Long = 27   ' AA 処分費額 前・出来形
Public Const SC_SA_REST  As Long = 28   ' AB 処分費額 前・残工事
Public Const SC_SB_ALL   As Long = 29   ' AC 処分費額 後・全体
Public Const SC_SB_REST  As Long = 30   ' AD 処分費額 後・残工事
Public Const SC_LAST     As Long = 30

Public Const SH_SLIDE    As String = "スライド計算表"
Public Const FIRST_ROW   As Long = 8

Public Const CLR_INPUT   As Long = 65535          ' 黄色（手入力）
Public Const CLR_HEAD    As Long = 15849925       ' 薄い青
Public Const CLR_TOTAL   As Long = 49407          ' オレンジ

' ---- M40 が参照する行範囲 ----
Public gDirectFirst As Long     ' 直接工事費部の先頭行
Public gDirectLast  As Long     ' 直接工事費部の最終行
Public gZFirst      As Long     ' 共通仮設費 積上げ部の先頭行（0＝なし）
Public gZLast       As Long     ' 　　　　　　　　　　最終行
Public gZItems      As Object   ' Dictionary 名称 → "先頭行|最終行"
Public gLastRow     As Long


'==============================================================
' 設定シート
'==============================================================
Public Sub 設定シート作成()
    Dim ws As Worksheet
    Dim r As Long

    If SheetExists(SH_CONFIG) Then
        If MsgBox("「" & SH_CONFIG & "」シートを作り直します。入力済みの内容は消えます。" & _
                  vbCrLf & "よろしいですか？", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = FreshSheet(SH_CONFIG)

    With ws.Range("A1")
        .Value = "インフレスライド計算表 作成設定"
        .Font.Size = 14
        .Font.Bold = True
    End With

    r = 3
    PutHead ws, r, "【工事情報】":                                                     r = r + 1
    PutItem ws, r, "工事名", "":                                                        r = r + 1
    PutItem ws, r, "工事場所", "":                                                      r = r + 1
    PutItem ws, r, "工期（自）", "":                                                    r = r + 1
    PutItem ws, r, "工期（至）", "":                                                    r = r + 1
    PutItem ws, r, "請負代金額", 0, "税込。様式4-2号の請負率の算出に使います":          r = r + 1
    PutItem ws, r, "基準日", "", "インフレスライドの請求日":                            r = r + 1
    PutItem ws, r, "発注者", "神戸市":                                                  r = r + 1
    PutItem ws, r, "受注者", "":                                                        r = r + 2

    PutHead ws, r, "【CSV設定】", "エスティマ 機能→CSV連携→抽出指示ファイル「スライド用csv出力」": r = r + 1
    PutItem ws, r, "旧単価CSVパス", "", "当初の単価適用日で出力したCSV。空欄なら実行時に選択":   r = r + 1
    PutItem ws, r, "新単価CSVパス", "", "基準日の単価適用日で出力したCSV。空欄なら実行時に選択": r = r + 1
    PutItem ws, r, "文字コード", "Shift_JIS", "UTF-8の場合は UTF-8 と入力":              r = r + 1
    PutItem ws, r, "区切り文字", ",":                                                   r = r + 1
    PutItem ws, r, "使用系列", "自動", "自動／当初／変更。通常は自動（①側＝最新を採用）": r = r + 2

    PutHead ws, r, "【経費計算の設定】", "経費率は経費計算シートで直接入力できます": r = r + 1
    PutItem ws, r, "契約保証費", 0, "当初設計時の額で固定":                             r = r + 1
    PutItem ws, r, "処分費控除率", 0.03, "③＝②－ROUNDDOWN((①－①'/2)×この率,0)":     r = r + 1
    PutItem ws, r, "消費税率", 0.1:                                                     r = r + 2

    PutHead ws, r, "【経費率の自動計算】", "経費計算シートの率欄の初期値。積算システムの値で上書きできます": r = r + 1
    PutItem ws, r, "共通仮設費率A", 1228.3, "率(%) = A × 対象額 ^ B":                   r = r + 1
    PutItem ws, r, "共通仮設費率B", -0.2614:                                            r = r + 1
    PutItem ws, r, "共通仮設費 地域補正", 1.2:                                          r = r + 1
    PutItem ws, r, "共通仮設費 週休補正", 1.04:                                         r = r + 1
    PutItem ws, r, "現場管理費率A", 458.2, "率(%) = A × 対象額 ^ B":                    r = r + 1
    PutItem ws, r, "現場管理費率B", -0.1508:                                            r = r + 1
    PutItem ws, r, "現場管理費 地域補正", 1.1:                                          r = r + 1
    PutItem ws, r, "現場管理費 週休補正", 1.06:                                         r = r + 1
    PutItem ws, r, "一般管理費率係数", -5.48972, "率(%) = 係数 × LOG10(対象額) + 定数": r = r + 1
    PutItem ws, r, "一般管理費率定数", 59.4977

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 58
    ws.Range("C:C").Font.Color = RGB(120, 120, 120)

    Application.ScreenUpdating = True
    ws.Activate
    ws.Range("B4").Select

    MsgBox "「" & SH_CONFIG & "」シートを作成しました。" & vbCrLf & _
           "工事情報を入力してから［スライド計算表作成］を実行してください。", vbInformation
End Sub


Private Sub PutHead(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    Optional ByVal note As String = "")
    With ws.Cells(r, 1)
        .Value = label
        .Font.Bold = True
        .Interior.Color = CLR_HEAD
    End With
    If note <> "" Then ws.Cells(r, 3).Value = note
End Sub


Private Sub PutItem(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
                    ByVal defaultValue As Variant, Optional ByVal note As String = "")
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = defaultValue
    ws.Cells(r, 2).Interior.Color = CLR_INPUT
    ws.Cells(r, 2).Borders.LineStyle = xlContinuous
    If note <> "" Then ws.Cells(r, 3).Value = note
End Sub


'==============================================================
' スライド計算表（明細）を作る
'==============================================================
Public Function BuildSlideSheet(ByVal cfg As Object, ByVal recOld As Collection, _
                                ByVal recNew As Collection) As String
    Dim ws As Worksheet
    Dim newTanka() As Double
    Dim warnText As String
    Dim i As Long, r As Long, n As Long
    Dim rec As Variant
    Dim lvl As String
    Dim lvlNum As Long
    Dim rowArr() As Long, lvlArr() As Long
    Dim cnt As Long, firstZ As Long, skippedG As Long
    Dim zCnt As Long, zRowArr() As Long, zLvlArr() As Long
    Dim curItem As String, curStart As Long

    n = recOld.Count
    ReDim newTanka(1 To n)
    warnText = MatchNewPrices(recOld, recNew, newTanka)

    Set ws = FreshSheet(SH_SLIDE)
    Set gZItems = CreateObject("Scripting.Dictionary")
    WriteHeader ws, cfg

    ReDim rowArr(1 To n)
    ReDim lvlArr(1 To n)
    cnt = 0
    r = FIRST_ROW - 1
    firstZ = 0
    gDirectFirst = FIRST_ROW

    '--- 直接工事費部（X1000配下）---
    For i = 1 To n
        rec = recOld(i)
        lvl = CStr(rec(R_LEVEL))

        If lvl = "G" Then
            skippedG = skippedG + 1
        ElseIf lvl = "Z" Or lvl = "YZ" Then
            firstZ = i
            Exit For
        Else
            r = r + 1
            lvlNum = LevelNum(lvl)
            WriteRow ws, r, rec, newTanka(i), lvlNum, lvl
            cnt = cnt + 1
            rowArr(cnt) = r
            lvlArr(cnt) = lvlNum
        End If
    Next i

    WriteSumFormulas ws, rowArr, lvlArr, cnt
    gDirectLast = r

    '--- 共通仮設費 積上げ部（Zコード配下）---
    gZFirst = 0: gZLast = 0
    If firstZ > 0 Then
        gZFirst = r + 1
        ReDim zRowArr(1 To n)
        ReDim zLvlArr(1 To n)
        zCnt = 0
        curItem = "": curStart = 0

        For i = firstZ To n
            rec = recOld(i)
            lvl = CStr(rec(R_LEVEL))
            If lvl = "G" Then GoTo NextZ

            r = r + 1
            lvlNum = LevelNum(lvl)
            WriteRow ws, r, rec, newTanka(i), lvlNum, lvl

            If lvl = "Z" Then
                If curItem <> "" Then gZItems(curItem) = curStart & "|" & (r - 1)
                curItem = CStr(rec(R_NAME))
                curStart = r
            End If

            zCnt = zCnt + 1
            zRowArr(zCnt) = r
            zLvlArr(zCnt) = lvlNum
NextZ:
        Next i

        If curItem <> "" Then gZItems(curItem) = curStart & "|" & r
        WriteSumFormulas ws, zRowArr, zLvlArr, zCnt
        gZLast = r
    End If

    gLastRow = r
    FinishSheet ws, r

    If skippedG > 0 Then
        warnText = warnText & IIf(warnText = "", "", vbCrLf) & _
                   "Gコード行 " & skippedG & " 行は計算表に出していません。"
    End If

    BuildSlideSheet = warnText
End Function


'--------------------------------------------------------------
' 旧CSVの各行に対応する新単価を求める
'--------------------------------------------------------------
Private Function MatchNewPrices(ByVal recOld As Collection, ByVal recNew As Collection, _
                                ByRef newTanka() As Double) As String
    Dim i As Long, miss As Long
    Dim positional As Boolean
    Dim d As Object
    Dim k As String
    Dim rec As Variant

    positional = (recOld.Count = recNew.Count)
    If positional Then
        For i = 1 To recOld.Count
            If RecKey(recOld(i)) <> RecKey(recNew(i)) Then
                positional = False
                Exit For
            End If
        Next i
    End If

    If positional Then
        For i = 1 To recOld.Count
            rec = recNew(i)
            newTanka(i) = CDbl(rec(R_TANKA))
        Next i
        Exit Function
    End If

    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To recNew.Count
        rec = recNew(i)
        k = RecKey(rec)
        If Not d.Exists(k) Then d.Add k, CDbl(rec(R_TANKA))
    Next i

    For i = 1 To recOld.Count
        rec = recOld(i)
        k = RecKey(rec)
        If d.Exists(k) Then
            newTanka(i) = d(k)
        Else
            newTanka(i) = CDbl(rec(R_TANKA))
            If CStr(rec(R_LEVEL)) = "" Then miss = miss + 1
        End If
    Next i

    MatchNewPrices = "旧CSVと新CSVで行がずれていたため、名称・規格・単位で突合しました。" & vbCrLf & _
                     "新単価が見つからず旧単価を据え置いた明細：" & miss & " 行"
End Function


Public Function LevelNum(ByVal lvl As String) As Long
    Select Case lvl
        Case "費目": LevelNum = 0
        Case "L1":   LevelNum = 1
        Case "L2":   LevelNum = 2
        Case "L3":   LevelNum = 3
        Case "L4":   LevelNum = 4
        Case "Z":    LevelNum = 1
        Case "YZ":   LevelNum = 2
        Case Else:   LevelNum = 9
    End Select
End Function


'==============================================================
' 見出し
'==============================================================
Private Sub WriteHeader(ByVal ws As Worksheet, ByVal cfg As Object)
    ws.Range("B1").Value = CfgVal(cfg, "工事名", "")
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Bold = True
    ws.Range("G1").Value = "基準日：" & CStr(CfgVal(cfg, "基準日", ""))

    ws.Cells(3, SC_CODE).Value = "コード"
    ws.Cells(3, SC_HIMOKU).Value = "費目"
    ws.Cells(3, SC_KUBUN).Value = "工事区分"
    ws.Cells(4, SC_KOSHU).Value = "工種"
    ws.Cells(5, SC_SHUBETSU).Value = "種別"
    ws.Cells(6, SC_SAIBETSU).Value = "細別"
    ws.Cells(3, SC_KIKAKU).Value = "規格"
    ws.Cells(3, SC_TANKA_O).Value = "単価"
    ws.Cells(5, SC_TANKA_O).Value = "スライド前"
    ws.Cells(5, SC_TANKA_N).Value = "スライド後"
    ws.Cells(3, SC_Q_ALL).Value = "数量"
    ws.Cells(5, SC_Q_ALL).Value = "全体"
    ws.Cells(5, SC_Q_DONE).Value = "出来形"
    ws.Cells(5, SC_Q_REST).Value = "残工事"
    ws.Cells(3, SC_TANI).Value = "単位"
    ws.Cells(3, SC_MK_SHOBU).Value = "区分（〇を入力）"
    ws.Cells(5, SC_MK_SHOBU).Value = "処分費"
    ws.Cells(5, SC_MK_KANZA).Value = "管材費"
    ws.Cells(5, SC_MK_SHIN).Value = "新工種"
    ws.Cells(3, SC_A_ALL).Value = "スライド前（旧単価）"
    ws.Cells(5, SC_A_ALL).Value = "全体"
    ws.Cells(5, SC_A_DONE).Value = "出来形"
    ws.Cells(5, SC_A_REST).Value = "残工事"
    ws.Cells(3, SC_B_ALL).Value = "スライド後（新単価）"
    ws.Cells(5, SC_B_ALL).Value = "全体"
    ws.Cells(5, SC_B_REST).Value = "残工事"
    ws.Cells(3, SC_TEKIYO).Value = "摘要"
    ws.Cells(3, SC_SP_O).Value = "処分費単価（手入力）"
    ws.Cells(5, SC_SP_O).Value = "スライド前"
    ws.Cells(5, SC_SP_N).Value = "スライド後"
    ws.Cells(3, SC_SA_ALL).Value = "処分費額"
    ws.Cells(5, SC_SA_ALL).Value = "前・全体"
    ws.Cells(5, SC_SA_DONE).Value = "前・出来形"
    ws.Cells(5, SC_SA_REST).Value = "前・残工事"
    ws.Cells(5, SC_SB_ALL).Value = "後・全体"
    ws.Cells(5, SC_SB_REST).Value = "後・残工事"

    With ws.Range(ws.Cells(3, 1), ws.Cells(6, SC_LAST))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.Color = CLR_HEAD
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range(ws.Cells(3, SC_MK_SHOBU), ws.Cells(6, SC_MK_SHIN)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(3, SC_SP_O), ws.Cells(6, SC_SP_N)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(6, 1), ws.Cells(6, SC_LAST)).Borders(xlEdgeBottom).LineStyle = xlDouble
End Sub


'==============================================================
' 1行を書き出す
'==============================================================
Private Sub WriteRow(ByVal ws As Worksheet, ByVal r As Long, ByVal rec As Variant, _
                     ByVal tankaNew As Double, ByVal lvlNum As Long, ByVal lvl As String)
    ws.Cells(r, SC_CODE).Value = rec(R_CODE)
    ws.Cells(r, SC_TEKIYO).Value = rec(R_TEKIYO)

    If lvl = "Z" Then
        ws.Cells(r, SC_KUBUN).Value = rec(R_NAME)
    ElseIf lvl = "YZ" Then
        ws.Cells(r, SC_KOSHU).Value = "②"
        ws.Cells(r, SC_SHUBETSU).Value = rec(R_NAME)
    Else
        Select Case lvlNum
            Case 0
                ws.Cells(r, SC_HIMOKU).Value = rec(R_NAME)
            Case 1
                ws.Cells(r, SC_KUBUN).Value = "①"
                ws.Cells(r, SC_KOSHU).Value = rec(R_NAME)
            Case 2
                ws.Cells(r, SC_KOSHU).Value = "②"
                ws.Cells(r, SC_SHUBETSU).Value = rec(R_NAME)
            Case 3
                ws.Cells(r, SC_SHUBETSU).Value = "③"
                ws.Cells(r, SC_SAIBETSU).Value = rec(R_NAME)
            Case 4
                ws.Cells(r, SC_SAIBETSU).Value = "④"
                ws.Cells(r, SC_NAME).Value = rec(R_NAME)
            Case Else
                ws.Cells(r, SC_NAME).Value = rec(R_NAME)
                ws.Cells(r, SC_KIKAKU).Value = rec(R_KIKAKU)
        End Select
    End If

    If lvlNum = 9 Then
        ws.Cells(r, SC_TANKA_O).Value = rec(R_TANKA)
        ws.Cells(r, SC_TANKA_N).Value = tankaNew
        ws.Cells(r, SC_Q_ALL).Value = rec(R_SURYO)
        ws.Cells(r, SC_TANI).Value = rec(R_TANI)
        WriteDetailFormulas ws, r
    Else
        ws.Cells(r, SC_TANI).Value = "式"
    End If
End Sub


'--------------------------------------------------------------
' 明細行の数式
'--------------------------------------------------------------
Public Sub WriteDetailFormulas(ByVal ws As Worksheet, ByVal r As Long)
    Dim aI As String, aJ As String, aK As String, aL As String, aM As String
    Dim aO As String, aX As String, aY As String
    Dim spO As String, spN As String

    aI = ws.Cells(r, SC_TANKA_O).Address(False, False)
    aJ = ws.Cells(r, SC_TANKA_N).Address(False, False)
    aK = ws.Cells(r, SC_Q_ALL).Address(False, False)
    aL = ws.Cells(r, SC_Q_DONE).Address(False, False)
    aM = ws.Cells(r, SC_Q_REST).Address(False, False)
    aO = ws.Cells(r, SC_MK_SHOBU).Address(False, False)
    aX = ws.Cells(r, SC_SP_O).Address(False, False)
    aY = ws.Cells(r, SC_SP_N).Address(False, False)

    ' 処分費単価：未入力ならその行の設計単価を使う
    spO = "IF(" & aX & "=""""," & aI & "," & aX & ")"
    spN = "IF(" & aY & "="""",IF(" & aX & "=""""," & aJ & "," & aX & ")," & aY & ")"

    ws.Cells(r, SC_Q_REST).Formula = "=" & aK & "-" & aL
    ws.Cells(r, SC_A_ALL).Formula = "=" & aK & "*" & aI
    ws.Cells(r, SC_A_DONE).Formula = "=" & aL & "*" & aI
    ws.Cells(r, SC_A_REST).Formula = "=" & aM & "*" & aI
    ws.Cells(r, SC_B_ALL).Formula = "=" & aK & "*" & aJ
    ws.Cells(r, SC_B_REST).Formula = "=" & aM & "*" & aJ

    ws.Cells(r, SC_SA_ALL).Formula = "=IF(" & aO & "=""〇""," & aK & "*" & spO & ",0)"
    ws.Cells(r, SC_SA_DONE).Formula = "=IF(" & aO & "=""〇""," & aL & "*" & spO & ",0)"
    ws.Cells(r, SC_SA_REST).Formula = "=IF(" & aO & "=""〇""," & aM & "*" & spO & ",0)"
    ws.Cells(r, SC_SB_ALL).Formula = "=IF(" & aO & "=""〇""," & aK & "*" & spN & ",0)"
    ws.Cells(r, SC_SB_REST).Formula = "=IF(" & aO & "=""〇""," & aM & "*" & spN & ",0)"
End Sub


'==============================================================
' 階層行の集計式（金額5列のみ。処分費額は明細行だけが持つ）
'==============================================================
Public Sub WriteSumFormulas(ByVal ws As Worksheet, ByRef rowArr() As Long, _
                            ByRef lvlArr() As Long, ByVal cnt As Long)
    Dim i As Long, j As Long, k As Long
    Dim parentIdx() As Long
    Dim stackLvl() As Long, stackIdx() As Long, sp As Long
    Dim cols As Variant, c As Variant
    Dim addrs As String
    Dim childRows() As Long, nc As Long
    Dim contiguous As Boolean

    If cnt = 0 Then Exit Sub

    ReDim parentIdx(1 To cnt)
    ReDim stackLvl(0 To cnt)
    ReDim stackIdx(0 To cnt)
    sp = 0

    For i = 1 To cnt
        Do While sp > 0
            If stackLvl(sp) >= lvlArr(i) Then sp = sp - 1 Else Exit Do
        Loop
        If sp > 0 Then parentIdx(i) = stackIdx(sp) Else parentIdx(i) = 0
        If lvlArr(i) < 9 Then
            sp = sp + 1
            stackLvl(sp) = lvlArr(i)
            stackIdx(sp) = i
        End If
    Next i

    cols = Array(SC_A_ALL, SC_A_DONE, SC_A_REST, SC_B_ALL, SC_B_REST)

    For i = 1 To cnt
        If lvlArr(i) < 9 Then
            ReDim childRows(1 To cnt)
            nc = 0
            For j = i + 1 To cnt
                If parentIdx(j) = i Then
                    nc = nc + 1
                    childRows(nc) = rowArr(j)
                ElseIf lvlArr(j) <= lvlArr(i) Then
                    Exit For
                End If
            Next j

            If nc > 0 Then
                contiguous = True
                For k = 2 To nc
                    If childRows(k) <> childRows(k - 1) + 1 Then contiguous = False: Exit For
                Next k

                For Each c In cols
                    If contiguous Then
                        addrs = ws.Cells(childRows(1), c).Address(False, False) & ":" & _
                                ws.Cells(childRows(nc), c).Address(False, False)
                    Else
                        addrs = ""
                        For k = 1 To nc
                            If addrs <> "" Then addrs = addrs & ","
                            addrs = addrs & ws.Cells(childRows(k), c).Address(False, False)
                        Next k
                    End If
                    ws.Cells(rowArr(i), c).Formula = "=SUM(" & addrs & ")"
                Next c
            End If
        End If
    Next i
End Sub


'==============================================================
' 体裁・印刷設定
'==============================================================
Public Sub FinishSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long

    If lastRow < FIRST_ROW Then lastRow = FIRST_ROW

    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, SC_LAST)).Borders.LineStyle = xlContinuous

    ws.Range(ws.Cells(FIRST_ROW, SC_TANKA_O), ws.Cells(lastRow, SC_TANKA_N)).NumberFormatLocal = "#,##0"
    ws.Range(ws.Cells(FIRST_ROW, SC_Q_ALL), ws.Cells(lastRow, SC_Q_REST)).NumberFormatLocal = "#,##0.00"
    ws.Range(ws.Cells(FIRST_ROW, SC_A_ALL), ws.Cells(lastRow, SC_B_REST)).NumberFormatLocal = "#,##0"
    ws.Range(ws.Cells(FIRST_ROW, SC_SP_O), ws.Cells(lastRow, SC_SB_REST)).NumberFormatLocal = "#,##0"

    ws.Range(ws.Cells(FIRST_ROW, SC_Q_DONE), ws.Cells(lastRow, SC_Q_DONE)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_MK_SHOBU), ws.Cells(lastRow, SC_MK_SHIN)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_SP_O), ws.Cells(lastRow, SC_SP_N)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(FIRST_ROW, SC_MK_SHOBU), ws.Cells(lastRow, SC_MK_SHIN)).HorizontalAlignment = xlCenter

    For r = FIRST_ROW To lastRow
        If ws.Cells(r, SC_HIMOKU).Value <> "" Then
            With ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO))
                .Interior.Color = RGB(47, 117, 181)
                .Font.Color = vbWhite
            End With
        ElseIf ws.Cells(r, SC_KUBUN).Value <> "" Then
            ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO)).Interior.Color = RGB(155, 194, 230)
        ElseIf ws.Cells(r, SC_KOSHU).Value <> "" Then
            ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO)).Interior.Color = RGB(189, 215, 238)
        ElseIf ws.Cells(r, SC_SHUBETSU).Value <> "" Then
            ws.Range(ws.Cells(r, SC_HIMOKU), ws.Cells(r, SC_TEKIYO)).Interior.Color = RGB(221, 235, 247)
        End If
        ' 手入力欄の黄色は階層行の塗りより優先する
        If ws.Cells(r, SC_TANKA_O).Value = "" Then
            ws.Range(ws.Cells(r, SC_MK_SHOBU), ws.Cells(r, SC_MK_SHIN)).Interior.ColorIndex = xlColorIndexNone
        End If
    Next r

    ws.Columns(SC_CODE).ColumnWidth = 9
    ws.Columns(SC_CODE).Font.Color = RGB(150, 150, 150)
    ws.Columns(SC_HIMOKU).ColumnWidth = 10
    ws.Range(ws.Columns(SC_KUBUN), ws.Columns(SC_SAIBETSU)).ColumnWidth = 4
    ws.Columns(SC_NAME).ColumnWidth = 22
    ws.Columns(SC_KIKAKU).ColumnWidth = 22
    ws.Range(ws.Columns(SC_TANKA_O), ws.Columns(SC_Q_REST)).ColumnWidth = 10
    ws.Columns(SC_TANI).ColumnWidth = 6
    ws.Range(ws.Columns(SC_MK_SHOBU), ws.Columns(SC_MK_SHIN)).ColumnWidth = 6
    ws.Range(ws.Columns(SC_A_ALL), ws.Columns(SC_B_REST)).ColumnWidth = 12
    ws.Columns(SC_TEKIYO).ColumnWidth = 24
    ws.Range(ws.Columns(SC_SP_O), ws.Columns(SC_SB_REST)).ColumnWidth = 11

    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$3:$6"
        .PrintArea = ws.Range(ws.Cells(1, SC_HIMOKU), ws.Cells(lastRow, SC_TEKIYO)).Address
        .CenterHorizontally = True
        .CenterFooter = "&P / &N"
    End With

    ws.Activate
    ws.Cells(FIRST_ROW, 1).Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub
