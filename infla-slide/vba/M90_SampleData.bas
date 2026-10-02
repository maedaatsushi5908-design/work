'==============================================================
' M90_SampleData  ―  動作確認用のサンプルCSVを作る（任意モジュール）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M90_SampleData" に。
' 実物のエスティマ「スライド用csv出力」と同じ構成
'   ・1行目  ヘッダ行
'   ・単価表部（Gコードの代価表）… X1000より前
'   ・内訳部（X1000〜Zコードの直前）… Y23系のコードで階層
'   ・諸経費部（Zコード）
' のサンプルCSVを、旧単価用と新単価用の2本作ります。
'
' 数量は ①＝変更設計、②＝当初設計 の順です。
'==============================================================
Option Explicit

Private mSb As String


Public Sub テスト用CSV作成()
    Dim baseDir As String, pOld As String, pNew As String

    baseDir = ThisWorkbook.Path
    If baseDir = "" Then
        MsgBox "先にこのブックを保存してください。", vbExclamation
        Exit Sub
    End If

    pOld = baseDir & Application.PathSeparator & "sample_旧単価.csv"
    pNew = baseDir & Application.PathSeparator & "sample_新単価.csv"

    WriteTextFile pOld, SampleCsv(False), "Shift_JIS"
    WriteTextFile pNew, SampleCsv(True), "Shift_JIS"

    If SheetExists(SH_CONFIG) Then
        SetConfigValue "旧単価CSVパス", pOld
        SetConfigValue "新単価CSVパス", pNew
        SetConfigValue "工事名", "○○地内　配水管布設工事（テスト）"
        SetConfigValue "工事場所", "神戸市○○区○○町地内"
        SetConfigValue "工期（自）", "令和7年4月1日"
        SetConfigValue "工期（至）", "令和8年3月20日"
        SetConfigValue "請負代金額", 120000000
        SetConfigValue "基準日", "令和7年10月1日"
        SetConfigValue "発注者", "神戸市"
        SetConfigValue "受注者", "株式会社○○建設"
        SetConfigValue "契約保証費", 77899
    End If

    MsgBox "サンプルCSVを作成しました。" & vbCrLf & vbCrLf & _
           pOld & vbCrLf & pNew & vbCrLf & vbCrLf & _
           "続けて［スライド計算表作成］を実行してください。", vbInformation
End Sub


Private Sub SetConfigValue(ByVal label As String, ByVal v As Variant)
    Dim ws As Worksheet
    Dim r As Long, lastR As Long

    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 1 To lastR
        If NormText(CStr(ws.Cells(r, 1).Value)) = NormText(label) Then
            ws.Cells(r, 2).Value = v
            Exit Sub
        End If
    Next r
End Sub


'--------------------------------------------------------------
' サンプル本体
'   isNew = False … 当初の単価適用日／True … 基準日の単価適用日
'--------------------------------------------------------------
Private Function SampleCsv(ByVal isNew As Boolean) As String
    mSb = ""

    ' 1行目はヘッダ行（積算システムが付ける制御行）
    mSb = "A1,8.69E+12,1,71001,80901,0,0,0,0,0,0,0,0" & vbCrLf

    '--- 単価表部（Gコードの代価表。内訳には入らない）---
    Lv "G0001", "／再掘削工", "式", "殻運搬処理　処分費含む", ""
    Det "舗装版破砕工", "㎡", "", "", 93, 93, IIf(isNew, 1420, 1350)
    Det "床掘", "", "", "", 150, 150, IIf(isNew, 336, 320)
    Lv "G0003", "／交通誘導警備員", "式", "", ""
    Det "交通誘導警備員B", "人日", "", "", 1152, 1152, IIf(isNew, 12080, 11870)
    Lv "G1005", "／仮設資材運搬", "式", "積込み取卸し含む", "往復分"
    Det "仮設材等の運搬", "t", "(鋼矢板､H形鋼等)", "往路", 2, 2, IIf(isNew, 5820, 5480)
    Lv "G1002", "／通水試験費", "式", "器具損料・諸雑費含む", ""
    Det "通水試験工", "日", "", "", 1, 1, IIf(isNew, 86400, 82300)
    Lv "G0007", "／スクラップ", "式", "撤去管・残管等", "；故銑A"
    Det "スクラップ", "ｔ", "故銑A", "", -7.9, -7.9, IIf(isNew, 23500, 9500)

    '--- 内訳部 ---
    Lv "X1000", "本工事費", "", "", ""

    Lv "Y2301", "管路(開削)", "", "", ""
    Lv "Y230101", "材料", "式", "", ""
    Lv "Y23010101", "管材料", "", "", ""
    Lv "Y2301010102", "GX形管", "", "", ""
    Det "ＧＸ形直管", "本", "継手材含む", "", 114, 114, IIf(isNew, 35690, 35400)
    Det "ＧＸ形直管", "本", "継手材含む", "", 30, 0, IIf(isNew, 35690, 35400)
    Lv "Y23010103", "ポリエチレン管", "", "", ""
    Lv "Y2301010301", "配水用ﾎﾟﾘｴﾁﾚﾝ管", "", "", ""
    Det "ＰＥ形 ＥＦ受口付直管", "本", "配水用", "", 500, 0, IIf(isNew, 8308, 8101)

    Lv "Y230102", "管布設工", "式", "", ""
    Lv "Y23010201", "管路土工", "", "", ""
    Lv "Y2301020101", "舗装版切断", "", "", ""
    Det "舗装版切断", "m", "", "", 5260, 5260, IIf(isNew, 1613, 1583)
    Lv "Y2301020110", "管路掘削", "", "", ""
    Det "床掘", "", "試掘", "", 1000, 80, IIf(isNew, 336, 320)
    Det "土砂等運搬", "m3", "現場～仮置場L=4.2km", "", 2000, 3010, IIf(isNew, 1180, 1150)
    Lv "Y2301020112", "再掘削", "", "", ""
    DetG "G0001", "／再掘削工", "式", "殻運搬処理　処分費含む", 1, 1, IIf(isNew, 452000, 410000)

    Lv "Y230104", "管工", "式", "", ""
    Lv "Y23010401", "管据付・撤去工", "", "", ""
    Lv "Y2301040101", "管据付", "", "", ""
    Det "鋳鉄管据付", "ｍ", "", "", 638.1, 638.1, IIf(isNew, 5820, 5430)

    Lv "Y230107", "付帯工", "式", "", ""
    Lv "Y23010701", "舗装撤去工", "", "", ""
    Lv "Y2301070101", "舗装版切断", "", "", ""
    Det "舗装版切断", "m", "", "", 650, 650, IIf(isNew, 1613, 1583)
    Lv "Y23010702", "道路復旧工", "", "", ""
    Lv "Y2301070201", "舗装復旧", "", "", ""
    Det "舗装復旧工", "㎡", "", "", 3120, 3120, IIf(isNew, 2240, 2080)

    Lv "Y230108", "仮設工", "式", "", ""
    Lv "Y23010803", "交通管理工", "", "", ""
    DetG "G0003", "／交通誘導警備員", "式", "", 1, 1, IIf(isNew, 13916160, 13674240)

    '--- 諸経費部 ---
    Lv "Z0001", "運搬費", "式", "", ""
    Lv "YZ000000003", "仮設材運搬費", "式", "", ""
    DetG "G1005", "／仮設資材運搬", "式", "積込み取卸し含む", 1, 1, IIf(isNew, 23280, 21920)
    Lv "Z0002", "準備費", "式", "", ""
    Lv "Z0006", "技術管理費", "式", "", ""
    DetG "G1002", "／通水試験費", "式", "器具損料・諸雑費含む", 1, 1, IIf(isNew, 86400, 82300)
    Lv "Z0013", "支給品費", "式", "", ""
    Lv "Z0040", "工期延長等に伴う増加費用", "式", "", ""
    Lv "Z0045", "設計委託費", "式", "", ""
    Lv "Z0047", "スクラップ", "式", "", ""
    DetG "G0007", "／スクラップ", "式", "撤去管・残管等", 1, 1, IIf(isNew, -185650, -75050)

    SampleCsv = mSb
End Function


' 階層行・見出し行（数量なし）
Private Sub Lv(ByVal code As String, ByVal nm As String, ByVal tani As String, _
               ByVal k1 As String, ByVal k2 As String)
    Row13 code, nm, tani, k1, k2, "", 0, 0, 0
End Sub


' 明細行（コードなし）
Private Sub Det(ByVal nm As String, ByVal tani As String, ByVal k1 As String, _
                ByVal k2 As String, ByVal qNew As Double, ByVal qOrg As Double, _
                ByVal tanka As Double)
    Row13 "", nm, tani, k1, k2, "", qNew, qOrg, tanka
End Sub


' 明細行（Gコードの代価を引く行）
Private Sub DetG(ByVal code As String, ByVal nm As String, ByVal tani As String, _
                 ByVal k1 As String, ByVal qNew As Double, ByVal qOrg As Double, _
                 ByVal tanka As Double)
    Row13 code, nm, tani, k1, "", "", qNew, qOrg, tanka
End Sub


' 13列。①＝変更設計、②＝当初設計
Private Sub Row13(ByVal code As String, ByVal nm As String, ByVal tani As String, _
                  ByVal k1 As String, ByVal k2 As String, ByVal tekiyo As String, _
                  ByVal qNew As Double, ByVal qOrg As Double, ByVal tanka As Double)
    mSb = mSb & Q(code) & "," & Q(nm) & "," & Q(tani) & "," & Q(k1) & "," & Q(k2) & "," & Q(tekiyo) & _
          "," & IIf(tanka = 0, "", Format$(tanka, "0")) & "," & IIf(tanka = 0, "", Format$(tanka, "0")) & _
          "," & IIf(qNew = 0 And qOrg = 0, "", Format$(qNew, "0.##")) & _
          "," & IIf(qNew = 0 And qOrg = 0, "", Format$(qOrg, "0.##")) & _
          "," & IIf(tanka = 0, "", Format$(qNew * tanka, "0")) & _
          "," & IIf(tanka = 0, "", Format$(qOrg * tanka, "0")) & "," & vbCrLf
End Sub


Private Function Q(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Then
        Q = """" & Replace(s, """", """""") & """"
    Else
        Q = s
    End If
End Function
