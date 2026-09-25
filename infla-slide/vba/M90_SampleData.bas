'==============================================================
' M90_SampleData  ―  動作確認用のサンプルCSVを作る（任意モジュール）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M90_SampleData" に。
' エスティマ「スライド用csv出力」と同じ形（12列・見出し行なし）の
' サンプルCSVを、旧単価用と新単価用の2本作ります。
' 実物のCSVが手元に無くても一通りの動作確認ができます。
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
        SetConfigValue "工事名", "○○地内　道路改良工事（テスト）"
        SetConfigValue "工事場所", "神戸市○○区○○町地内"
        SetConfigValue "工期（自）", "令和7年4月1日"
        SetConfigValue "工期（至）", "令和8年3月20日"
        SetConfigValue "請負代金額", 120000000
        SetConfigValue "基準日", "令和7年10月1日"
        SetConfigValue "発注者", "神戸市"
        SetConfigValue "受注者", "株式会社○○建設"
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
' サンプル（エスティマ「スライド用csv出力」形式・12列・見出しなし）
'   isNew = False … 当初の単価適用日
'   isNew = True  … 基準日の単価適用日（労務比率の高い工種ほど上昇）
'--------------------------------------------------------------
Private Function SampleCsv(ByVal isNew As Boolean) As String
    mSb = ""

    Head "G0100", "○○地内　道路改良工事"
    Head "X1000", "本工事費"

    Head "Y10001", "道路改良"
    Head "Y20001", "土工"
    Head "Y30001", "掘削工"
    Head "Y40001", "機械掘削"
    Det "床掘り", "m3", "土砂", "オープンカット", "", 1250, IIf(isNew, 336, 320)
    Head "Y40002", "埋戻工"
    Det "埋戻し", "m3", "流用土", "人力併用", "", 880, IIf(isNew, 452, 410)
    Det "土材料", "m3", "再生クラッシャラン0~40", "", "裏込め部", 40, IIf(isNew, 1800, 1800)
    Head "Y40003", "残土処理工"
    Det "土砂等運搬", "m3", "DT10t", "L=5km", "片道9.5km", 620, IIf(isNew, 1613, 1583)
    Det "残土等処分", "m3", "布施畑環境センター", "", "", 620, IIf(isNew, 3272, 3272)

    Head "Y30002", "擁壁工"
    Head "Y40011", "場所打擁壁工"
    Det "コンクリート", "m3", "24-12-25(20)(高炉)", "", "躯体", 145, IIf(isNew, 35690, 35400)
    Det "型枠", "m2", "一般型枠", "", "", 980, IIf(isNew, 8308, 8101)
    Det "鉄筋工", "t", "SD345", "D16～D25", "", 12.4, IIf(isNew, 135500, 117500)

    Head "Y20002", "舗装工"
    Head "Y30011", "路盤工"
    Head "Y40021", "下層路盤"
    Det "下層路盤", "m2", "クラッシャラン", "t=200", "", 3100, IIf(isNew, 1420, 1350)
    Head "Y40022", "上層路盤"
    Det "上層路盤", "m2", "粒度調整砕石", "t=150", "", 3100, IIf(isNew, 1680, 1590)
    Head "Y30012", "表層工"
    Head "Y40031", "表層"
    Det "表層", "m2", "密粒度AS", "t=50", "", 3100, IIf(isNew, 2240, 2080)

    ' 直接工事費の中の仮設工（YZコードを使う。ここで諸経費部に切り替わってはいけない）
    Head "Y10003", "仮設工"
    Head "YZ1001", "土留工"
    Det "鋼矢板打込・引抜", "m2", "Ⅲ型", "L=6.0m", "", 480, IIf(isNew, 5820, 5480)
    Det "鋼矢板賃料", "m2月", "Ⅲ型", "", "", 2880, IIf(isNew, 268, 268)

    Head "Y10002", "交通管理"
    Head "Y20011", "交通管理工"
    Head "Y30021", "交通誘導警備員"
    Det "交通誘導警備員B", "人日", "", "", "", 352, IIf(isNew, 12080, 11870)

    ' 諸経費（共通仮設費の積上げ分）
    Head "Z0001", "運搬費"
    Head "YZ0001", "重機分解組立輸送"
    Det "重機分解組立輸送", "回", "分解組立+輸送(往復)", "", "", 1, IIf(isNew, 831400, 828600)
    Head "Z0010", "準備費"
    Head "YZ0011", "木根等処分費"
    Det "高木伐採・根株撤去", "本", "幹周90cm～110cm", "", "集材含む", 2, IIf(isNew, 77750, 75485)
    Det "生木処分費", "t", "枝葉", "", "", 5.09, IIf(isNew, 16000, 16000)
    Head "Z0020", "技術管理費"
    Head "YZ0021", "土質試験費"
    Det "土の一軸圧縮試験", "試料", "", "", "2供試体/試料", 3, IIf(isNew, 10400, 10400)

    ' Z0040以降は積算システム側の計算行。共通仮設費の積上分に混ぜてはいけない
    Head "Z0040", "共通仮設費率分"
    Det "共通仮設費率分", "式", "率計上", "", "", 1, IIf(isNew, 3620000, 3500000)
    Head "Z0045", "一般管理費等"
    Det "一般管理費等", "式", "率計上", "", "", 1, IIf(isNew, 5120000, 4820000)

    ' スクラップ（数量がマイナス。㉒として工事価格の後に足す）
    Head "Z0047", "スクラップ"
    Head "YZ0048", "ｽｸﾗｯﾌﾟ"
    Det "ｽｸﾗｯﾌﾟ", "t", "", "", "", -1.93, IIf(isNew, 23500, 9500)

    SampleCsv = mSb
End Function


' 階層行（コードあり・単価/数量なし）
Private Sub Head(ByVal code As String, ByVal nm As String)
    Row12 code, nm, "式", "", "", "", 0, 0
End Sub


' 明細行（コードなし）
Private Sub Det(ByVal nm As String, ByVal tani As String, ByVal k1 As String, _
                ByVal k2 As String, ByVal tekiyo As String, _
                ByVal suryo As Double, ByVal tanka As Double)
    Row12 "", nm, tani, k1, k2, tekiyo, suryo, tanka
End Sub


Private Sub Row12(ByVal code As String, ByVal nm As String, ByVal tani As String, _
                  ByVal k1 As String, ByVal k2 As String, ByVal tekiyo As String, _
                  ByVal suryo As Double, ByVal tanka As Double)
    mSb = mSb & Q(code) & "," & Q(nm) & "," & Q(tani) & "," & Q(k1) & "," & Q(k2) & "," & Q(tekiyo) & _
          "," & Format$(tanka, "0") & ",0," & _
          Format$(suryo, "0.00") & ",0," & _
          Format$(suryo * tanka, "0") & ",0" & vbCrLf
End Sub


Private Function Q(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Then
        Q = """" & Replace(s, """", """""") & """"
    Else
        Q = s
    End If
End Function
