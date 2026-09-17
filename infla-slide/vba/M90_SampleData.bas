'==============================================================
' M90_SampleData  ―  動作確認用のサンプルCSVを作る（任意モジュール）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M90_SampleData" に。
' 実物のCSVが手元に無くても、これで一通りの動作確認ができます。
' 本番運用時は削除して構いません。
'==============================================================
Option Explicit


Public Sub テスト用CSV作成()
    Dim baseDir As String
    Dim pOld As String, pNew As String

    baseDir = ThisWorkbook.Path
    If baseDir = "" Then
        MsgBox "先にこのブックを保存してください。", vbExclamation
        Exit Sub
    End If

    pOld = baseDir & Application.PathSeparator & "sample_変動前.csv"
    pNew = baseDir & Application.PathSeparator & "sample_変動後.csv"

    WriteTextFile pOld, SampleCsv(False), "Shift_JIS"
    WriteTextFile pNew, SampleCsv(True), "Shift_JIS"

    ' 設定シートがあればパスと工事情報を流し込む
    If SheetExists(SH_CONFIG) Then
        SetConfigValue "変動前CSVパス", pOld
        SetConfigValue "変動後CSVパス", pNew
        SetConfigValue "工事名", "○○地内　道路改良工事（テスト）"
        SetConfigValue "工事場所", "○○県○○市○○地内"
        SetConfigValue "工期（自）", "令和7年4月1日"
        SetConfigValue "工期（至）", "令和8年3月20日"
        SetConfigValue "請負代金額", 58500000
        SetConfigValue "基準日", "令和7年10月1日"
        SetConfigValue "発注者", "○○県○○土木事務所"
        SetConfigValue "受注者", "株式会社○○建設"
    End If

    MsgBox "サンプルCSVを作成しました。" & vbCrLf & vbCrLf & _
           pOld & vbCrLf & pNew & vbCrLf & vbCrLf & _
           "続けて［インフレスライド設計書作成］を実行してください。", vbInformation
End Sub


'--------------------------------------------------------------
' 設定シートの値を書き換える（A列の項目名で検索）
'--------------------------------------------------------------
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
' サンプル明細（基準日以降の残工事を想定）
'   isNew = False … 変動前（旧労務単価）
'   isNew = True  … 変動後（新労務単価）
'--------------------------------------------------------------
Private Function SampleCsv(ByVal isNew As Boolean) As String
    Dim sb As String

    ' 新単価では労務比率の高い工種ほど単価が上がる想定
    sb = "工種,種別,細別,規格,単位,数量,単価,金額,摘要" & vbCrLf

    sb = sb & CsvRow("土工", "掘削工", "機械掘削", "土砂　オープンカット", "m3", 1250, IIf(isNew, 336, 320))
    sb = sb & CsvRow("土工", "埋戻工", "埋戻し", "流用土　人力併用", "m3", 880, IIf(isNew, 452, 410))
    sb = sb & CsvRow("土工", "残土処理工", "残土運搬", "DT10t　L=5km", "m3", 620, IIf(isNew, 1180, 1150))
    sb = sb & CsvRow("擁壁工", "場所打擁壁工", "コンクリート", "18-8-25(高炉)", "m3", 145, IIf(isNew, 22400, 21000))
    sb = sb & CsvRow("擁壁工", "場所打擁壁工", "型枠", "一般型枠", "m2", 980, IIf(isNew, 3620, 3280))
    sb = sb & CsvRow("擁壁工", "場所打擁壁工", "鉄筋", "SD345　D16", "t", 12.4, IIf(isNew, 138000, 128000))
    sb = sb & CsvRow("舗装工", "路盤工", "下層路盤", "クラッシャラン　t=200", "m2", 3100, IIf(isNew, 1420, 1350))
    sb = sb & CsvRow("舗装工", "路盤工", "上層路盤", "粒度調整砕石　t=150", "m2", 3100, IIf(isNew, 1680, 1590))
    sb = sb & CsvRow("舗装工", "表層工", "表層", "密粒度AS　t=50", "m2", 3100, IIf(isNew, 2240, 2080))
    sb = sb & CsvRow("排水工", "側溝工", "L型側溝", "300×300", "m", 460, IIf(isNew, 9800, 9100))
    sb = sb & CsvRow("排水工", "集水桝工", "集水桝", "600×600　H=1.0m", "箇所", 18, IIf(isNew, 86000, 79000))
    sb = sb & CsvRow("付属物工", "防護柵工", "ガードレール", "Gr-A-2E", "m", 320, IIf(isNew, 7400, 7050))
    sb = sb & CsvRow("付属物工", "区画線工", "区画線", "溶融式　実線15cm", "m", 2800, IIf(isNew, 410, 395))
    sb = sb & CsvRow("共通仮設費", "共通仮設費", "共通仮設費", "率計上", "式", 1, IIf(isNew, 4180000, 3900000))
    sb = sb & CsvRow("現場管理費", "現場管理費", "現場管理費", "率計上", "式", 1, IIf(isNew, 7350000, 6880000))
    sb = sb & CsvRow("一般管理費等", "一般管理費等", "一般管理費等", "率計上", "式", 1, IIf(isNew, 5120000, 4820000))

    SampleCsv = sb
End Function


Private Function CsvRow(ByVal koshu As String, ByVal shubetsu As String, ByVal saibetsu As String, _
                      ByVal kikaku As String, ByVal tani As String, _
                      ByVal suryo As Double, ByVal tanka As Double) As String
    CsvRow = koshu & "," & shubetsu & "," & saibetsu & "," & _
           """" & kikaku & """" & "," & tani & "," & _
           Format$(suryo, "0.00") & "," & Format$(tanka, "0") & "," & _
           Format$(suryo * tanka, "0") & "," & vbCrLf
End Function
