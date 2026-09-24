'==============================================================
' M40_Chosho  ―  スライド調書（様式4-2号）
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M40_Chosho" にしてください。
'
' 金額はすべて「スライド計算表」の経費計算部からリンクで引いています。
' スライド計算表の出来形や〇を直すと、この調書も自動で追従します。
'==============================================================
Option Explicit

Public Const SH_CHOSHO As String = "スライド調書（様式4-2号）"


Public Sub BuildChosho(ByVal cfg As Object)
    Dim ws As Worksheet, sl As Worksheet
    Dim contract As Double
    Dim fKoujihi As String, fKakaku As String, fDekidaka As String
    Dim fZanNew As String, fZei As String, fShinNuki As String, fShinNukiD As String

    Set sl = ThisWorkbook.Worksheets(SH_SLIDE)
    Set ws = FreshSheet(SH_CHOSHO)
    contract = CDbl(CfgVal(cfg, "請負代金額", 0))

    ' スライド計算表へのリンク（S=前全体 T=前出来形 V=後全体 W=後残工事 X=新抜き全体 Y=新抜き出来形）
    fKoujihi = SlideRef(sl, gRowKoujihi, SC_A_ALL)
    fKakaku = SlideRef(sl, gRowKakaku2, SC_A_ALL)
    fDekidaka = SlideRef(sl, gRowKakaku2, SC_A_DONE)
    fZanNew = SlideRef(sl, gRowKakaku2, SC_B_REST)
    fZei = SlideRef(sl, gRowZei, SC_A_ALL)
    fShinNuki = SlideRef(sl, gRowKakaku2, SC_N_ALL)
    fShinNukiD = SlideRef(sl, gRowKakaku2, SC_N_DONE)

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
    ws.Range("C8").Formula = "=" & fKoujihi
    ws.Range("H8").Value = "請負率(%)"
    If contract > 0 Then
        ws.Range("I8").Formula = "=" & Format$(contract, "0.##########") & "/C8*100"
    Else
        ws.Range("I8").Value = 100
    End If
    ws.Range("I8").Interior.Color = CLR_INPUT

    ws.Range("C9").Value = "②": ws.Range("D9").Value = "⑤"
    ws.Range("E9").Value = "⑦=②-⑤": ws.Range("F9").Value = "⑧"
    ws.Range("B10").Value = "　工事価格（税抜）"
    ws.Range("C10").Formula = "=" & fKakaku
    ws.Range("D10").Formula = "=" & fDekidaka
    ws.Range("E10").Formula = "=C10-D10"
    ws.Range("F10").Formula = "=" & fZanNew

    ws.Range("B12").Value = "　消費税相当額"
    ws.Range("C12").Formula = "=" & fZei

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
    ws.Range("C20").Formula = "=ROUNDDOWN(" & fShinNuki & "*C14/C8,0)"
    ws.Range("D20").Formula = "=ROUNDDOWN(" & fShinNukiD & "*C14/C8,0)"
    ws.Range("E20").Formula = "=C20-D20"

    ws.Range("B22").Value = "※ P1''（新工種を抜いた請負ベースの残工事）が受注者負担1%の母数です"
    ws.Range("B23").Value = "※ 金額はすべて「スライド計算表」の経費計算部からリンクしています"
    ws.Range("B22:B23").Font.Color = RGB(120, 120, 120)

    '--- 体裁 ---
    ws.Range("C5:C6").Merge
    ws.Range("D5:D6").Merge
    ws.Range("E5:F5").Merge
    ws.Range("G5:G6").Merge
    With ws.Range("C5:G6")
        .HorizontalAlignment = xlCenter
        .Interior.Color = CLR_HEAD
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    ws.Range("B8:G20").Borders.LineStyle = xlContinuous
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


' スライド計算表の1セルへの参照文字列を作る
Private Function SlideRef(ByVal sl As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If r = 0 Then
        SlideRef = "0"
    Else
        SlideRef = "'" & SH_SLIDE & "'!" & sl.Cells(r, c).Address(True, True)
    End If
End Function
