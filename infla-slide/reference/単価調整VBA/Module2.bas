Attribute VB_Name = "Module2"
Sub ☆単価調整Excel作成()

Call ①様式整理レベル表記
Call ②単価調整からL1計算まで
Call ③Zコード計算
Call ④諸経費計算

End Sub


Sub ①様式整理レベル表記()


'
 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ

'Gコード前の空白挿入
Dim r As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim rmax As Long
Dim cmax As Long
Dim rx1000 As Long
rmax = Cells(Rows.Count, 1).End(xlUp).Row

For r = rmax To 1 Step -1
If Range("A" & r).Value Like "*X1000*" Then
Rows(r).Insert
End If
Next r

rx1000 = Columns(1).Find("X1000").Row

For r = rx1000 To 1 Step -1
If Range("A" & r).Value Like "G*" Then
Rows(r).Insert
End If
Next r

' 様式整理
'
    Columns("A:B").Insert Shift:=xlToRight
    Rows("1:2").Insert Shift:=xlDown
    Cells.Replace what:=" ", Replacement:="", LookAt:=xlPart
    
    rx1000 = Columns(3).Find("X1000").Row
    
    If Range("N" & rx1000 + 1) = 0 Then
    Range("C2") = "コード"
    Range("D2") = "施工単価名称"
    Range("E2") = "単価"
    Range("F2") = "規格1"
    Range("G2") = "規格2"
    Range("H2") = "摘要"
    Range("I2") = "設計単価"
    Range("J2") = "設計単価"
    Range("K2") = "数量"
    Range("L2") = "数量"
    Range("M2") = "設計金額"
    Range("N2") = "設計金額"
    Range("I1") = "当初"
    Range("J1") = "変更"
    Range("K1") = "当初"
    Range("L1") = "変更"
    Range("M1") = "当初"
    Range("N1") = "変更"
    Range("O1") = "当初"
    Range("O2") = "調整単価"
    Range("P1") = "変更"
    Range("P2") = "調整単価"
    Range("Q1") = "当初"
    Range("Q2") = "調整金額"
    Range("R1") = "変更"
    Range("R2") = "調整金額"
    Range("T1") = "落札率↓"
    Range("T2") = "1"
    Range("U1") = "何桁目丸め↓"
    Range("U2") = "1"
    Range("B2") = "レベル"
    Range("O3") = ""
    Else
    Range("C2") = "コード"
    Range("D2") = "施工単価名称"
    Range("E2") = "単価"
    Range("F2") = "規格1"
    Range("G2") = "規格2"
    Range("H2") = "摘要"
    Range("I2") = "設計単価"
    Range("J2") = "設計単価"
    Range("K2") = "数量"
    Range("L2") = "数量"
    Range("M2") = "設計金額"
    Range("N2") = "設計金額"
    Range("I1") = "変更"
    Range("J1") = "当初"
    Range("K1") = "変更"
    Range("L1") = "当初"
    Range("M1") = "変更"
    Range("N1") = "当初"
    Range("O1") = "変更"
    Range("O2") = "調整単価"
    Range("P1") = "当初"
    Range("P2") = "調整単価"
    Range("Q1") = "変更"
    Range("Q2") = "調整金額"
    Range("R1") = "当初"
    Range("R2") = "調整金額"
    Range("T1") = "落札率↓"
    Range("T2") = "1"
    Range("U1") = "何桁目丸め↓"
    Range("U2") = "1"
    Range("B2") = "レベル"
    Range("O3") = ""
    End If
    

'レベル表記
    rmax = Cells(Rows.Count, 3).End(xlUp).Row

    For r = 1 To rmax
    If Range("C" & r).Value Like "*Y23*" Then
    If Len(Range("C" & r)) = 5 Then
    Range("B" & r).Value = "L1"
    End If
    If Len(Range("C" & r)) = 7 Then
    Range("B" & r).Value = "L2"
    End If
    If Len(Range("C" & r)) = 9 Then
    Range("B" & r).Value = "L3"
    End If
    If Len(Range("C" & r)) = 11 Then
    Range("B" & r).Value = "L4"
    End If
    End If
    If Range("C" & r).Value Like "*Y10*" Then
    Range("B" & r).Value = "L1"
    End If
    If Range("C" & r).Value Like "*Y18*" Then
    Range("B" & r).Value = "工L2"
    End If
    If Range("C" & r).Value Like "*Y20*" Then
    Range("B" & r).Value = "L2"
    End If
    If Range("C" & r).Value Like "*Y30*" Then
    Range("B" & r).Value = "L3"
    End If
    If Range("C" & r).Value Like "*Y40*" Then
    Range("B" & r).Value = "L4"
    End If
    Next
    

    

    
'工事原価計、一般管理費挿入
    rmax = Cells(Rows.Count, 3).End(xlUp).Row
    For r = rmax To 1 Step -1
    If Range("C" & r).Value Like "Z0045*" Then
    Rows(r).Insert
    Rows(r + 1).Insert
    Range("D" & r) = "工事原価計"
    Range("D" & r + 1) = "一般管理費等"
    Range("E" & r) = "式"
    Range("E" & r + 1) = "式"
    
    End If
    Next r
    
'共通仮設費～現場管理費挿入
    For r = rmax To 1 Step -1
    If Range("C" & r).Value Like "*Z0040*" Then
    Rows(r).Insert
    Rows(r).Insert
    Rows(r).Insert
    Rows(r).Insert
    Range("D" & r) = "共通仮設費率分"
    Range("D" & r + 1) = "共通仮設費計"
    Range("D" & r + 2) = "純工事費"
    Range("D" & r + 3) = "現場管理費"
    Range("E" & r) = "式"
    Range("E" & r + 1) = "式"
    Range("E" & r + 2) = "式"
    Range("E" & r + 3) = "式"
    End If
    Next r
    
    
'直接工事費挿入
    For r = rmax To 1 Step -1
    If Range("C" & r).Value Like "Z0001*" Then
    Rows(r).Insert
    Range("D" & r) = "直接工事費計"
    Range("E" & r) = "式"
    sr = r
    End If
    Next r
    
'工事価格～工事費記入
    rmax = Cells(Rows.Count, 3).End(xlUp).Row
    
    Range(Cells(rmax + 1, 2), Cells(rmax + 3, 18)).Borders.LineStyle = xlContinuous
    Range("D" & rmax + 1) = "工事価格"
    Range("D" & rmax + 2) = "消費税相当額"
    Range("D" & rmax + 3) = "工事費"
    Range("E" & rmax + 1) = "式"
    Range("E" & rmax + 2) = "式"
    Range("E" & rmax + 3) = "式"
    
'罫線入力
    rmax = Cells(Rows.Count, 3).End(xlUp).Row
    cmax = Cells(1, Columns.Count).End(xlToLeft).Column
    rx1000 = Columns(3).Find("X1000").Row
    
    Range(Cells(1, 2), Cells(rmax, 18)).Borders.LineStyle = xlContinuous
    
    Range("T1:U2").Borders.LineStyle = xlContinuous
    
    Range("B2:R2").Borders(xlEdgeBottom).LineStyle = xlDouble
    
    Range(Cells(rx1000, 2), Cells(rx1000, 18)).Borders(xlEdgeTop).LineStyle = xlDouble
    
    
'Gコード背景色変更
    rx1000 = Columns(3).Find("X1000").Row
     For i = 1 To rx1000
         If Cells(i, "C").Value Like "G*" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(221, 235, 247)
         Next
    
rmax = Cells(Rows.Count, 4).End(xlUp).Row

'レベル毎に背景色変更
     For i = 4 To rmax
     If Cells(i, "B") = "L1" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(47, 117, 181)
     If Cells(i, "B") = "L2" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(155, 194, 230)
     If Cells(i, "B") = "L3" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(189, 215, 238)
     If Cells(i, "B") = "L4" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(221, 235, 247)
     Next

'Zコード辺り背景色変更
    For i = 4 To rmax
     If Cells(i, "C").Value Like "Z0*" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(189, 215, 238)
     If Cells(i, "C").Value Like "YZ*" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(221, 235, 247)
     Next
     
'工事費など背景色変更
    For i = 4 To rmax
    If Cells(i, "D").Value = "直接工事費計" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(255, 192, 0)
    If Cells(i, "D").Value = "純工事費" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(255, 192, 0)
    If Cells(i, "D").Value = "工事原価計" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(255, 192, 0)
    If Cells(i, "D").Value = "工事価格" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(255, 192, 0)
    If Cells(i, "D").Value = "消費税相当額" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(255, 192, 0)
    If Cells(i, "D").Value = "工事費" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(255, 192, 0)
    Next
    
'様式　微調整
    Range("T2:U2").Interior.Color = RGB(255, 255, 0)
    Range("S3").Value = ""
    
 
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開


End Sub



Sub ②単価調整からL1計算まで()


 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ

Dim r As Long
Dim r1 As Long
Dim r2 As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim rmax As Long
Dim cmax As Long
Dim rx1000 As Long
rmax = Cells(Rows.Count, 3).End(xlUp).Row

'レベル毎に背景色変更
     For i = 4 To rmax
     If Cells(i, "B") = "L1" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(47, 117, 181)
     If Cells(i, "B") = "L2" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(155, 194, 230)
     If Cells(i, "B") = "L3" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(189, 215, 238)
     If Cells(i, "B") = "L4" Then Range(Cells(i, "B"), Cells(i, "R")).Interior.Color = RGB(221, 235, 247)
     Next

' 単価調整

    For r = 3 To rmax
    If Range("B" & r) = "" And Range("I" & r) <> "" Then
    Range("O" & r) = "=ROUNDDOWN(" & Range("I" & r).Address(False, False) & "*" & Range("T2").Address & ",-" & Range("U2").Address & ")"
    Range("P" & r) = "=ROUNDDOWN(" & Range("J" & r).Address(False, False) & "*" & Range("T2").Address & ",-" & Range("U2").Address & ")"
    Range("Q" & r) = "=ROUNDDOWN(" & Range("O" & r).Address(False, False) & "*" & Range("K" & r).Address(False, False) & ",0)"
    Range("R" & r) = "=ROUNDDOWN(" & Range("P" & r).Address(False, False) & "*" & Range("L" & r).Address(False, False) & ",0)"
    End If
    Next
    
    
'L4計算
    For r = 1 To rmax
    If Range("B" & r).Value = "L4" Then
    sr = r
    r = r + 1
    r1 = r
    Do Until Range("B" & r).Value = "L4" Or Range("B" & r).Value = "L3" Or Range("B" & r).Value = "L2" Or Range("B" & r).Value = "L1" Or Range("C" & r).Value = "" Or r > rmax
    If Range("B" & r).Value = "" Then
    r2 = r
    End If
    r = r + 1
    Loop
    Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
    Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
    r = r - 1
    End If
    Next r

   
'新L3計算
    For r = 1 To rmax
    If Range("B" & r).Value = "L3" Then
        If Range("B" & r + 1).Value = "L4" Then
        sr = r
        buf1 = "="
        buf2 = "="
        r = r + 1
        
        Do Until Range("B" & r).Value = "L3" Or r > rmax
            If Range("B" & r).Value = "L4" Then
                If buf1 = "=" Then
                buf1 = buf1 & Range("Q" & r).Address(False, False)
                buf2 = buf2 & Range("R" & r).Address(False, False)
                Else
                buf1 = buf1 & "+" & Range("Q" & r).Address(False, False)
                buf2 = buf2 & "+" & Range("R" & r).Address(False, False)
                End If
            End If
        r = r + 1
        Loop
        
        Range("Q" & sr).Value = buf1
        Range("R" & sr).Value = buf2
        r = r - 1
        
        Else
        sr = r
        r = r + 1
        r1 = r
        r2 = 1
        
        Do Until Range("B" & r).Value = "L4" Or Range("B" & r).Value = "L3" Or Range("B" & r).Value = "L2" Or Range("B" & r).Value = "L1" Or Range("C" & r).Value = "" Or r > rmax
            If Range("B" & r).Value = "" Then
            r2 = r
            End If
        r = r + 1
        Loop
        Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
        Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
        r = r - 1
        End If
    
    End If
    Next r
    
    
'新L2計算
    For r = 1 To rmax
    If Range("B" & r).Value = "L2" Then
        If Range("B" & r + 1).Value = "L3" Then
        sr = r
        buf1 = "="
        buf2 = "="
        r = r + 1
        
        Do Until Range("B" & r).Value = "L2" Or r > rmax
            If Range("B" & r).Value = "L3" Then
                If buf1 = "=" Then
                buf1 = buf1 & Range("Q" & r).Address(False, False)
                buf2 = buf2 & Range("R" & r).Address(False, False)
                Else
                buf1 = buf1 & "+" & Range("Q" & r).Address(False, False)
                buf2 = buf2 & "+" & Range("R" & r).Address(False, False)
                End If
            End If
        r = r + 1
        Loop
        
        Range("Q" & sr).Value = buf1
        Range("R" & sr).Value = buf2
        r = r - 1
        
        Else
        sr = r
        r = r + 1
        r1 = r
        r2 = 1
        
        Do Until Range("B" & r).Value = "L4" Or Range("B" & r).Value = "L3" Or Range("B" & r).Value = "L2" Or Range("B" & r).Value = "L1" Or Range("C" & r).Value = "" Or r > rmax
            If Range("B" & r).Value = "" Then
            r2 = r
            End If
        r = r + 1
        Loop
        Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
        Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
        r = r - 1
        End If
    
    End If
    Next r
    

'新L1計算
    For r = 1 To rmax
    If Range("B" & r).Value = "L1" Then
        If Range("B" & r + 1).Value = "L2" Then
        sr = r
        buf1 = "="
        buf2 = "="
        r = r + 1
        
        Do Until Range("B" & r).Value = "L1" Or r > rmax
            If Range("B" & r).Value = "L2" Then
                If buf1 = "=" Then
                buf1 = buf1 & Range("Q" & r).Address(False, False)
                buf2 = buf2 & Range("R" & r).Address(False, False)
                Else
                buf1 = buf1 & "+" & Range("Q" & r).Address(False, False)
                buf2 = buf2 & "+" & Range("R" & r).Address(False, False)
                End If
            End If
        r = r + 1
        Loop
        
        Range("Q" & sr).Value = buf1
        Range("R" & sr).Value = buf2
        r = r - 1
        
        Else
        sr = r
        r = r + 1
        r1 = r
        r2 = 1
        
        Do Until Range("B" & r).Value = "L4" Or Range("B" & r).Value = "L3" Or Range("B" & r).Value = "L2" Or Range("B" & r).Value = "L1" Or Range("C" & r).Value = "" Or r > rmax
            If Range("B" & r).Value = "" Then
            r2 = r
            End If
        r = r + 1
        Loop
        Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
        Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
        r = r - 1
        End If
    
    End If
    Next r
    
rmax = Cells(Rows.Count, 1).End(xlUp).Row
rx1000 = Columns(3).Find("X1000").Row

'Gコード計算
    For r = 1 To rx1000
    If Range("C" & r).Value Like "G*" Then
    sr = r
    r = r + 1
    r1 = r
    r2 = 1
    Do Until Range("C" & r).Value = "" Or r > rx1000
    If Not Range("C" & r).Value = "" Then
    r2 = r
    End If
    r = r + 1
    Loop
    Range("M" & sr).Value = "=SUM(" & Range("M" & r1).Address(False, False) & ":" & Range("M" & r2).Address(False, False) & ")"
    Range("N" & sr).Value = "=SUM(" & Range("N" & r1).Address(False, False) & ":" & Range("N" & r2).Address(False, False) & ")"
    Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
    Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
    r = r - 1
    End If
    Next r

    
'Q列Gコード明細と内訳書のGコードをリンク
    rmax = Cells(Rows.Count, 3).End(xlUp).Row
    rx1000 = Columns(3).Find("X1000").Row
    
    For r = rx1000 To rmax
    If Range("C" & r).Value Like "G*" Then
    sr = r
    aa = Range("C" & r).Value
       If Range("E" & r).Value = "式" Then
         Range("O" & r).Value = ""
         For i = 1 To rx1000
         If Range("C" & i).Value = aa Then
         Range("Q" & sr).Value = "=" & Range("Q" & i).Address(False, False) & "*" & Range("K" & sr).Address(False, False)
         End If
         Next i
       Else
         For i = 1 To rx1000
         If Range("C" & i).Value = aa Then
         Range("O" & sr).Value = "=" & Range("Q" & i).Address(False, False) & "*" & Range("K" & sr).Address(False, False)
         End If
         Next i
       End If

    End If
    Next r
    
'R列Gコード明細と内訳書のGコードをリンク
    For r = rx1000 To rmax
    If Range("C" & r).Value Like "G*" Then
    sr = r
    aa = Range("C" & r).Value
       If Range("E" & r).Value = "式" Then
         Range("P" & r).Value = ""
         For i = 1 To rx1000
         If Range("C" & i).Value = aa Then
         Range("R" & sr).Value = "=" & Range("R" & i).Address(False, False) & "*" & Range("L" & sr).Address(False, False)
         End If
         Next i
       Else
         For i = 1 To rx1000
         If Range("C" & i).Value = aa Then
         Range("P" & sr).Value = "=" & Range("R" & i).Address(False, False) & "*" & Range("L" & sr).Address(False, False)
         End If
         Next i
       End If
    End If
    Next r
    
    
'直接工事費計算
    For r = 1 To rmax
    If Range("D" & r).Value Like "直接工事費計" Then
    sr = r
    End If
    Next r
    
    buf = "="
    For r = 1 To rmax
    If Range("B" & r).Value = "L1" Then
    If buf = "=" Then
    buf = buf & Range("Q" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("Q" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("Q" & sr).Value = buf
    
    buf = "="
    For r = 1 To rmax
    If Range("B" & r).Value = "L1" Then
    If buf = "=" Then
    buf = buf & Range("R" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("R" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("R" & sr).Value = buf
    
    
 
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開


End Sub

Sub ③Zコード計算()

 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ

'Zコード計算
Dim r As Long
Dim r1 As Long
Dim r2 As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim buf3 As String
Dim buf4 As String
Dim rmax As Long
Dim cmax As Long

rmax = Cells(Rows.Count, 3).End(xlUp).Row

'Zコード内のYZコード計算
    For r = 1 To rmax
    If Range("C" & r).Value Like "YZ*" Then
    sr = r
    r = r + 1
    r1 = r
    r2 = 1
    Do Until Range("C" & r).Value Like "YZ*" Or r > rmax Or Range("C" & r) = "" Or Range("C" & r).Value Like "Z0*"
    r2 = r
    r = r + 1
    Loop
    Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
    Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
    Range("M" & sr).Value = "=SUM(" & Range("M" & r1).Address(False, False) & ":" & Range("M" & r2).Address(False, False) & ")"
    Range("N" & sr).Value = "=SUM(" & Range("N" & r1).Address(False, False) & ":" & Range("N" & r2).Address(False, False) & ")"
    r = r - 1
    End If
    Next r

'Zコード計算
    For r = 1 To rmax
    If Range("C" & r).Value Like "Z0*" Then
        If Range("C" & r + 1).Value Like "YZ*" Then
        sr = r
        buf1 = "="
        buf2 = "="
        buf3 = "="
        buf4 = "="
        r = r + 1
        
        Do Until Range("C" & r).Value Like "Z0*" Or r > rmax
            If Range("C" & r).Value Like "YZ*" Then
                If buf1 = "=" Then
                buf1 = buf1 & Range("Q" & r).Address(False, False)
                buf2 = buf2 & Range("R" & r).Address(False, False)
                buf3 = buf3 & Range("M" & r).Address(False, False)
                buf4 = buf4 & Range("N" & r).Address(False, False)
                Else
                buf1 = buf1 & "+" & Range("Q" & r).Address(False, False)
                buf2 = buf2 & "+" & Range("R" & r).Address(False, False)
                buf3 = buf3 & "+" & Range("M" & r).Address(False, False)
                buf4 = buf4 & "+" & Range("N" & r).Address(False, False)
                End If
            End If
        r = r + 1
        Loop
        
        Range("Q" & sr).Value = buf1
        Range("R" & sr).Value = buf2
        Range("M" & sr).Value = buf3
        Range("N" & sr).Value = buf4
        r = r - 1
    
        Else
        
        r2 = 1
        sr = r
        r = r + 1
        r1 = r
        Do Until Range("C" & r).Value Like "Z0*" Or r > rmax Or Range("C" & r) = ""
        r2 = r
        r = r + 1
        Loop
        
        
        
        Range("Q" & sr).Value = "=SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & ")"
        Range("R" & sr).Value = "=SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & ")"
        Range("M" & sr).Value = "=SUM(" & Range("M" & r1).Address(False, False) & ":" & Range("M" & r2).Address(False, False) & ")"
        Range("N" & sr).Value = "=SUM(" & Range("N" & r1).Address(False, False) & ":" & Range("N" & r2).Address(False, False) & ")"
        r = r - 1
        End If
    End If
    
    Next r
    
    For r = 1 To rmax
    If Range("C" & r).Value Like "Z0*" And Range("C" & r + 1).Value Like "Z0*" Then
    Range("Q" & r) = ""
    Range("R" & r) = ""
    Range("M" & r) = ""
    Range("N" & r) = ""
    End If
    Next r
    
    For r = 1 To rmax
    If Range("C" & r).Value Like "Z0*" And Range("C" & r + 1) = "" Then
    Range("Q" & r) = ""
    Range("R" & r) = ""
    Range("M" & r) = ""
    Range("N" & r) = ""
    End If
    Next r
    
    For r = 1 To rmax
    If Range("C" & r).Value Like "Z0047*" Then
    sr = r
    r = r + 1
    r1 = r
    r2 = 1
    Do Until Range("C" & r).Value Like "Z0*" Or r > rmax Or Range("C" & r) = ""
    r2 = r
    r = r + 1
    Loop
    Range("Q" & sr).Value = "=ROUNDDOWN(SUM(" & Range("Q" & r1).Address(False, False) & ":" & Range("Q" & r2).Address(False, False) & "),-3)"
    Range("R" & sr).Value = "=ROUNDDOWN(SUM(" & Range("R" & r1).Address(False, False) & ":" & Range("R" & r2).Address(False, False) & "),-3)"
    Range("M" & sr).Value = "=ROUNDDOWN(SUM(" & Range("M" & r1).Address(False, False) & ":" & Range("M" & r2).Address(False, False) & "),-3)"
    Range("N" & sr).Value = "=ROUNDDOWN(SUM(" & Range("N" & r1).Address(False, False) & ":" & Range("N" & r2).Address(False, False) & "),-3)"
    r = r - 1
    End If
    Next r

 
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開

End Sub

Sub ④諸経費計算()

 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ


Dim r As Long
Dim r1 As Long
Dim r2 As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim rmax As Long
Dim cmax As Long

Dim junkoujihi As Long
Dim genbakanrihi As Long
Dim koujigenkakei As Long
Dim ippankanri As Long
Dim koujikakaku As Long
Dim buf1 As String
Dim buf2 As String
Dim buf3 As String
Dim buf4 As String

rmax = Cells(Rows.Count, 3).End(xlUp).Row



'直接工事費計算（設計単価）

    For r = rmax To 1 Step -1
    If Range("D" & r).Value Like "直接工事費計" Then
    sr = r
    End If
    Next r

    buf = "="
    For r = 1 To rmax
    If Range("B" & r).Value = "L1" Then
    If buf = "=" Then
    buf = buf & Range("M" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("M" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("M" & sr).Value = buf
    
    buf = "="
    For r = 1 To rmax
    If Range("B" & r).Value = "L1" Then
    If buf = "=" Then
    buf = buf & Range("N" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("N" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("N" & sr).Value = buf

'共通仮設費率分計算
    For r = rmax To 1 Step -1
    If Range("D" & r).Value Like "共通仮設費率分" Then
    Range("M" & r).Interior.Color = RGB(255, 255, 0)
    Range("N" & r).Interior.Color = RGB(255, 255, 0)
    Range("Q" & r) = "=ROUNDDOWN(" & Range("M" & r).Address(False, False) & "*" & Range("T2").Address & ",-3)"
    Range("R" & r) = "=ROUNDDOWN(" & Range("N" & r).Address(False, False) & "*" & Range("T2").Address & ",-3)"
    End If
    Next r
    
'共通仮設費計計算
    For r = 1 To rmax
    If Range("D" & r).Value Like "共通仮設費計" Then
    sr = r
    End If
    Next r
    
    buf = "="
    For r = 1 To sr
    If Range("C" & r).Value Like "Z0*" And Not Range("D" & r).Value Like "支給品費*" Then
    If buf = "=" Then
    buf = buf & Range("M" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("M" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("M" & sr).Value = buf & "+" & Range("M" & sr - 1).Address(False, False)

    buf = "="
    For r = 1 To sr
    If Range("C" & r).Value Like "Z0*" And Not Range("D" & r).Value Like "支給品費*" Then
    If buf = "=" Then
    buf = buf & Range("N" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("N" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("N" & sr).Value = buf & "+" & Range("N" & sr - 1).Address(False, False)
    
    buf = "="
    For r = 1 To sr
    If Range("C" & r).Value Like "Z0*" And Not Range("D" & r).Value Like "支給品費*" Then
    If buf = "=" Then
    buf = buf & Range("Q" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("Q" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("Q" & sr).Value = buf & "+" & Range("Q" & sr - 1).Address(False, False)
    
    buf = "="
    For r = 1 To sr
    If Range("C" & r).Value Like "Z0*" And Not Range("D" & r).Value Like "支給品費*" Then
    If buf = "=" Then
    buf = buf & Range("R" & r).Address(False, False)
    Else
    buf = buf & "+" & Range("R" & r).Address(False, False)
    End If
    End If
    Next r
    
    Range("R" & sr).Value = buf & "+" & Range("R" & sr - 1).Address(False, False)

'純工事費計算
    For r = 1 To rmax
    If Range("D" & r).Value Like "純工事費" Then
    sr = r
    junkoujihi = r
    End If
    If Range("D" & r).Value Like "直接工事費計" Then
    r1 = r
    End If
    If Range("D" & r).Value Like "共通仮設費計" Then
    r2 = r
    End If
    Next r
  
    Range("M" & sr) = "=" & Range("M" & r1).Address(False, False) & "+" & Range("M" & r2).Address(False, False)
    Range("N" & sr) = "=" & Range("N" & r1).Address(False, False) & "+" & Range("N" & r2).Address(False, False)
    Range("Q" & sr) = "=" & Range("Q" & r1).Address(False, False) & "+" & Range("Q" & r2).Address(False, False)
    Range("R" & sr) = "=" & Range("R" & r1).Address(False, False) & "+" & Range("R" & r2).Address(False, False)

'現場管理費計算～工事原価計計算
    For r = rmax To 1 Step -1
    If Range("D" & r).Value Like "現場管理費" Then
    Range("M" & r).Interior.Color = RGB(255, 255, 0)
    Range("N" & r).Interior.Color = RGB(255, 255, 0)
    Range("Q" & r) = "=ROUNDDOWN(" & Range("M" & r).Address(False, False) & "*" & Range("T2").Address & ",-3)"
    Range("R" & r) = "=ROUNDDOWN(" & Range("N" & r).Address(False, False) & "*" & Range("T2").Address & ",-3)"
    r1 = r
    genbakanrihi = r
    End If
    Next r
    

    
    For r = 1 To rmax
    If Range("D" & r).Value Like "工事原価計" Then
    koujigenkakei = r
    End If
    Next
    
    
    For r = genbakanrihi To koujigenkakei
    If Range("C" & r).Value Like "Z0*" Then
    buf1 = buf1 + "+" & Range("M" & r).Address(False, False)
    buf2 = buf2 + "+" & Range("N" & r).Address(False, False)
    buf3 = buf3 + "+" & Range("Q" & r).Address(False, False)
    buf4 = buf4 + "+" & Range("Q" & r).Address(False, False)
    End If
    Next
    
    Range("M" & koujigenkakei) = "=" & Range("M" & junkoujihi).Address(False, False) & "+" & Range("M" & genbakanrihi).Address(False, False) & buf1
    Range("N" & koujigenkakei) = "=" & Range("N" & junkoujihi).Address(False, False) & "+" & Range("N" & genbakanrihi).Address(False, False) & buf2
    Range("Q" & koujigenkakei) = "=" & Range("Q" & junkoujihi).Address(False, False) & "+" & Range("Q" & genbakanrihi).Address(False, False) & buf3
    Range("R" & koujigenkakei) = "=" & Range("R" & junkoujihi).Address(False, False) & "+" & Range("R" & genbakanrihi).Address(False, False) & buf4
    
'一般管理費等計算

    rmax = Cells(Rows.Count, 4).End(xlUp).Row
        
    For i = 1 To rmax
    If Range("D" & i).Value Like "工事価格*" Then
    koujikakaku = i
    End If
    Next
    
    Range("Q" & koujikakaku).Interior.Color = RGB(255, 255, 0)
    Range("R" & koujikakaku).Interior.Color = RGB(255, 255, 0)
    
    For i = 1 To rmax
    If Range("D" & i).Value Like "一般管理費等" Then
    Range("M" & i).Interior.Color = RGB(255, 255, 0)
    Range("N" & i).Interior.Color = RGB(255, 255, 0)
    r1 = i
    ippankanrihi = i
    
    buf1 = ""
    buf2 = ""
    
    For r = ippankanrihi To koujikakaku
    If Range("C" & r).Value Like "Z0*" Then
    buf1 = buf1 + "+" & Range("Q" & r).Address(False, False)
    buf2 = buf2 + "+" & Range("R" & r).Address(False, False)
    End If
    Next
    
    Range("Q" & i) = "=" & Range("Q" & koujikakaku).Address(False, False) & "-(" & buf1 & ")" & "-" & Range("Q" & koujigenkakei).Address(False, False)
    Range("R" & i) = "=" & Range("R" & koujikakaku).Address(False, False) & "-(" & buf2 & ")" & "-" & Range("R" & koujigenkakei).Address(False, False)
    
    End If
    Next
    
        
    buf1 = ""
    buf2 = ""
        
    For r = koujigenkakei To koujikakaku
    If Range("C" & r).Value Like "Z0*" Then
    buf1 = buf1 + "+" & Range("M" & r).Address(False, False)
    buf2 = buf2 + "+" & Range("N" & r).Address(False, False)
    End If
    Next
        
    Range("M" & koujikakaku) = "=" & Range("M" & koujigenkakei).Address(False, False) & "+" & Range("M" & ippankanrihi).Address(False, False) & buf1
    Range("N" & koujikakaku) = "=" & Range("N" & koujigenkakei).Address(False, False) & "+" & Range("N" & ippankanrihi).Address(False, False) & buf2

'消費税相当額
    For i = 1 To rmax
    If Range("D" & i).Value Like "消費税*" Then
    sr = i
    End If
    Next
    
    Range("M" & sr) = "=ROUNDDOWN(" & Range("M" & sr - 1).Address(False, False) & "*0.1,0)"
    Range("N" & sr) = "=ROUNDDOWN(" & Range("N" & sr - 1).Address(False, False) & "*0.1,0)"
    Range("Q" & sr) = "=ROUNDDOWN(" & Range("Q" & sr - 1).Address(False, False) & "*0.1,0)"
    Range("R" & sr) = "=ROUNDDOWN(" & Range("R" & sr - 1).Address(False, False) & "*0.1,0)"
    
    Range("M" & rmax) = "=" & Range("M" & sr - 1).Address(False, False) & "+" & Range("M" & sr).Address(False, False)
    Range("N" & rmax) = "=" & Range("N" & sr - 1).Address(False, False) & "+" & Range("N" & sr).Address(False, False)
    Range("Q" & rmax) = "=" & Range("Q" & sr - 1).Address(False, False) & "+" & Range("Q" & sr).Address(False, False)
    Range("R" & rmax) = "=" & Range("R" & sr - 1).Address(False, False) & "+" & Range("R" & sr).Address(False, False)
    
    Range("I3:J" & rmax).NumberFormatLocal = "#,###"
    Range("M3:R" & rmax).NumberFormatLocal = "#,###"
    
 Application.CutCopyMode = False  'セルコピー終了
 
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開
    
End Sub

Sub 業者渡す用修正()

 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ


Dim r As Long
Dim r1 As Long
Dim r2 As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim rmax As Long
Dim cmax As Long

Dim junkoujihi As Long
Dim genbakanrihi As Long
Dim koujigenkakei As Long
Dim ippankanri As Long
Dim koujikakaku As Long
Dim buf1 As String
Dim buf2 As String
Dim buf3 As String
Dim buf4 As String


'";"を判定し、消去
    rmax = Cells(Rows.Count, 3).End(xlUp).Row
    rx1000 = Columns(3).Find("X1000").Row
    
    For r = 3 To rmax
    If Range("F" & r).Value Like "*；*" Then
    Range("F" & r) = ""
    End If
    Next r
    
    For r = 3 To rmax
    If Range("G" & r).Value Like "*；*" Then
    Range("G" & r) = ""
    End If
    Next r

    For r = 3 To rmax
    If Range("H" & r).Value Like "*；*" Then
    Range("H" & r) = ""
    End If
    Next r

 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開

End Sub

Sub OCR()

 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ



Dim rmax As Long
Dim r As Long
Dim rx As Long


'別シート追加
    Worksheets.Add
    ActiveSheet.Name = "Table 2"
    Worksheets("Table 1").Activate

'別シートへ並び替え出力
    rmax = Cells(Rows.Count, 14).End(xlUp).Row
    ActiveSheet.UsedRange.UnMerge
    On Error Resume Next
    
    rx = Columns(18).Find("当初金額", LookAt:=xlPart).Row
    '           ↑の数字を「当初金額」がある列番号に変更する。A列＝1、B列＝2・・・となる。初期値はS列を想定して19。
    
    If Err.Number Then
    MsgBox "S列に「当初金額」は見つかりませんでした。「当初金額」があるセルの列に対応する数字をマクロに入力してください。"
    Exit Sub
    Else
    
    MsgBox rx
    
     For r = rx + 6 To rmax Step 18
     Sheets("Table 2").Range("A" & r - rx - 5).Value = Sheets("Table 1").Range("H" & r).Value
     Sheets("Table 2").Range("B" & r - rx - 5).Value = Sheets("Table 1").Range("H" & r + 1).Value
     Sheets("Table 2").Range("A" & r - rx - 4).Value = Sheets("Table 1").Range("H" & r + 2).Value
     Sheets("Table 2").Range("B" & r - rx - 4).Value = Sheets("Table 1").Range("H" & r + 3).Value
     Sheets("Table 2").Range("A" & r - rx - 3).Value = Sheets("Table 1").Range("H" & r + 4).Value
     Sheets("Table 2").Range("B" & r - rx - 3).Value = Sheets("Table 1").Range("H" & r + 5).Value
     Sheets("Table 2").Range("A" & r - rx - 2).Value = Sheets("Table 1").Range("H" & r + 6).Value
     Sheets("Table 2").Range("B" & r - rx - 2).Value = Sheets("Table 1").Range("H" & r + 7).Value
     Sheets("Table 2").Range("A" & r - rx - 1).Value = Sheets("Table 1").Range("H" & r + 8).Value
     Sheets("Table 2").Range("B" & r - rx - 1).Value = Sheets("Table 1").Range("H" & r + 9).Value
     Sheets("Table 2").Range("A" & r - rx).Value = Sheets("Table 1").Range("H" & r + 10).Value
     Sheets("Table 2").Range("B" & r - rx).Value = Sheets("Table 1").Range("H" & r + 11).Value
     Next r
     
     For r = 7 To rmax Step 6
     Sheets("Table 2").Range(r & ":" & r + 11).Delete
     Next r
    
    End If

    
     
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開

End Sub

Sub 任意工種1式変更()

Cells(Selection.Row, 17).Copy
Cells(Selection.Row, 15).PasteSpecial xlPasteValues

Cells(Selection.Row, 5).Value = "式"
Cells(Selection.Row, 11).Value = "1"

Application.CutCopyMode = False


End Sub

Sub 任意工種1式変更_型枠など一括()

 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ
 
Dim rmax As Long
Dim r As Long

rmax = Cells(Rows.Count, 4).End(xlUp).Row

For r = 1 To rmax
If Range("D" & r) = "型枠" Then

Range("Q" & r).Copy
Range("O" & r).PasteSpecial xlPasteValues
Range("E" & r).Value = "式"
Range("K" & r).Value = "1"

End If
Next r

Application.CutCopyMode = False
     
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開
End Sub

Sub 規格欄3規格欄4追加()

 Application.ScreenUpdating = False  '画面更新停止
 
 Application.DisplayAlerts = False  '警告メッセージオフ

Dim r As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim rmax As Long
Dim cmax As Long
Dim rx1000 As Long
rmax = Cells(Rows.Count, 1).End(xlUp).Row

Columns(8).Insert
Columns(8).Insert
Range("H2") = "規格3"
Range("I2") = "規格4"

Application.CutCopyMode = False
     
 Application.DisplayAlerts = True   '警告メッセージオン

 Application.ScreenUpdating = True  '画面更新再開

End Sub

Sub 規格欄3規格欄4記載()

Application.ScreenUpdating = False  '画面更新停止
Application.DisplayAlerts = False  '警告メッセージオフ

Dim r As Long
Dim i As Long
Dim sr As Long
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim rmax As Long
Dim cmax As Long
Dim ws1x1000 As Long
Dim ws1z0001 As Long
Dim ws2x1000 As Long
Dim ws2z0001 As Long
Dim ws1 As Worksheet
Dim ws2 As Worksheet

Set ws1 = Worksheets("ファイル名")
Set ws2 = Worksheets("内訳明細入力表")

rmax = Cells(Rows.Count, 1).End(xlUp).Row
ws1x1000 = ws1.Columns(3).Find("X1000").Row
ws1z0001 = ws1.Columns(3).Find("Z0001").Row
ws2x1000 = ws2.Columns(1).Find("本工事費").Row
ws2z0001 = ws2.Columns(3).Find("直接工事費計").Row


r = 1
For i = ws1x1000 + 1 To ws1z0001 - 2
    If ws1.Range("D" & i) <> "" And ws1.Range("B" & i) = "" Then
        ws1.Range("A" & i) = r
        r = r + 1
    End If
Next

r = 1
For i = ws2x1000 To ws2z0001
    If ws2.Range("K" & i) <> "" And ws2.Range("K" & i) <> "施工単価名称，規格（施工条件）" And ws2.Range("K" & i) <> "実　施　設　計　工　事　費　内　訳　表" Then
        ws2.Range("EZ" & i) = r
        r = r + 1
    End If
Next

r = 1
For i = ws1x1000 + 1 To ws1z0001 - 2
    If ws1.Range("A" & i) = r Then
        sr = ws2.Columns(156).Find(r).Row
        ws1.Range("H" & i) = ws2.Range("M" & sr + 1) + ws2.Range("M" & sr + 2)
        ws1.Range("I" & i) = ws2.Range("M" & sr + 3) + ws2.Range("M" & sr + 4)
        If ws1.Range("H" & i) = "0" Then
            ws1.Range("H" & i) = ""
        End If
        If ws1.Range("I" & i) = "0" Then
            ws1.Range("I" & i) = ""
        End If
        r = r + 1
    End If
Next


Application.CutCopyMode = False
Application.DisplayAlerts = True   '警告メッセージオン
Application.ScreenUpdating = True  '画面更新再開

End Sub


Sub ☆施工条件追記()

Call 規格欄3規格欄4追加
Call 規格欄3規格欄4記載

End Sub




Option Explicit

' コードの版数。貼り替え忘れの確認用に、更新のたびに増やす。
' 実行後のメッセージボックスにこの番号が表示される。
Const MACRO_VERSION As String = "v59"

' 「ファイル名」シートの２行目で「施工単価名称」列を探し、
' そのセルの文字列に指定キーワードを含む行を丸ごと、
' セル参照の数式（=ファイル名!$A$5 形式）として
' 「再生資源」シート（無ければ新規作成）へコピーする。
' 数式で参照するため、元データを変更すると自動的に反映される。
Sub 再生資源()

    Const SRC_SHEET_NAME As String = "ファイル名"
    Const DEST_SHEET_NAME As String = "再生資源"
    Const HEADER_ROW As Long = 2
    Const HEADER_TEXT As String = "施工単価名称"

    Dim keywords As Variant
    keywords = Array("街渠工", "舗装復旧工", "先行路盤工", "殻運搬処理", _
                      "床掘", "土砂運搬処理")

    Dim wb As Workbook
    Set wb = ActiveWorkbook

    Dim wsSrc As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If Trim(ws.Name) = SRC_SHEET_NAME Then
            Set wsSrc = ws
            Exit For
        End If
    Next ws

    If wsSrc Is Nothing Then
        Dim sheetList As String
        For Each ws In wb.Worksheets
            sheetList = sheetList & "・" & ws.Name & vbCrLf
        Next ws
        MsgBox "シート「" & SRC_SHEET_NAME & "」が見つかりません。処理を中止します。" & vbCrLf & vbCrLf & _
               "「" & wb.Name & "」ブック内の実際のシート名:" & vbCrLf & sheetList, vbExclamation
        Exit Sub
    End If

    Dim wsDest As Worksheet
    For Each ws In wb.Worksheets
        If Trim(ws.Name) = DEST_SHEET_NAME Then
            Set wsDest = ws
            Exit For
        End If
    Next ws

    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDest.Name = DEST_SHEET_NAME
    End If

    ' 実行するたびに「再生資源」シートの内容を全て消してから作り直す
    wsDest.Cells.Clear

    ' １行目・２行目のうち、より右まで使われている方に合わせて列数を決定
    Dim lastColSrc As Long
    Dim lastColRow1 As Long
    lastColRow1 = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    lastColSrc = wsSrc.Cells(HEADER_ROW, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastColRow1 > lastColSrc Then lastColSrc = lastColRow1

    Dim targetCol As Long
    targetCol = 0

    Dim c As Long
    For c = 1 To lastColSrc
        If InStr(1, CStr(wsSrc.Cells(HEADER_ROW, c).Value), HEADER_TEXT, vbTextCompare) > 0 Then
            targetCol = c
            Exit For
        End If
    Next c

    If targetCol = 0 Then
        MsgBox "シート「" & SRC_SHEET_NAME & "」の" & HEADER_ROW & "行目に「" & _
               HEADER_TEXT & "」列が見つかりません。", vbExclamation
        Exit Sub
    End If

    ' １行目・２行目（見出し）を数式リンクでコピー
    For c = 1 To lastColSrc
        wsDest.Cells(1, c).Formula = "='" & SRC_SHEET_NAME & "'!" & wsSrc.Cells(1, c).Address
        wsDest.Cells(2, c).Formula = "='" & SRC_SHEET_NAME & "'!" & wsSrc.Cells(HEADER_ROW, c).Address
    Next c

    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, targetCol).End(xlUp).Row

    Dim outRow As Long
    outRow = 3

    Dim r As Long, i As Long
    Dim cellText As String
    Dim matched As Boolean

    For r = HEADER_ROW + 1 To lastRowSrc
        cellText = CStr(wsSrc.Cells(r, targetCol).Value)
        If Len(cellText) > 0 Then

            matched = False
            For i = LBound(keywords) To UBound(keywords)
                If InStr(1, cellText, CStr(keywords(i)), vbTextCompare) > 0 Then
                    matched = True
                    Exit For
                End If
            Next i

            ' E～H列のいずれかに「発生材」を含む行は対象から除外する
            If matched Then
                Dim ecol As Long
                For ecol = 5 To 8
                    If InStr(1, CStr(wsSrc.Cells(r, ecol).Value), "発生材", vbTextCompare) > 0 Then
                        matched = False
                        Exit For
                    End If
                Next ecol
            End If

            If matched Then
                ' 行全体（全列）をセル参照の数式でコピー
                For c = 1 To lastColSrc
                    wsDest.Cells(outRow, c).Formula = "='" & SRC_SHEET_NAME & "'!" & wsSrc.Cells(r, c).Address
                Next c
                outRow = outRow + 1
            End If
        End If
    Next r

    ' データ列の右端の次から、材料数量集計用の見出しを追記
    ' （シート全体を実行冒頭で消しているので、毎回まっさらな状態に書く）
    Dim extraCol As Long
    extraCol = lastColSrc + 1

    Dim extraHeaders(1 To 19, 1 To 2) As String
    extraHeaders(1, 1) = "": extraHeaders(1, 2) = "単位Co量(m3/施工単位)"
    extraHeaders(2, 1) = "": extraHeaders(2, 2) = "Co量(m3)"
    extraHeaders(3, 1) = "": extraHeaders(3, 2) = "処分無筋Co量(m3)"
    extraHeaders(4, 1) = "": extraHeaders(4, 2) = "処分鉄筋Co量(m3)"
    extraHeaders(5, 1) = "粗粒度": extraHeaders(5, 2) = "単位As量(t/m2)"
    extraHeaders(6, 1) = "粗粒度": extraHeaders(6, 2) = "As量(t)"
    extraHeaders(7, 1) = "密粒度": extraHeaders(7, 2) = "単位As量(t/m2)"
    extraHeaders(8, 1) = "密粒度": extraHeaders(8, 2) = "As量(t)"
    extraHeaders(9, 1) = "細粒度": extraHeaders(9, 2) = "単位As量(t/m2)"
    extraHeaders(10, 1) = "細粒度": extraHeaders(10, 2) = "As量(t)"
    extraHeaders(11, 1) = "開粒度": extraHeaders(11, 2) = "単位As量(t/m2)"
    extraHeaders(12, 1) = "開粒度": extraHeaders(12, 2) = "As量(t)"
    extraHeaders(13, 1) = "改質アスコン": extraHeaders(13, 2) = "単位As量(t/m2)"
    extraHeaders(14, 1) = "改質アスコン": extraHeaders(14, 2) = "As量(t)"
    extraHeaders(15, 1) = "": extraHeaders(15, 2) = "処分As量(t)"
    extraHeaders(16, 1) = "": extraHeaders(16, 2) = "再生砕石量(m3)"
    extraHeaders(17, 1) = "": extraHeaders(17, 2) = "粒調砕石量(m3)"
    extraHeaders(18, 1) = "": extraHeaders(18, 2) = "掘削土量(m3)"
    extraHeaders(19, 1) = "": extraHeaders(19, 2) = "処分土量(m3)"

    For i = 1 To 19
        If Len(extraHeaders(i, 1)) > 0 Then wsDest.Cells(1, extraCol + i - 1).Value = extraHeaders(i, 1)
        If Len(extraHeaders(i, 2)) > 0 Then wsDest.Cells(2, extraCol + i - 1).Value = extraHeaders(i, 2)
    Next i

    ' 材料数量集計用見出し（19列、extraCol～extraCol+18）のすぐ右の
    ' 固定位置に、単位数量の見出しを縦に記載する。列位置を毎回スキャンで
    ' 探すのではなく固定にすることで、既存データの内容に影響されず、
    ' 再実行しても列がずれたり増えたりしない
    Dim unitLabels As Variant
    unitLabels = Array("単位Co量(m3/施工単位)", "粗粒度単位As量(t/m2)", "密粒度単位As量(t/m2)", _
                        "細粒度単位As量(t/m2)", "開粒度単位As量(t/m2)", "改質アスコン単位As量(t/m2)", _
                        "処分As単位量(t/m3)")

    Dim unitLabelCol As Long
    unitLabelCol = extraCol + 19

    Dim labelRow As Long
    labelRow = 2
    For i = LBound(unitLabels) To UBound(unitLabels)
        wsDest.Cells(labelRow + i, unitLabelCol).Value = unitLabels(i)
    Next i

    ' 一番端の「単位Co量」セルの横に、数値・単位・参照先を記載
    wsDest.Cells(labelRow, unitLabelCol + 1).Value = "0.138"
    wsDest.Cells(labelRow, unitLabelCol + 2).Value = "m3/m"
    wsDest.Cells(labelRow, unitLabelCol + 3).Value = "標準図 NG-L-FA参照"

    ' 「改質アスコン単位As量」の下の「処分As単位量(t/m3)」セルの横に、
    ' 数値・単位・注記を記載
    wsDest.Cells(labelRow + 6, unitLabelCol + 1).Value = "2.30"
    wsDest.Cells(labelRow + 6, unitLabelCol + 2).Value = "t/m3"
    wsDest.Cells(labelRow + 6, unitLabelCol + 3).Value = "(歩車道密粒細粒などの平均として)"

    ' D列に「街渠工」を含む行の「単位Co量」列（extraCol）に、
    ' 上で記載した0.138セルを絶対参照する数式を入れる
    Dim unitCoRefAddr As String
    unitCoRefAddr = wsDest.Cells(labelRow, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

    ' 同じ行の「単位Co量」列（extraCol）の右隣（Co量(m3)列）に、
    ' 単位Co量×L列 の数式を入れる
    For r = 3 To outRow - 1
        If InStr(1, CStr(wsDest.Cells(r, 4).Value), "街渠工", vbTextCompare) > 0 Then
            wsDest.Cells(r, extraCol).Formula = "=" & unitCoRefAddr
            wsDest.Cells(r, extraCol + 1).Formula = "=" & wsDest.Cells(r, extraCol).Address(False, False) & _
                                                     "*" & wsDest.Cells(r, 12).Address(False, False)
        End If
    Next r

    ' D列に「殻運搬処理」を含み、かつE列またはI列に「Co」を含む行の
    ' 「処分無筋Co量(m3)」列に、その行のL列を絶対参照する数式を入れる
    Dim disposalConcreteCol As Long
    disposalConcreteCol = FindHeaderColumn(wsDest, extraCol, "処分無筋Co量(m3)")

    If disposalConcreteCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "殻運搬処理", vbTextCompare) > 0 And _
               (InStr(1, StrConv(CStr(wsDest.Cells(r, 5).Value), vbNarrow), "Co", vbTextCompare) > 0 Or _
                InStr(1, StrConv(CStr(wsDest.Cells(r, 9).Value), vbNarrow), "Co", vbTextCompare) > 0) Then
                wsDest.Cells(r, disposalConcreteCol).Formula = "=$L$" & r
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつ（G列に「２－２号工」もしくは「３号工」、
    ' またはE列に「10-1号工(乗入部)」を含む）行の「粗粒度」列
    ' （１行目=粗粒度、２行目=単位As量(t/m2)）に、粗粒度単位As量の
    ' 0.115セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim roughAsCol As Long
    roughAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "粗粒度", "単位As量(t/m2)")

    If roughAsCol > 0 Then
        Dim roughAsRefAddr As String
        roughAsRefAddr = wsDest.Cells(labelRow + 1, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "2-2号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "3号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "10-1号工(乗入部)", vbTextCompare) > 0) Then
                wsDest.Cells(r, roughAsCol).Formula = "=" & roughAsRefAddr
                wsDest.Cells(r, roughAsCol + 1).Formula = "=" & wsDest.Cells(r, roughAsCol).Address(False, False) & _
                                                           "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「３号工」を含む行の「密粒度」列
    ' （１行目=密粒度、２行目=単位As量(t/m2)）に、密粒度単位As量の
    ' 0.118セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim fineAsCol As Long
    fineAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "密粒度", "単位As量(t/m2)")

    If fineAsCol > 0 Then
        Dim fineAsRefAddr As String
        fineAsRefAddr = wsDest.Cells(labelRow + 2, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "3号工", vbTextCompare) > 0 Then
                wsDest.Cells(r, fineAsCol).Formula = "=" & fineAsRefAddr
                wsDest.Cells(r, fineAsCol + 1).Formula = "=" & wsDest.Cells(r, fineAsCol).Address(False, False) & _
                                                          "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「５号工」もしくは「１０号工」を含む
    ' 行の「細粒度」列（１行目=細粒度、２行目=単位As量(t/m2)）に、細粒度単位As量の
    ' 0.115セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim fineFineAsCol As Long
    fineFineAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "細粒度", "単位As量(t/m2)")

    If fineFineAsCol > 0 Then
        Dim fineFineAsRefAddr As String
        fineFineAsRefAddr = wsDest.Cells(labelRow + 3, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "5号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "10号工", vbTextCompare) > 0) Then
                wsDest.Cells(r, fineFineAsCol).Formula = "=" & fineFineAsRefAddr

                ' 右隣（As量(t)列）に、G列の記載内容に応じた係数付きの数式を入れる
                Dim fineFineG As String
                fineFineG = NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value))

                Dim fineFineAjAddr As String, fineFineLAddr As String
                fineFineAjAddr = wsDest.Cells(r, fineFineAsCol).Address(False, False)
                fineFineLAddr = wsDest.Cells(r, 12).Address(False, False)

                If InStr(1, fineFineG, "4cm", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr & "*4/5"
                ElseIf InStr(1, fineFineG, "5cm", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr
                ElseIf InStr(1, fineFineG, "10cm", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr & "*2"
                ElseIf InStr(1, fineFineG, "10号工", vbTextCompare) > 0 Then
                    wsDest.Cells(r, fineFineAsCol + 1).Formula = "=" & fineFineAjAddr & "*" & fineFineLAddr & "*4/5"
                End If
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「２－２号工」もしくは「１０－１号工」を
    ' 含む行の「開粒度」列（１行目=開粒度、２行目=単位As量(t/m2)）に、開粒度単位As量の
    ' 0.097セルを絶対参照する数式を入れる（全角/半角・ダッシュの表記ゆれを吸収）
    Dim coarseAsCol As Long
    coarseAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "開粒度", "単位As量(t/m2)")

    If coarseAsCol > 0 Then
        Dim coarseAsRefAddr As String
        coarseAsRefAddr = wsDest.Cells(labelRow + 4, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "2-2号工", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "10-1号工", vbTextCompare) > 0) Then
                wsDest.Cells(r, coarseAsCol).Formula = "=" & coarseAsRefAddr

                ' 右隣（As量(t)列）に、G列の記載内容に応じた係数付きの数式を入れる
                Dim coarseG As String
                coarseG = NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value))

                Dim coarseAlAddr As String, coarseLAddr As String
                coarseAlAddr = wsDest.Cells(r, coarseAsCol).Address(False, False)
                coarseLAddr = wsDest.Cells(r, 12).Address(False, False)

                If InStr(1, coarseG, "2-2号工", vbTextCompare) > 0 Then
                    wsDest.Cells(r, coarseAsCol + 1).Formula = "=" & coarseAlAddr & "*" & coarseLAddr & "*2"
                ElseIf InStr(1, coarseG, "10-1号工", vbTextCompare) > 0 Then
                    wsDest.Cells(r, coarseAsCol + 1).Formula = "=" & coarseAlAddr & "*" & coarseLAddr & "*4/5"
                End If
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつG列に「７号工」を含む行の
    ' 「Co量(m3)」列に、＝L列×0.15 の数式を入れる（全角/半角の表記ゆれを吸収）
    Dim coQuantityCol As Long
    coQuantityCol = FindHeaderColumn(wsDest, extraCol, "Co量(m3)")

    If coQuantityCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "7号工", vbTextCompare) > 0 Then
                wsDest.Cells(r, coQuantityCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False) & "*0.15"
            End If
        Next r
    End If

    ' D列に「舗装復旧工」を含み、かつE列に「１０－１号工」と「乗入」の両方を
    ' 含む行の「改質アスコン」列（１行目=改質アスコン、２行目=単位As量(t/m2)）に、
    ' 改質アスコン単位As量の0.115セルを絶対参照する数式を入れる
    ' （全角/半角・ダッシュの表記ゆれを吸収）
    Dim modifiedAsCol As Long
    modifiedAsCol = FindHeaderColumnByRow1Row2(wsDest, extraCol, "改質アスコン", "単位As量(t/m2)")

    If modifiedAsCol > 0 Then
        Dim modifiedAsRefAddr As String
        modifiedAsRefAddr = wsDest.Cells(labelRow + 5, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 4).Value)), "舗装復旧工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "10-1号工", vbTextCompare) > 0 And _
               InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "乗入", vbTextCompare) > 0 Then
                wsDest.Cells(r, modifiedAsCol).Formula = "=" & modifiedAsRefAddr
                wsDest.Cells(r, modifiedAsCol + 1).Formula = "=" & wsDest.Cells(r, modifiedAsCol).Address(False, False) & _
                                                              "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「殻運搬処理」を含み、かつE列・F列・G列・H列のいずれかに「As」を
    ' 含む行の「処分As量(t)」列に、＝処分As単位量(2.30)×L列 の数式を入れる
    ' （全角/半角の表記ゆれを吸収）
    Dim disposalAsCol As Long
    disposalAsCol = FindHeaderColumn(wsDest, extraCol, "処分As量(t)")

    If disposalAsCol > 0 Then
        Dim disposalAsUnitRefAddr As String
        disposalAsUnitRefAddr = wsDest.Cells(labelRow + 6, unitLabelCol + 1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "殻運搬処理", vbTextCompare) > 0 And _
               (InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 5).Value)), "As", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 6).Value)), "As", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "As", vbTextCompare) > 0 Or _
                InStr(1, NormalizeForMatch(CStr(wsDest.Cells(r, 8).Value)), "As", vbTextCompare) > 0) Then
                wsDest.Cells(r, disposalAsCol).Formula = "=" & disposalAsUnitRefAddr & "*" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「先行路盤」を含み、かつG列に「再生砕石」を含む行の「再生砕石量(m3)」
    ' 列に、＝L×0.01×(G列のcmの直前の数値) の数式を入れる（全角/半角の表記ゆれを吸収）
    Dim recycledCrushedCol As Long
    recycledCrushedCol = FindHeaderColumn(wsDest, extraCol, "再生砕石量(m3)")

    If recycledCrushedCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "先行路盤", vbTextCompare) > 0 And _
               InStr(1, CStr(wsDest.Cells(r, 7).Value), "再生砕石", vbTextCompare) > 0 Then
                Dim recycledCmNumber As String
                recycledCmNumber = ExtractNumberBeforeUnit(NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "cm")

                If Len(recycledCmNumber) > 0 Then
                    wsDest.Cells(r, recycledCrushedCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False) & _
                                                                   "*0.01*" & recycledCmNumber
                End If
            End If
        Next r
    End If

    ' D列に「先行路盤」を含み、かつG列に「粒調砕石」を含む行の「粒調砕石量(m3)」
    ' 列に、＝L×0.01×(G列のcmの直前の数値) の数式を入れる（全角/半角の表記ゆれを吸収）
    Dim gradedCrushedCol As Long
    gradedCrushedCol = FindHeaderColumn(wsDest, extraCol, "粒調砕石量(m3)")

    If gradedCrushedCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "先行路盤", vbTextCompare) > 0 And _
               InStr(1, CStr(wsDest.Cells(r, 7).Value), "粒調砕石", vbTextCompare) > 0 Then
                Dim gradedCmNumber As String
                gradedCmNumber = ExtractNumberBeforeUnit(NormalizeForMatch(CStr(wsDest.Cells(r, 7).Value)), "cm")

                If Len(gradedCmNumber) > 0 Then
                    wsDest.Cells(r, gradedCrushedCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False) & _
                                                                 "*0.01*" & gradedCmNumber
                End If
            End If
        Next r
    End If

    ' D列に「床掘」を含む行の「掘削土量(m3)」列に、＝L の数式を入れる
    Dim excavationCol As Long
    excavationCol = FindHeaderColumn(wsDest, extraCol, "掘削土量(m3)")

    If excavationCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "床掘", vbTextCompare) > 0 Then
                wsDest.Cells(r, excavationCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' D列に「土砂運搬処理」を含む行の「処分土量(m3)」列に、＝L の数式を入れる
    Dim disposalSoilCol As Long
    disposalSoilCol = FindHeaderColumn(wsDest, extraCol, "処分土量(m3)")

    If disposalSoilCol > 0 Then
        For r = 3 To outRow - 1
            If InStr(1, CStr(wsDest.Cells(r, 4).Value), "土砂運搬処理", vbTextCompare) > 0 Then
                wsDest.Cells(r, disposalSoilCol).Formula = "=" & wsDest.Cells(r, 12).Address(False, False)
            End If
        Next r
    End If

    ' 粗粒度/密粒度/細粒度/開粒度/改質アスコン、各単位As量セルの横に、数値・単位・注記を記載
    Dim asValues As Variant
    asValues = Array("0.115", "0.118", "0.115", "0.097", "0.115")

    For i = 1 To 5
        wsDest.Cells(labelRow + i, unitLabelCol + 1).Value = asValues(i - 1)
        wsDest.Cells(labelRow + i, unitLabelCol + 2).Value = "t/m2"
        wsDest.Cells(labelRow + i, unitLabelCol + 3).Value = "t=5cm"
    Next i

    ' データ最終行の次の行（outRow）に、AB?AT列（extraCol?extraCol+18）
    ' それぞれのSUM式を入れる（２行目の見出しに「単位」を含む列は、
    ' 単価・原単位のような比率であり合計する意味がないため除外する）
    Dim sumRow As Long
    sumRow = outRow

    For c = extraCol To extraCol + 18
        If InStr(1, CStr(wsDest.Cells(2, c).Value), "単位", vbTextCompare) = 0 Then
            wsDest.Cells(sumRow, c).Formula = "=SUM(" & wsDest.Cells(3, c).Address(False, False) & _
                                               ":" & wsDest.Cells(outRow - 1, c).Address(False, False) & ")"
        End If
    Next c

    ' SUM行の下２行に、その列の１行目・２行目の見出しを数式で持ってくる
    For c = extraCol To extraCol + 18
        wsDest.Cells(sumRow + 1, c).Formula = "=" & wsDest.Cells(1, c).Address(False, False)
        wsDest.Cells(sumRow + 2, c).Formula = "=" & wsDest.Cells(2, c).Address(False, False)
    Next c

    wsDest.Columns.AutoFit

    MsgBox (outRow - 3) & " 件の行を「" & DEST_SHEET_NAME & "」シートにリンク（数式）でコピーしました。" & vbCrLf & _
           "(" & MACRO_VERSION & ")", vbInformation

End Sub

' ２行目が headerText と一致する列番号を返す。見つからなければ0を返す。
Private Function FindHeaderColumn(ws As Worksheet, searchFromCol As Long, headerText As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(2, c).Value) = headerText Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

' １行目が row1Text、２行目が row2Text と一致する列番号を返す。
' （"単位As量(t/m2)"のように２行目だけでは複数該当する見出しを区別するため）
' 見つからなければ0を返す。
Private Function FindHeaderColumnByRow1Row2(ws As Worksheet, searchFromCol As Long, _
                                             row1Text As String, row2Text As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = searchFromCol To lastCol
        If CStr(ws.Cells(1, c).Value) = row1Text And CStr(ws.Cells(2, c).Value) = row2Text Then
            FindHeaderColumnByRow1Row2 = c
            Exit Function
        End If
    Next c
    FindHeaderColumnByRow1Row2 = 0
End Function

' 全角英数字・記号を半角に揃え、長音記号やダッシュ類も半角ハイフンに
' 統一してから比較できるようにする（表記ゆれ対策）。
' StrConv(vbNarrow)は環境によって全角英数字を変換しないことがあるため、
' 文字コード（Unicode）を直接シフトして半角化する。
Private Function NormalizeForMatch(ByVal s As String) As String
    Dim result As String
    result = ""

    Dim i As Long, code As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch)
        If code >= &HFF01 And code <= &HFF5E Then
            ' 全角の！～～（英数字・記号を含む）を半角へ
            ch = ChrW(code - &HFEE0)
        ElseIf code = &H3000 Then
            ' 全角スペース -> 半角スペース
            ch = " "
        End If
        result = result & ch
    Next i

    result = Replace(result, "‐", "-")
    result = Replace(result, "?", "-")
    result = Replace(result, "?", "-")
    result = Replace(result, "?", "-")
    result = Replace(result, "ー", "-")
    result = Replace(result, "ｰ", "-")

    ' cm/m2/m3などの単位記号（CJK互換文字の１文字表記）を通常の英字表記に展開
    result = Replace(result, "㎝", "cm")
    result = Replace(result, "㎜", "mm")
    result = Replace(result, "㎡", "m2")
    result = Replace(result, "?", "m3")

    NormalizeForMatch = result
End Function

' normalizedText内でunitText（例："cm"）の直前に連続する数字（と小数点）を取り出す。
' 見つからなければ空文字を返す。
Private Function ExtractNumberBeforeUnit(ByVal normalizedText As String, ByVal unitText As String) As String
    Dim pos As Long
    pos = InStr(1, normalizedText, unitText, vbTextCompare)
    If pos = 0 Then
        ExtractNumberBeforeUnit = ""
        Exit Function
    End If

    Dim numStr As String
    numStr = ""

    Dim j As Long, ch As String
    j = pos - 1
    Do While j >= 1
        ch = Mid(normalizedText, j, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            numStr = ch & numStr
            j = j - 1
        Else
            Exit Do
        End If
    Loop

    ExtractNumberBeforeUnit = numStr
End Function

