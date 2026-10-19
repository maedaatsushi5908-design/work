'==============================================================
' M60_Help  ―  ブックの中に「使い方」シートを作る
'--------------------------------------------------------------
' 標準モジュールとして貼り付け、モジュール名を "M60_Help" に。
'
' ［使い方シート作成］を実行すると、ブックの先頭に「使い方」シートを作ります。
' ［設定シート作成］のときにも、無ければ自動で作ります。
' 内容を書き換えても、もう一度実行すれば元に戻ります。
'==============================================================
Option Explicit

Public Const SH_HELP As String = "使い方"

' 使い方シートを組み立てるときの作業用
Private mHws As Worksheet
Private mHr As Long

' 段落の折り返し幅（半角文字の数）。A〜D列の幅の合計に合わせる
Private Const H_WRAP As Long = 112

Private Const H_C1 As Long = 26     ' A列の幅
Private Const H_C2 As Long = 52     ' B列の幅
Private Const H_C3 As Long = 17     ' C列の幅
Private Const H_C4 As Long = 17     ' D列の幅


'==============================================================
' 入口
'==============================================================
Public Sub 使い方シート作成()
    BuildHelpSheet False
End Sub


Public Sub BuildHelpSheet(ByVal quiet As Boolean)
    Application.ScreenUpdating = False

    Set mHws = FreshSheet(SH_HELP)
    mHr = 0

    HTitle "インフレスライド設計書　作成マクロ　使い方"
    HNote "このシートは［使い方シート作成］で作り直せます。書き換えても元に戻せます。"
    HHead2 "最新版の置き場所", "GitHub"
    HRow "VBA（貼り付けるファイル）", "infla-slide/vba/ALL_IN_ONE.bas"
    HRow "手順書（このシートの元）", "infla-slide/docs/使い方.md"
    HRow "リポジトリ", "https://github.com/maedaatsushi5908-design/work"
    HBlank

    WriteHelpOverview
    WriteHelpPrepare
    WriteHelpSteps
    WriteHelpInputs
    WriteHelpSheetMap
    WriteHelpRules
    WriteHelpConfig
    WriteHelpMacros
    WriteHelpTrouble

    FinishHelpSheet

    Application.ScreenUpdating = True
    If Not quiet Then
        mHws.Activate
        mHws.Range("A1").Select
        MsgBox "「" & SH_HELP & "」シートを作りました。" & vbCrLf & _
               "書き換えても、もう一度［使い方シート作成］を実行すれば元に戻ります。", vbInformation
    End If
End Sub


'==============================================================
' 中身
'==============================================================
Private Sub WriteHelpOverview()
    HSec "0. これは何をするマクロか"
    HP "エスティマのCSVから、スライド計算表（明細＋経費計算）と" & _
       "スライド調書（様式4-2号）を作ります。"
    HBlank
    HHead2 "項目", "内容"
    HRow "入力するCSV", "スライド用csv 2本（旧単価・新単価）" & vbLf & _
                        "（任意）特殊集計区分一覧表csv 2本（スライド前・スライド後）"
    HRow "手で入れるもの", "出来形数量（L列）と、調書の請負代金額（C14）だけ"
    HRow "出てくる答え", "スライド調書（様式4-2号）の G16 ＝ スライド額 S'"
    HBlank
End Sub


Private Sub WriteHelpPrepare()
    HSec "1. 準備（最初の1回だけ）"

    HSub "1-1　マクロ有効ブックとして保存する"
    HP "「名前を付けて保存」→ ファイルの種類を「Excel マクロ有効ブック (*.xlsm)」にします。" & vbLf & _
       "通常の .xlsx で保存するとマクロが消えます。"
    HBlank

    HSub "1-2　VBAを貼り付ける（貼るファイルは1つだけ）"
    HHead2 "版", "GitHubのファイル名"
    HRow "新工種を考慮する（通常）", "infla-slide/vba/ALL_IN_ONE.bas"
    HRow "新工種を考慮しない", "infla-slide/vba/ALL_IN_ONE_新工種なし.bas"
    HP "中身は同じで、設定シートの「新工種を考慮する」の初期値だけが違います。" & _
       "どちらを貼っても、あとから設定シートで切り替えられます。"
    HBlank
    HP "GitHubでファイルを開くと、コードの枠の右上にボタンが並んでいます。" & vbLf & _
       "四角が2枚重なったアイコン（Copy raw file）を押すとファイル全体がコピーされます。" & vbLf & _
       "［Raw］を押して文字だけのページにしてから Ctrl+A → Ctrl+C でも同じです。"
    HBlank
    HHead2 "貼り付けの手順", "操作"
    HRow "1", "Excelで Alt+F11 を押す（VBEが開く）"
    HRow "2", "メニューの「挿入」→「標準モジュール」"
    HRow "3", "右側の白い画面をクリックして Ctrl+V"
    HRow "4", "Alt+F11 でExcelに戻る"
    HP "モジュール名の変更も、参照設定の追加も要りません。"
    HBlank

    HSub "1-3　マクロを有効にする"
    HP "ブックを開き直して「セキュリティの警告 マクロが無効にされました」が出たら、" & _
       "「コンテンツの有効化」を押します。"
    HBlank
End Sub


Private Sub WriteHelpSteps()
    HSec "2. 毎回の手順"

    HSub "手順1　エスティマからCSVを出す"
    HCode "機能 → CSV連携 → 抽出指示ファイル「スライド用csv出力」→ csv抽出"
    HP "単価適用日を変えて2回出力します。数量・工種構成は同じで、単価だけが違うCSVになります。"
    HHead3 "CSV", "単価適用日", "ファイル名の例"
    HRow3 "旧単価CSV", "当初", "〇〇工事_旧単価.csv"
    HRow3 "新単価CSV", "スライド基準日", "〇〇工事_新単価.csv"
    HP "特殊集計区分一覧表のCSVも出せる場合は、同じく単価適用日を変えて2本出しておくと、" & _
       "処分費・管材費の〇と金額が自動で入ります（手順4）。"
    HBlank

    HSub "手順2　設定シートを作る"
    HCode "Alt+F8 →「設定シート作成」→ 実行"
    HP "「設定」シートができます。黄色いセルに入力します。最低限いるのは次の2つで、" & _
       "残りは既定のままで動きます。"
    HHead2 "項目", "入れるもの"
    HRow "請負代金額", "税込。様式4-2号の請負率の計算に使います"
    HRow "基準日", "インフレスライドの請求日"
    HP "経費率の係数は道路改良工事・公園工事等の値が入っています。" & _
       "違う工種区分なら、ここを直すか、あとで経費計算部の率セルを直接上書きしてください。"
    HBlank

    HSub "手順3　スライド計算表を作る"
    HCode "Alt+F8 →「スライド計算表作成」→ 実行"
    HHead2 "操作", "内容"
    HRow "1", "「【旧単価】…」→ 旧単価CSVを選ぶ"
    HRow "2", "「【新単価】…」→ 新単価CSVを選ぶ"
    HBlank
    HHead2 "できるシート", "中身"
    HRow "スライド計算表", "明細部＋経費計算部（①〜㉔を7系列）"
    HRow "スライド調書（様式4-2号）", "計算表からリンク。G16がスライド額"
    HRow "_単価表（非表示）", "代価表の中身の索引。手順4で使います"
    HP "実行後にメッセージが出ます。★が付いた件数があれば必ず確認してください" & _
       "（新単価が見つからなかった行など。9章）。"
    HBlank

    HSub "手順4　特殊集計区分CSVを取り込む（任意・あると楽）"
    HCode "Alt+F8 →「特殊集計区分CSV取込」→ 実行"
    HHead2 "操作", "内容"
    HRow "1", "【スライド前】の一覧表CSVを選ぶ → AA列（処分費額 前）に入ります"
    HRow "2", "【スライド後】の一覧表CSVを選ぶ → AB列（処分費額 後）に入ります" & vbLf & _
              "　無ければキャンセル。AB列は空のままで、計算式はAA列の額を使います"
    HRow "3", "「いま入っている〇を消してから入れ直すか」→ 初回は［はい］"
    HRow "4", "_特殊集計取込結果 シートに結果が出ます"
    HBlank
    HHead2 "集計区分名称", "入るところ"
    HRow "直工内の処分費等", "O列に〇、AA列（前）・AB列（後）に金額"
    HRow "水道工事における管財費", "P列に〇"
    HRow "特殊集計区分外", "何も入れません"
    HBlank
    HWarn "スクラップには処分費の〇を付けません。売却益なので処分費ではなく、㉒で別に扱います。"
    HWarn "金額の列が空だった行は赤字で出ます。AA列を必ず手入力してください。" & _
          "空のままだとその行の金額を全額 処分費として扱い、過大になります。"
    HP "突合できなかった行には、さがしたコードがそのまま出ます。" & _
       "まず一覧表とスライド用csvが同じ設計書から出ているか確かめてください。" & vbLf & _
       "先に［特殊集計区分CSV構造チェック］を実行すると、列の並びだけを書き出します" & _
       "（金額・数量・単価は「（数値）」に置き換わるのでそのまま送れます）。"
    HBlank

    HSub "手順5　出来形数量を入れる"
    HP "スライド計算表の L列（数量・出来形）に、出来高検査の数量を入れます。" & vbLf & _
       "明細行だけでOK。階層行（工種・種別・細別）は自動で集計されます。" & _
       "M列（残工事）は ＝K−L で自動です。"
    HBlank

    HSub "手順6　調書の請負代金額を入れる"
    HP "スライド調書の C14（請負代金額・税込）だけが手入力です" & _
       "（設定シートの値が初期値で入っています）。" & vbLf & _
       "請負率(%)・請負工事価格（税抜）・消費税相当額は、ここから計算で出ます。"
    HBlank

    HSub "手順7　できあがり"
    HP "スライド調書の G16 がスライド額 S' です。E3に「増額スライド／減額スライド」が出ます。" & vbLf & _
       "調書は計算表からリンクしているので、出来形や〇を直すと調書も自動で追従します。"
    HBlank
End Sub


Private Sub WriteHelpInputs()
    HSec "3. 手入力するところ（黄色いセル）"

    HSub "スライド計算表・明細部"
    HHead3 "列", "項目", "入れるもの"
    HRow3 "L", "数量（出来形）", "必須。出来高検査の数量"
    HRow3 "O", "処分費〇", "処分費の行に〇。手順4で自動入力できます"
    HRow3 "P", "管材費〇", "水道工事の管材費の行に〇。手順4で自動入力できます"
    HRow3 "Q", "新工種〇", "基準日以降に追加された新工種の行に〇。" & _
                           "当初数量0・変更数量ありの明細には自動で入ります。" & _
                           "基準日より前の変更なら外してください"
    HRow3 "AA", "処分費額（前）", "全数量に対する処分費の総額。空欄ならその行の金額を全額 処分費として扱います"
    HRow3 "AB", "処分費額（後）", "空欄ならAA列の値を使います。処分費は据置が多いのでAA列だけで足ります"
    HP "R列は明細部では使いません。経費計算部のZコード項目の区分だけに使います。"
    HBlank

    HSub "スライド計算表・経費計算部（明細のすぐ下）"
    HHead2 "項目", "入れるもの"
    HRow "Zコード項目の区分（R列）", "自動で入ります。課の運用と違えば直してください（下記）"
    HRow "⑦⑪⑮ 経費率", "自動計算した値が入っています。積算システムの値があれば上書きしてください"
    HRow "⑯契約保証費", "全体の列（S・V・X）だけが手入力。出来高・残工事の列は自動です"
    HP "率を自前で持つのは S列（前・全体）・V列（後・全体）・X列（新工種抜き・全体）の3つだけです。" & _
       "T列はS列、W列はV列、Y列はX列の率を参照するので、" & _
       "積算システムで率を取る場合でも3回で済みます。"
    HBlank

    HSub "Zコード項目の区分（R列）"
    HHead3 "区分", "意味", "自動判定"
    HRow3 "積", "共通仮設費の積上分（運搬費・準備費・技術管理費など）。④に算入", "Z0001〜Z0039"
    HRow3 "原", "⑬工事原価に加算（工期延長等に伴う増加費用など）", "Z0040〜Z0044"
    HRow3 "価", "㉑工事価格に加算（設計委託費など）", "Z0045 以降"
    HRow3 "ス", "スクラップ。㉒に算入（1000円単位に切捨）", "名称に「スクラップ」を含む"
    HRow3 "空欄", "計算に使わない", "支給品費"
    HP "課の運用と違う場合はR列を直してください。全系列（S〜Y列）に一度に効きます。"
    HBlank

    HSub "スライド調書（様式4-2号）"
    HHead2 "セル", "項目"
    HRow "C14", "請負代金額（税込）。ここだけ手入力"
    HBlank
End Sub


Private Sub WriteHelpSheetMap()
    HSec "4. 画面の見方"

    HSub "スライド計算表の列"
    HCode "A        コード（非表示にしてOK）" & vbLf & _
          "B〜F     費目／工事区分／工種／種別／細別　①②③④のマーカーで階層を表現" & vbLf & _
          "G,H      施工単価名称／規格" & vbLf & _
          "I,J      単価（スライド前／スライド後）　← 旧単価CSV／新単価CSV" & vbLf & _
          "K,L,M    数量（全体／出来形★／残工事）" & vbLf & _
          "N        単位" & vbLf & _
          "O,P,Q    処分費〇★／管材費〇★／新工種〇☆" & vbLf & _
          "R        区分（経費計算部のZコード項目でだけ使う）★" & vbLf & _
          "S,T,U    スライド前（全体／出来形／残工事・検算用）" & vbLf & _
          "V,W      スライド後（全体／残工事）" & vbLf & _
          "X,Y      新工種抜き（全体／出来形）" & vbLf & _
          "Z        摘要" & vbLf & _
          "AA,AB    処分費額（全数量分・前★／後★）" & vbLf & _
          "AC〜AI   処分費額（7系列）" & vbLf & vbLf & _
          "★＝手入力　☆＝自動（修正可）"
    HP "明細部は8行目から。Zコードより前の行がすべて直接工事費です。" & _
       "そのすぐ下が経費計算部で、①〜③ → Zコード項目を1行ずつ → ④〜㉔ を7系列で計算します。" & vbLf & _
       "背景色は階層のレベルです。濃い青＝費目、薄い青＝工事区分、以下同様。"
    HBlank

    HSub "7系列の意味"
    HHead4 "列", "内容", "単価", "経費率"
    HRow4 "S", "スライド前・全体（＝元設計）", "旧", "自前"
    HRow4 "T", "スライド前・出来形（＝出来高）", "旧", "S列の率"
    HRow4 "U", "スライド前・残工事（検算用。S＝T＋U の確認にだけ使います）", "旧", "素材のみ"
    HRow4 "V", "スライド後・全体（＝スライド設計書）＝T＋W", "旧＋新", "自前"
    HRow4 "W", "スライド後・残工事", "新", "V列の率"
    HRow4 "X", "新工種抜き・全体", "旧", "自前"
    HRow4 "Y", "新工種抜き・出来形", "旧", "X列の率"
    HP "「新工種を考慮する」を「しない」にすると X・Y列は出ません（S〜Wの5系列）。" & vbLf & _
       "様式4-2号が使うのは S・T・W・X・Y の5列です。" & vbLf & _
       "U列はどこからも使われません。経費計算部でも「素材のみ」で④まで、他の系列の率も参照しません。" & vbLf & _
       "いまどちらのモードかは、スライド計算表の K1 セルとスライド調書の G4 セルに出ます" & _
       "（考慮しないときは赤字）。"
    HBlank

    HSub "スライド調書（様式4-2号）"
    HCode "P1'  ＝ 請負工事価格 － 出来高（請負率考慮）　　　　　 … 残工事" & vbLf & _
          "P1'' ＝ 新工種抜き　 － 新工種抜き出来高（請負率考慮） … 1%の母数" & vbLf & _
          "S'   ＝ P2' － P1' － P1'' × 1%　（増額スライド）" & vbLf & _
          "     ＝ P2' － P1' ＋ P1'' × 1%　（減額スライド）" & vbLf & vbLf & _
          "※「新工種を考慮する」を「しない」にした場合は P1'' ＝ P1' になります。"
    HBlank
End Sub


Private Sub WriteHelpRules()
    HSec "5. 計算のきまり"

    HSub "金額は1円未満切り捨て"
    HP "単価×数量 はすべて ROUNDDOWN(…,0) です。処分費額を数量比で割り振るときも同じです。"
    HBlank

    HSub "スライド後は「出来形＝旧単価 ＋ 残工事＝新単価」"
    HP "V列（スライド後・全体）は「全数量×新単価」ではありません。" & _
       "出来高はもう払い終わっているので旧単価のまま、残工事だけ新単価にします。" & _
       "1行の中で出来形が中途半端な工種は、その行の中で旧単価と新単価の両方を使います。"
    HCode "単価 1,000→1,100　全数量 10　出来形 4　残工事 6" & vbLf & vbLf & _
          "  T 出来形 ＝ 4 × 1,000 ＝ 4,000    （旧単価。もう払い終わっている）" & vbLf & _
          "  W 残工事 ＝ 6 × 1,100 ＝ 6,600    （新単価）" & vbLf & _
          "  V 全体   ＝ 4,000 ＋ 6,600 ＝ 10,600"
    HP "全数量×新単価だと 10×1,100＝11,000 になり、" & _
       "すでに払い終わった4単位分の値上がり（4×100＝400円）まで乗ってしまいます。" & vbLf & _
       "AF列（処分費額 後・全体）も同じ考え方で AD＋AG です。"
    HBlank

    HSub "③率計算の対象外費は按分"
    HCode "全体　　　　　 ＝ MAX(0, ② － ROUNDDOWN((① － ①'/2) × 3%, 0))" & vbLf & _
          "出来高・残工事 ＝ ROUNDDOWN(全体の③ × その系列の② ÷ 全体の②, 0)"
    HP "旧単価の系列は旧単価の全体から、新単価の系列は新単価の全体から按分します（混ざりません）。" & _
       "系列ごとに算式を当てると出来高＋残工事が全体に一致しないので、按分にしています。" & vbLf & _
       "設定シートの「③の出来高・残工事」を「系列ごとに計算」にすると従来の挙動に戻ります。"
    HBlank

    HSub "⑯契約保証費は出来高の側に乗せる"
    HP "契約保証費は請負契約に1回だけ付くので、出来高に乗せ、残工事は0にします。" & _
       "出来高が0（まだ何も出来ていない）ときだけ、残工事に乗せます。"
    HHead3 "系列", "出来高がある場合", "出来高が0の場合"
    HRow3 "S・V・X（全体）※手入力", "契約保証費", "契約保証費"
    HRow3 "T・Y（出来高）", "契約保証費", "0"
    HRow3 "W（残工事）", "0", "契約保証費"
    HP "こうしないとスライド額が契約保証費の分だけ過大になります。" & _
       "P1' は「全体－出来高」の引き算なので両方に入っていると打ち消えますが、" & _
       "P2' は残工事そのものなので残工事に入ったままになるためです。"
    HBlank

    HSub "旧CSVと新CSVの突合"
    HP "2本のCSVの行がそろっていないときは、上から順に試します。"
    HHead3 "順", "突合のしかた", "J列の印"
    HRow3 "0", "全体の行数も並びも同じ → 行の位置", "なし"
    HRow3 "A", "部ごと（単価表部／内訳部）に行数が同じ → その部の行の位置", "なし"
    HRow3 "B", "コード＋名称＋規格＋単位", "なし"
    HRow3 "C", "コード＋名称＋単位", "黄"
    HRow3 "D", "コード＋名称", "黄"
    HRow3 "E", "コードだけ", "黄"
    HRow3 "－", "どれも当たらない → 旧単価を据置", "赤"
    HRow3 "－", "新単価が0・旧単価が0でない → 旧単価を据置", "赤"
    HP "値上がり0で正しく同じ場合は印が付きません。" & _
       "「J列がI列と同じ」を見たとき、正常なのか突合できなかったのか区別できます。"
    HBlank
End Sub


Private Sub WriteHelpConfig()
    HSec "6. 設定シートの項目"

    HSub "【工事情報】"
    HHead3 "項目", "既定", "意味"
    HRow3 "工事名／工事場所／工期（自）（至）", "（空）", "任意"
    HRow3 "請負代金額", "0", "税込。様式4-2号の請負率の算出に使います"
    HRow3 "基準日", "（空）", "インフレスライドの請求日"
    HRow3 "発注者／受注者", "神戸市／（空）", "任意"
    HBlank

    HSub "【CSV設定】"
    HHead3 "項目", "既定", "意味"
    HRow3 "旧単価CSVパス／新単価CSVパス", "（空）", "空欄なら実行時にダイアログで選択"
    HRow3 "文字コード", "Shift_JIS", "文字化けしたら UTF-8"
    HRow3 "区切り文字", ",", ""
    HRow3 "使用系列", "変更", "変更／当初。CSVの①列＝変更設計、②列＝当初設計"
    HRow3 "新工種を考慮する", "する", "しないにすると新工種の列・系列を出さず、1%の母数は残工事そのもの"
    HRow3 "新工種の自動判定", "する", "当初数量0・変更数量ありの明細にQ列の〇を自動で付ける"
    HBlank

    HSub "【特殊集計区分CSVの設定】"
    HHead3 "項目", "既定", "意味"
    HRow3 "特殊集計区分CSVパス（スライド前）（スライド後）", "（空）", "空欄なら実行時に選択"
    HRow3 "特殊集計CSV列指定", "自動", "列の並びを手で決めたいとき"
    HRow3 "処分費の区分名", "処分", "この文字を含む区分をO列の〇に。カンマ区切りで複数可"
    HRow3 "管材費の区分名", "水道,管材,管財", "同じくP列の〇に"
    HRow3 "特殊集計CSVから処分費額も入れる", "する", "しないにすると〇だけ入れて金額は入れません"
    HBlank

    HSub "【経費計算の設定】"
    HHead3 "項目", "既定", "意味"
    HRow3 "契約保証費", "0", "当初設計時の額。出来高に乗り、残工事は0"
    HRow3 "処分費控除率", "0.03", "③の全体＝② － ROUNDDOWN((① － ①'/2)×この率, 0)"
    HRow3 "③の出来高・残工事", "按分", "按分／系列ごとに計算"
    HRow3 "消費税率", "0.1", ""
    HBlank

    HSub "【経費率の自動計算】"
    HHead3 "項目", "既定", "算式"
    HRow3 "共通仮設費率A／B", "485.4／-0.2231", "率(%) ＝ A × 対象額 ^ B"
    HRow3 "共通仮設費 地域補正／週休補正", "1.5／1.01", ""
    HRow3 "現場管理費率A／B", "202.3／-0.1034", "率(%) ＝ A × 対象額 ^ B"
    HRow3 "現場管理費 地域補正／週休補正", "1.2／1.02", ""
    HRow3 "一般管理費率係数／定数", "-4.97802／56.92101", "率(%) ＝ 係数 × LOG10(対象額) ＋ 定数"
    HBlank
    HP "古い設定シートを使い回す場合、［スライド計算表作成］を実行すると" & _
       "足りない設定が赤字で自動追加されます。確認してからもう一度実行してください。"
    HBlank
End Sub


Private Sub WriteHelpMacros()
    HSec "7. 実行できるマクロ（Alt+F8 で出る一覧）"
    HHead2 "マクロ", "内容"
    HRow "設定シート作成", "「設定」シートを作る（この使い方シートも無ければ作ります）"
    HRow "スライド計算表作成", "計算表と調書を作る（メイン）"
    HRow "特殊集計区分CSV取込", "処分費〇・管材費〇・処分費額を自動入力"
    HRow "特殊集計区分CSV構造チェック", "特殊集計区分CSVの列の並びを金額抜きで書き出す"
    HRow "CSV構造チェック", "スライド用csvの階層判定を金額抜きで書き出す"
    HRow "テスト用CSV作成", "動作確認用のサンプルCSVを2本作る"
    HRow "使い方シート作成", "このシートを作り直す"
    HP "CSV構造チェックと特殊集計区分CSV構造チェックは金額・数量を出しません。" & _
       "結果シートをそのままコピーして送れます。"
    HBlank
End Sub


Private Sub WriteHelpTrouble()
    HSec "8. やり直したいとき"
    HWarn "［スライド計算表作成］を実行すると、シートは作り直しになります。" & _
          "手入力した出来形や〇は消えます。やり直す前にシートをコピーして退避してください。"
    HBlank

    HSec "9. 困ったとき"
    HHead2 "症状", "原因と対処"
    HRow "マクロが一覧に出てこない", "貼り付け先が「標準モジュール」か確認。.xlsm で保存しているかも確認"
    HRow "文字化けする", "設定シートの「文字コード」を UTF-8 にして再実行"
    HRow "「明細を1行も読み取れませんでした」", _
         "CSVが「スライド用csv出力」の形式か確認。CSV構造チェックの結果を送ってください"
    HRow "階層がおかしい", "CSV構造チェックを実行して結果を送ってください"
    HRow "J列がI列と同じ（色なし）", "正常です。突合できていて、単価が動かなかっただけです"
    HRow "J列が赤い＋摘要に★新単価なし／★新単価が0", _
         "新CSVにその行が見つからなかった（または0だった）ので旧単価を据え置いています。" & _
         "金入設計書と見比べてください"
    HRow "J列が黄色＋摘要に★新単価は…で突合", _
         "規格や単位が2本のCSVで一致せず、ゆるい条件で当てた行です。" & _
         "どの条件で当てたか摘要に出ます"
    HRow "「旧CSVと新CSVで行がずれていた」と出る", _
         "2本のCSVで行がそろっていません。段階的に突合し直すので多くは問題ありませんが、" & _
         "★の件数が多いときは要確認"
    HRow "仮設材運搬費・技術管理費が積上分に入らない", _
         "経費計算部のZコード項目の行で、R列に「積」と入れてください"
    HRow "共通仮設費が大きすぎる", _
         "Z0040以降に「積」が付いていないか確認。付いていると二重計上になります"
    HRow "直接工事費に仮設工が入っていない", _
         "Zコードより前の行はすべて直接工事費に入ります。" & _
         "入っていなければCSV構造チェックの結果を送ってください"
    HRow "経費率が課の値と違う", _
         "経費計算部の⑦⑪⑮を直接上書きするか、設定シートの係数を変えてください"
    HRow "工事費が金入設計書と合わない", _
         "経費率を積算システムの値で上書き。" & _
         "それでも合わなければ⑯契約保証費・㉒スクラップ・端数整理を確認"
    HRow "③の出来高・残工事が元の経費計算シートと違う", _
         "既定を按分にしています。元に戻すには設定シートの「③の出来高・残工事」を" & _
         "「系列ごとに計算」に"
    HRow "特殊集計区分CSVが1行も取り込めない", _
         "メッセージに読み取った列の並びが出ます。★が付いていればその列を読めていないので、" & _
         "「特殊集計CSV列指定」に列を書いてください"
    HRow "特殊集計区分CSVの取りこぼしが多い", _
         "［スライド計算表作成］をやり直すと _単価表 シートが作られます。" & _
         "古いブックには入っていません"
    HRow "新工種なし版を貼ったのに新工種の列が出る", _
         "古い設定シートに項目が無く既定に戻っていました。" & _
         "［スライド計算表作成］で設定が追加されるので、確認してもう一度実行してください"
    HBlank
End Sub


'==============================================================
' 書き出しの部品
'==============================================================
Private Sub HTitle(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = s
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
End Sub


Private Sub HNote(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = s
        .Font.Color = RGB(120, 120, 120)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    FitPara mHws.Rows(mHr), s
End Sub


Private Sub HSec(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = s
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = vbWhite
        .Interior.Color = RGB(47, 117, 181)
        .VerticalAlignment = xlCenter
        .RowHeight = 24
    End With
End Sub


Private Sub HSub(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = s
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
End Sub


Private Sub HP(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = s
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    FitPara mHws.Rows(mHr), s
End Sub


Private Sub HWarn(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = "★ " & s
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Color = RGB(192, 0, 0)
        .Interior.Color = RGB(255, 242, 204)
    End With
    FitPara mHws.Rows(mHr), s & "　★"
End Sub


Private Sub HCode(ByVal s As String)
    mHr = mHr + 1
    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .Merge
        .Value = s
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Name = "ＭＳ ゴシック"
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    FitPara mHws.Rows(mHr), s
End Sub


Private Sub HBlank()
    mHr = mHr + 1
    mHws.Rows(mHr).RowHeight = 7
End Sub


Private Sub HHead2(ByVal a As String, ByVal b As String)
    HRowCore a, b, "", "", 2, True
End Sub

Private Sub HHead3(ByVal a As String, ByVal b As String, ByVal c As String)
    HRowCore a, b, c, "", 3, True
End Sub

Private Sub HHead4(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String)
    HRowCore a, b, c, d, 4, True
End Sub

Private Sub HRow(ByVal a As String, ByVal b As String)
    HRowCore a, b, "", "", 2, False
End Sub

Private Sub HRow3(ByVal a As String, ByVal b As String, ByVal c As String)
    HRowCore a, b, c, "", 3, False
End Sub

Private Sub HRow4(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String)
    HRowCore a, b, c, d, 4, False
End Sub


Private Sub HRowCore(ByVal a As String, ByVal b As String, ByVal c As String, ByVal d As String, _
                     ByVal nCol As Long, ByVal isHead As Boolean)
    mHr = mHr + 1
    mHws.Cells(mHr, 1).Value = a

    Select Case nCol
        Case 2
            mHws.Range(mHws.Cells(mHr, 2), mHws.Cells(mHr, 4)).Merge
            mHws.Cells(mHr, 2).Value = b
        Case 3
            mHws.Cells(mHr, 2).Value = b
            mHws.Range(mHws.Cells(mHr, 3), mHws.Cells(mHr, 4)).Merge
            mHws.Cells(mHr, 3).Value = c
        Case Else
            mHws.Cells(mHr, 2).Value = b
            mHws.Cells(mHr, 3).Value = c
            mHws.Cells(mHr, 4).Value = d
    End Select

    With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
        .WrapText = True
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
    End With

    If isHead Then
        With mHws.Range(mHws.Cells(mHr, 1), mHws.Cells(mHr, 4))
            .Font.Bold = True
            .Interior.Color = CLR_HEAD
            .HorizontalAlignment = xlCenter
        End With
    End If

    ' 各セルの折り返し行数のうち、いちばん多いものに合わせる
    FitRowCells mHr, a, b, c, d, nCol
End Sub


Private Sub FitRowCells(ByVal rw As Long, ByVal a As String, ByVal b As String, _
                        ByVal c As String, ByVal d As String, ByVal nCol As Long)
    Dim n As Long
    n = WrapLines(a, H_C1 - 2)
    Select Case nCol
        Case 2
            n = MaxL(n, WrapLines(b, H_C2 + H_C3 + H_C4 - 2))
        Case 3
            n = MaxL(n, WrapLines(b, H_C2 - 2))
            n = MaxL(n, WrapLines(c, H_C3 + H_C4 - 2))
        Case Else
            n = MaxL(n, WrapLines(b, H_C2 - 2))
            n = MaxL(n, WrapLines(c, H_C3 - 2))
            n = MaxL(n, WrapLines(d, H_C4 - 2))
    End Select
    mHws.Rows(rw).RowHeight = 14.5 * n + 3
End Sub


Private Function MaxL(ByVal x As Long, ByVal y As Long) As Long
    If y > x Then MaxL = y Else MaxL = x
End Function


'--------------------------------------------------------------
' 行の高さを、折り返し行数から決める
'   結合セルは AutoFit が効かないので自分で計算する
'--------------------------------------------------------------
Private Sub FitPara(ByVal rw As Range, ByVal s As String)
    rw.RowHeight = 14.5 * WrapLines(s, H_WRAP) + 3
End Sub


Private Function WrapLines(ByVal s As String, ByVal w As Long) As Long
    Dim a As Variant, p As Variant
    Dim n As Long
    If w < 4 Then w = 4
    a = Split(s, vbLf)
    For Each p In a
        n = n + Int(DispLen(CStr(p)) / w) + 1
    Next p
    If n < 1 Then n = 1
    WrapLines = n
End Function


' 全角を2・半角を1として数えた表示幅
Private Function DispLen(ByVal s As String) As Long
    Dim i As Long, n As Long, ch As Long
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        If ch < 0 Or ch > 255 Then n = n + 2 Else n = n + 1
    Next i
    DispLen = n
End Function


'--------------------------------------------------------------
' 仕上げ
'--------------------------------------------------------------
Private Sub FinishHelpSheet()
    mHws.Columns(1).ColumnWidth = H_C1
    mHws.Columns(2).ColumnWidth = H_C2
    mHws.Columns(3).ColumnWidth = H_C3
    mHws.Columns(4).ColumnWidth = H_C4
    mHws.Columns(5).ColumnWidth = 2

    On Error Resume Next
    mHws.Move Before:=ThisWorkbook.Worksheets(1)
    With mHws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ""
    End With
    On Error GoTo 0
End Sub
