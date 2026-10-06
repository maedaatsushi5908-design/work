import io, re
MODS = [('M00_Main.bas','メイン処理・CSV読み込み・2本の突合'),
        ('M10_CsvIO.bas','CSV読み書き・数値／文字列変換'),
        ('M20_Slide.bas','設定シート生成・スライド計算表（明細＋経費計算）'),
        ('M40_Chosho.bas','スライド調書（様式4-2号）'),
        ('M90_SampleData.bas','動作確認用サンプルCSV生成'),
        ('M92_CsvCheck.bas','CSVの構造チェック（診断用）')]
PROC = re.compile(r'^\s*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)\s')

def build(outfile, useShin):
    decls, bodies = [], []
    for fn, desc in MODS:
        lines = io.open(fn, encoding='utf-8').read().split('\n')
        if not useShin:
            lines = [l.replace('DefaultUseShin = "する"', 'DefaultUseShin = "しない"') for l in lines]
        cut = next(i for i,l in enumerate(lines) if PROC.match(l))
        while cut > 0 and (lines[cut-1].strip().startswith("'") or lines[cut-1].strip()==''): cut -= 1
        head = [l for l in lines[:cut] if l.strip() != 'Option Explicit']
        while head and (head[0].strip()=='' or head[0].strip().startswith("'")): head.pop(0)
        while head and head[-1].strip()=='': head.pop()
        if head:
            decls.append(f"'--- {fn[:-4]} の宣言 ---"); decls += head; decls.append('')
        bodies += ['', "'"+'='*62, f"' 《{fn[:-4]}》 {desc}", "'"+'='*62] + lines[cut:]
    title = ("' インフレスライド計算表 作成マクロ  ―  全部入り1ファイル版" if useShin
             else "' インフレスライド計算表 作成マクロ  ―  全部入り1ファイル版（新工種を考慮しない）")
    shin = (["' 新工種を考慮します（設定シートの「新工種を考慮する」＝する）。",
             "' 考慮しない場合はその設定を「しない」に変えるか、新工種なし版を貼ってください。"] if useShin else
            ["' 新工種を考慮しません（設定シートの「新工種を考慮する」＝しない）。",
             "' 新工種の列・系列は出さず、受注者負担1%の母数は残工事そのもの（P1''＝P1'）です。",
             "' 考慮したくなったらその設定を「する」に変えれば、同じファイルのまま切り替わります。"])
    out = ["'"+'='*62, title, "'"+'-'*62,
      "' このファイル1つを標準モジュールに貼り付ければ動きます。",
      "' モジュール名（オブジェクト名）は何でも構いません。", "'",
      "' 手順",
      "'   1) Excelを「マクロ有効ブック(.xlsm)」で保存する",
      "'   2) Alt+F11 → 挿入 → 標準モジュール",
      "'   3) このファイルを全部コピーして貼り付ける",
      "'   4) Alt+F8 →「設定シート作成」→「スライド計算表作成」", "'",
      "' 作られるシート",
      "'   ・スライド計算表（明細＋経費計算）",
      "'   ・スライド調書（様式4-2号）… スライド計算表からリンク", "'"] + shin + ["'",
      "' CSVの前提：①列＝変更設計、②列＝当初設計",
      "' 金額は 単価×数量 を1円未満切り捨て（ROUNDDOWN）で計算します。", "'",
      "' そのほかのマクロ",
      "'   ・CSV構造チェック … CSVのコードと名称だけを書き出す（金額は出しません）",
      "'   ・テスト用CSV作成 … 動作確認用のサンプルCSVを2本作る", "'",
      "' 参照設定の追加は不要です（すべて CreateObject の遅延バインディング）。",
      "'"+'='*62, 'Option Explicit', '',
      "'"+'='*62, "' 宣言部", "'"+'='*62] + decls + bodies
    txt='\n'.join(out)
    while '\n\n\n\n' in txt: txt=txt.replace('\n\n\n\n','\n\n\n')
    io.open(outfile,'w',encoding='utf-8').write(txt)
    lines=txt.split('\n')
    P=re.compile(r'^(Public |Private |Friend )?(Static )?(Sub|Function)\s', re.I)
    D=re.compile(r'^(Public|Private)\s+(?!Sub|Function)', re.I)
    def st(l):
        o,q=[],False
        for ch in l:
            if ch=='"': q=not q; o.append(ch)
            elif ch=="'" and not q: break
            else: o.append(ch)
        return ''.join(o).strip()
    fp=next(i for i,l in enumerate(lines) if P.match(st(l)))
    d=0; bad=[]
    for i,l in enumerate(lines[fp:], start=fp):
        t=st(l)
        if P.match(t): d+=1
        elif t.lower() in ('end sub','end function'): d-=1
        elif d==0 and D.match(t): bad.append((i+1,t))
    names={}
    for i,l in enumerate(lines,1):
        m=re.match(r'^\s*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)\s+([^\s(]+)', l)
        if m: names.setdefault(m.group(1),[]).append(i)
    dup={k:v for k,v in names.items() if len(v)>1}
    print(f'{outfile}: {len(lines)} 行 / 手続き外宣言 {bad or "なし"} / 重複名 {dup or "なし"}')

build('ALL_IN_ONE.bas', True)
build('ALL_IN_ONE_新工種なし.bas', False)
