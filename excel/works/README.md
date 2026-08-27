# 工事ごとの数量計算書

```
excel/works/
  01_higashishirakawa/   東白川特２高層配水池揚水管取替工事（令和7年度）
  02/                    2件目。給水・仮配のある工事
```

各フォルダの `files.md` に、元のファイル名との対応と工事の概要を書く。

## 新しい工事を入れるとき

### 1. ファイル名を半角にする

GitHub のアップローダーは、**全角スペース・拡張子前のスペース・`【】（）～`** を
含むファイル名で `Something went really wrong, and we can't process that file.`
を返す。アップロードする前に、手元でこの形に変えておく。

```
01_soukatsu.xlsx        01 総括表…
02_chutetsukan.xls      02 …鋳鉄管製造 数量計算書
03_kirikan.xlsx         03 …切管表
04_unpanhi.xls          04 …運搬費 数量計算書
05-1_dokou_konkyo.xlsx  05-1 …根拠土工模式図集計表
06_dokou_hosou.xlsx     06 土工事・舗装復旧 数量計算書   ← これがいちばん大事
06_hosou_shukei.xlsx    06 …舗装復旧面積他集計表
07_futaikou.xls         07 …付帯工
08_koukan.xls           08 …鋼管工事数量
```

名前は上と同じでなくてよい。**半角英数字・`_`・`-`・`.` だけ**にしてあればよい。
元の名前は `files.md` に控える。

### 2. アップロードする

<https://github.com/maedaatsushi5908-design/work> で
`excel/works/02` を開き、右上の **Add file → Upload files** にドラッグ＆ドロップ。
下の **Commit changes** を押す。

- `Uploads are disabled` と出るときは、GitHub に
  `maedaatsushi5908-design` でログインできていない
- ブランチは `claude/excel-vba-customization-uacsmu` を選ぶ

`.xls`（97-2003 形式）のままでよい。こちらで LibreOffice を使って
`.xlsx` に変換して読む。

### 3. 読み取る

```
python3 excel/docs/scan_work.py excel/works/02
```

総括表の見出し・シート名・舗装版破砕ブロックを突き合わせ、
`M_Hasai` の先頭に貼る `TARGET_SHEET` と `COL_MAP` の下書きを出す。

1件目の工事で流すと、手で確かめた `COL_MAP` と**完全に一致する**ことを確認済み。
候補が絞れない列には `★` が付くので、そこだけ見出しとシート名を見比べて決める。
