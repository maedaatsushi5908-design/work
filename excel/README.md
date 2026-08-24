# Excel 様式 VBA 作業フォルダ

業務で使う Excel 様式を VBA で自動化するための作業場所です。

## フォルダ構成

| フォルダ | 用途 | 置くもの |
|---|---|---|
| `original/` | **あなたがアップロードする場所** | 元の Excel 様式（`.xlsx` / `.xlsm` / `.xltx` / `.xltm`） |
| `vba/` | Claude が生成する VBA コード | `.bas`（標準モジュール）、`.cls`（クラス）、`.frm`（ユーザーフォーム） |
| `docs/` | 仕様メモ・設計メモ | 解析結果、やりたいことのメモ（Markdown） |

## アップロードの手順（GitHub の画面から）

1. 下のリンクを開く（作業ブランチの `excel/original/` に直接アップロードできます）

   https://github.com/maedaatsushi5908-design/work/upload/claude/excel-vba-customization-uacsmu/excel/original

2. Excel ファイルをドラッグ＆ドロップ
3. 画面下部で **「Commit directly to the `claude/excel-vba-customization-uacsmu` branch」** が選ばれていることを確認
4. **Commit changes** をクリック

### コマンドラインから入れる場合

```bash
git checkout claude/excel-vba-customization-uacsmu
git pull origin claude/excel-vba-customization-uacsmu
cp ~/Desktop/様式.xlsm excel/original/
git add excel/original/様式.xlsm
git commit -m "Excel様式を追加"
git push -u origin claude/excel-vba-customization-uacsmu
```

## アップロード時に一緒に伝えてほしいこと

VBA の作り込みに直結するので、分かる範囲で教えてください。

- **何を自動化したいか** — 例：別シートの一覧から様式へ転記、複数ファイルへ一括出力、入力チェック、PDF 出力、印刷範囲の自動調整
- **入力元のデータ** — 手入力か、別ファイルからか、システムからの CSV か
- **実行のきっかけ** — ボタンを押す／ファイルを開いたら自動／シート変更時に自動
- **Excel のバージョンと bit 数** — 例：Microsoft 365（64bit）、Excel 2019（32bit）
  ※ Windows API を使う場合、32/64bit で書き方が変わります
- **シート保護・ブック保護の有無**（かかっている場合はパスワードの要否）
- **ファイルを配る相手がいるか** — 他の人も使うなら、マクロ有効化の案内やエラー処理を厚くします

## 機密情報について

社外秘のデータや個人情報が入っている場合は、**様式（レイアウト・項目名・数式）だけ残して中身のデータを消したサンプル**をアップロードしてください。それでも十分に VBA は書けます。

## 注意事項

- `.xlsx` にはマクロを保存できません。VBA を組み込む場合は **`.xlsm`（マクロ有効ブック）** で保存し直す必要があります。
- この環境には Excel がないため、**VBA の実行テストはお手元でお願いします**。こちらでは構文とロジックを詰めた状態でお渡しします。
- Excel を開いたまま `git add` すると、一時ファイル（`~$様式.xlsm`）が混ざることがあります。`.gitignore` で除外済みですが、**閉じてからコミット**するのが安全です。
