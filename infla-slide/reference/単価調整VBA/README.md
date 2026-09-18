# 単価調整Excel から抽出した元VBAソース

`infla-slide/input/単価調整Excel（VBA格納ファイル）20260817.xlsm` の `vbaProject.bin` から
抽出したVBAソースです。**GitHubへのアップロードでVBAは失われていません。**

| ファイル | 行数 | 内容 |
|---|---:|---|
| `Module2.bas` | 1855 | 本体。①〜④のメイン処理と補助マクロ一式 |
| `Module1.bas` / `Module3.bas` / `Module4.bas` | 1〜2 | 空 |
| `ThisWorkbook.bas` / `Sheet1.bas` / `Sheet11.bas` | 8 | 空（シートモジュール） |

冒頭の `Attribute VB_Name = "Module2"` の行は、VBEへコピペで戻すときは**削除**してください
（貼り付けるとエラーになります）。ファイルとして取り込む場合はそのままで構いません。

この中身を解析した結果は [`../../docs/CSV仕様_スライド用csv出力.md`](../../docs/CSV仕様_スライド用csv出力.md) にまとめてあります。
