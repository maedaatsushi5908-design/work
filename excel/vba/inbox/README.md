# 手元で直したモジュールを戻す場所

VBE で直したものをここへ置くと、差分だけ見て `src` に取り込みます。
チャットにコードを貼るより、こちらのほうが速くて安く済みます。

## やり方

1. VBE の左ツリーで `M_Hasai` を右クリック → **ファイルのエクスポート**
2. `M_Hasai.bas` として保存（名前はそのままでよい）
3. このフォルダ（`excel/vba/inbox/`）へアップロード
   <https://github.com/maedaatsushi5908-design/work/tree/claude/excel-vba-customization-uacsmu/excel/vba/inbox>
4. チャットで「inbox に置きました」と一言

こちらで `diff` を取り、変わったところだけを `src/M_Hasai.bas` に入れて
`build_vba.py` で `dist` を作り直します。取り込んだらこのフォルダは空にします。

## 直した中身を覚えているなら、言葉だけで十分

「COL_MAP の O 列を消した」「TARGET_SHEET を〇〇にした」程度なら、
エクスポートも不要です。こちらで `src` を直します。

## 大事なこと

**`src/M_Hasai.bas` が正本です。** VBE で直しただけだと、次にこちらが
`dist` を作り直したときに消えます。残したい直しは必ず戻してください。

## 直してよいところ

`dist/M_Hasai.bas` の先頭、

```
' ここだけ工事に合わせて直す        ← 46行目あたり
   …
' ここから下は工事が変わっても直さない  ← 136行目あたり
```

に挟まれた定数だけが、工事ごとに変えることを想定した場所です。

| 定数 | 何を決めるか |
|---|---|
| `TARGET_SHEET` | 総括表のシート名 |
| `COL_MAP` | 総括表の列 → 転記元シート |
| `KARA_MAP` | 殻運搬の副見出し → 転記元ブロック |
| `SECTION_LABEL` / `BLOCK_LABEL` | 舗装版破砕の工種名とブロック見出し |
| `CUT_LABEL` / `CUT_BLOCK` | 舗装切断工の〃 |
| `TORI_BLOCK` / `TORI_KIND` | 給水2度の舗装版取壊の〃 |
| `INPUT_COLOR` | 入力セルの色（黄色） |

ここより下を直したときは、`lint_vba.py` に通してから戻してください。
