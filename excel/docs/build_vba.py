#!/usr/bin/env python3
"""貼り付け用の1ファイルモジュールを組み立てる

保守は excel/vba/src/ の小さいモジュールで行い、配布は1ブック1ファイルにする。
同じ関数を2箇所で直さずに済ませるため。

    python3 excel/docs/build_vba.py

出力:
    excel/vba/dist/M_Hasai.bas      06 土工事・舗装復旧 用
    excel/vba/dist/M_Tenki.bas      01 総括表 用
    excel/vba/dist/sjis/*.bas       同じものの Shift-JIS 版（インポート用）
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "vba", "src")
DIST = os.path.join(ROOT, "vba", "dist")

BUNDLES = [
    {
        "name": "M_Hasai",
        # 単独で動くので M_Util は混ぜない。読む量を減らすため。
        "parts": ["M_Hasai.bas"],
        "header": [
            "' 06 土工事・舗装復旧 数量計算書 用",
            "'",
            "' このファイル1つだけを標準モジュールに貼り付ければ動く。",
            "' マクロは「舗装版破砕を転記する」1本。",
        ],
    },
    {
        "name": "M_Tenki",
        "parts": ["M_Util.bas", "M_Config.bas", "M_Engine.bas", "M_Main.bas"],
        # モジュールをまたぐために Public にしてあるが、
        # 結合後は同じモジュール内なので Private でよい。
        # Public のままだと Alt+F8 のマクロ一覧に出てしまう。
        "make_private": ["InitSources", "CloseSources"],
        "header": [
            "' 01 総括表 用",
            "'",
            "' このファイル1つだけを標準モジュールに貼り付ければ動く。",
            "' マクロは「定義シートを作成」「照合実行」「転記実行」。",
        ],
    },
]

ATTR = re.compile(r'^Attribute VB_Name = "[^"]*"\s*$')
OPT = re.compile(r"^Option Explicit\s*$")


PROC = re.compile(
    r"^(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function|Property)\s", re.I
)


def load(part, make_private=()):
    """1モジュールを読み、宣言部と手続き部に分けて返す

    VBA は Type / Const / Enum / モジュール変数を、すべての手続きより前に
    置かないとコンパイルできない。単純に連結すると2つ目以降のモジュールの
    宣言が手続きの後ろに来て構文エラーになるため、ここで分けておく。
    """
    with open(os.path.join(SRC, part), encoding="utf-8") as f:
        lines = f.read().split("\n")
    lines = [l for l in lines if not ATTR.match(l) and not OPT.match(l)]

    for i, l in enumerate(lines):
        if PROC.match(l):
            cut = i
            break
    else:
        cut = len(lines)

    # 手続きの直前にあるコメント・空行は、その手続きの説明なので手続き側へ回す
    while cut > 0:
        prev = lines[cut - 1].strip()
        if prev == "" or prev.startswith("'"):
            cut -= 1
        else:
            break

    decl = "\n".join(lines[:cut]).strip("\n")
    body = "\n".join(lines[cut:]).strip("\n")

    for name in make_private:
        body = re.sub(rf"^Public (Sub|Function) {re.escape(name)}\b",
                      rf"Private \1 {name}", body, flags=re.M)
    return decl, body


def build(bundle):
    out = [f'Attribute VB_Name = "{bundle["name"]}"']
    out.append("'" + "=" * 66)
    out.extend(bundle["header"])
    out.append("'")
    out.append("' 元は excel/vba/src/ の " + " / ".join(bundle["parts"]) + " を")
    out.append("' つなげたもの。直すときは src 側を直して build_vba.py を実行する。")
    out.append("'" + "=" * 66)
    out.append("Option Explicit")

    parts = [(p,) + load(p, bundle.get("make_private", ())) for p in bundle["parts"]]

    # 宣言はすべて先頭に集める（VBA の決まり）
    for name, decl, _body in parts:
        if not decl:
            continue
        out.append("")
        out.append("'--- " + name + " の宣言 " + "-" * max(0, 50 - len(name)))
        out.append(decl)

    # 手続きはそのあとに並べる
    for name, _decl, body in parts:
        if not body:
            continue
        out.append("")
        out.append("'" + "=" * 66)
        out.append(f"' ここから {name}")
        out.append("'" + "=" * 66)
        out.append(body)

    return "\n".join(out) + "\n"


def main():
    os.makedirs(os.path.join(DIST, "sjis"), exist_ok=True)
    for b in BUNDLES:
        text = build(b)
        path = os.path.join(DIST, b["name"] + ".bas")
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        sjis = os.path.join(DIST, "sjis", b["name"] + ".bas")
        with open(sjis, "w", encoding="cp932") as f:
            f.write(text)
        print(f"  {b['name']}.bas  {text.count(chr(10)) + 1} 行  ← " + " + ".join(b["parts"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
