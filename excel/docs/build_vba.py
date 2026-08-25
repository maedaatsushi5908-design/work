#!/usr/bin/env python3
"""貼り付け用の1ファイルモジュールを組み立てる

保守は excel/vba/src/ の小さいモジュールで行い、配布は1ブック1ファイルにする。
同じ関数を2箇所で直さずに済ませるため。

    python3 excel/docs/build_vba.py

出力:
    excel/vba/dist/M_Link.bas       06 土工事・舗装復旧 用
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
        "name": "M_Link",
        "parts": ["M_Util.bas", "M_Link.bas"],
        "header": [
            "' 06 土工事・舗装復旧 数量計算書 用",
            "'",
            "' このファイル1つだけを標準モジュールに貼り付ければ動く。",
            "' マクロは「総括表の数式を作り直す」1本。",
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


def load(part, make_private=()):
    """1モジュールを読み、Attribute と Option Explicit を落とす"""
    with open(os.path.join(SRC, part), encoding="utf-8") as f:
        lines = f.read().split("\n")
    text = "\n".join(l for l in lines if not ATTR.match(l) and not OPT.match(l)).strip("\n")
    for name in make_private:
        text = re.sub(rf"^Public (Sub|Function) {re.escape(name)}\b",
                      rf"Private \1 {name}", text, flags=re.M)
    return text


def build(bundle):
    out = [f'Attribute VB_Name = "{bundle["name"]}"']
    out.append("'" + "=" * 66)
    out.extend(bundle["header"])
    out.append("'")
    out.append("' 元は excel/vba/src/ の " + " / ".join(bundle["parts"]) + " を")
    out.append("' つなげたもの。直すときは src 側を直して build_vba.py を実行する。")
    out.append("'" + "=" * 66)
    out.append("Option Explicit")
    for p in bundle["parts"]:
        out.append("")
        out.append("'" + "=" * 66)
        out.append(f"' ここから {p}")
        out.append("'" + "=" * 66)
        out.append(load(p, bundle.get("make_private", ())))
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
