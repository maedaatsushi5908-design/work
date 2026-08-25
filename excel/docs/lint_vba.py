#!/usr/bin/env python3
"""VBA モジュールの簡易チェック

作成環境に Excel が無くコンパイルできないため、機械的に潰せるものだけ潰す。

  1. 予約語を変数名・引数名・Type メンバ名に使っていないか
     （`Dim any As Boolean` は VBA では構文エラー）
  2. Sub / Function と End Sub / End Function の対応
  3. 全角スペースがコード行に混ざっていないか

    python3 excel/docs/lint_vba.py excel/vba/*.bas
"""
import re
import sys

RESERVED = {w.lower() for w in """
Any As Boolean ByRef ByVal Byte Call Case Close Const Currency Date Declare
Dim Do Double Each Else ElseIf Empty End Enum Eqv Erase Error Event Exit False
For Friend Function Get Global GoSub GoTo If Imp Implements In Input Integer
Is Len Let Like Line Load Lock Long Loop LSet Me Mod New Next Not Nothing Null
Object On Open Option Optional Or ParamArray Preserve Print Private Property
Public Put RaiseEvent ReDim Rem Resume Return RSet Seek Select Set Single
Static Step Stop String Sub Then To True Type TypeOf Unload Until Variant Wend
While With WithEvents Write Xor Name Time Width Circle Scale Spc Tab Erl Err
""".split()}

def strip_literals(line):
    """文字列リテラルの中身を伏字にし、行コメント以降を落とす"""
    out, i, in_str = [], 0, False
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    out.append("..")
                    i += 2
                    continue
                in_str = False
                out.append('"')
            else:
                out.append(".")
        elif ch == '"':
            in_str = True
            out.append('"')
        elif ch == "'":
            break  # 行コメント
        else:
            out.append(ch)
        i += 1
    return "".join(out).strip()


VAR_AS = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\s+As\s+", re.I)
FOR_EACH = re.compile(r"\bFor\s+Each\s+([A-Za-z_][A-Za-z0-9_]*)\b", re.I)
OPENER = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property)\b", re.I)
CLOSER = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)

# モジュールレベルの宣言。VBA ではすべての手続きより前に置く決まりがあり、
# 後ろに来るとコンパイルできない。モジュールを連結したときに起きやすい。
MODULE_DECL = re.compile(
    r"^(?:Public\s+|Private\s+|Global\s+)?(?:Const|Type|Enum|Dim|Declare)\s", re.I
)
MODULE_VAR = re.compile(
    r"^(?:Public|Private)\s+(?:WithEvents\s+)?[A-Za-z_][A-Za-z0-9_]*\s+As\s", re.I
)


def check(path):
    problems = []
    depth = 0
    seen_proc = False
    for ln, raw in enumerate(open(path, encoding="utf-8"), 1):
        line = raw.rstrip("\n")
        stripped = line.strip()
        if stripped.startswith("'"):
            continue

        # 文字列リテラルを潰してから調べる。シート名の "　印刷　" などを
        # 全角スペースとして誤検出しないため。VBA の "" は文字列中の " 1個。
        code = strip_literals(stripped)

        for name in VAR_AS.findall(code):
            if name.lower() in RESERVED and name.lower() != "as":
                problems.append((ln, f"予約語を名前に使用: {name}", code[:78]))
        for name in FOR_EACH.findall(code):
            if name.lower() in RESERVED:
                problems.append((ln, f"For Each の変数が予約語: {name}", code[:78]))

        if "　" in code:
            problems.append((ln, "コード行に全角スペース", code[:78]))

        # 手続きの外に出ている宣言が、最初の手続きより後ろに来ていないか
        if depth == 0 and (MODULE_DECL.match(code) or MODULE_VAR.match(code)):
            if seen_proc:
                problems.append(
                    (ln, "モジュールレベルの宣言が手続きより後ろにある"
                         "（VBA では全ての手続きより前に置く必要がある）", code[:78])
                )

        if OPENER.match(code) and not re.search(r"\bDeclare\b", code, re.I):
            depth += 1
            seen_proc = True
        elif CLOSER.match(code):
            depth -= 1

    if depth != 0:
        problems.append((0, f"Sub/Function の開始と End の数が合わない（差 {depth}）", ""))
    return problems


def main(paths):
    total = 0
    for p in paths:
        probs = check(p)
        total += len(probs)
        if probs:
            print(f"--- {p}")
            for ln, msg, code in probs:
                where = f"{ln:>5}" if ln else "    -"
                print(f"  {where}  {msg}")
                if code:
                    print(f"         {code}")
        else:
            print(f"--- {p}  問題なし")
    print(f"\n指摘 {total} 件")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:] or ["excel/vba/M_Util.bas"]))
