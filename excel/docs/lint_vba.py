#!/usr/bin/env python3
"""VBA モジュールの簡易チェック

作成環境に Excel が無くコンパイルできないため、機械的に潰せるものだけ潰す。

  1. 予約語を変数名・引数名・Type メンバ名に使っていないか
     （`Dim any As Boolean` は VBA では構文エラー）
  2. Sub / Function と End Sub / End Function の対応
  3. 手続き名と同じ名前の変数・引数を使っていないか
     （`Private Function LastRow()` があるのに `Dim lastRow As Long` と
      書くと、VBA は大文字小文字を区別しないので同じ名前になり壊れる）
  4. VBA / Excel の関数と同じ名前を変数に使っていないか
     （`Dim Val As String` と書くと、そのあと Val() を呼べなくなる）
  5. 全角スペースがコード行に混ざっていないか

    python3 excel/docs/lint_vba.py            # src と dist をまとめて
    python3 excel/docs/lint_vba.py excel/vba/src/M_Link.bas
"""
import glob
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

# 変数名に使うと、そのあとその関数を呼べなくなるもの。
# 予約語ではないので VBA は宣言を通すが、呼び出し側で壊れる。
RISKY = {w.lower() for w in """
Left Right Mid Val Split Replace InStr InStrRev Join Format Trim LTrim RTrim
UCase LCase Chr ChrW Asc AscW Abs Int Fix Round Sgn Sqr Array Now Timer Space
IsEmpty IsNull IsNumeric IsDate IsArray IsObject IsError TypeName VarType
CInt CLng CDbl CSng CStr CDate CBool CVar StrComp StrConv StrReverse Hex Oct
MsgBox InputBox RGB Environ Dir Shell Choose Switch IIf DateAdd DateDiff
Day Month Year Hour Minute Second Weekday
Range Cells Rows Columns Sheets Worksheets Workbooks ActiveCell ActiveSheet
ActiveWorkbook ThisWorkbook Selection Application Union Intersect Evaluate
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
PROC_NAME = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(?:Sub|Function|Property(?:\s+(?:Get|Let|Set))?)\s+([A-Za-z_][A-Za-z0-9_\u3000-\u9fff]*)",
    re.I,
)
CLOSER = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)

# モジュールレベルの宣言。VBA ではすべての手続きより前に置く決まりがあり、
# 後ろに来るとコンパイルできない。モジュールを連結したときに起きやすい。
MODULE_DECL = re.compile(
    r"^(?:Public\s+|Private\s+|Global\s+)?(?:Const|Type|Enum|Dim|Declare)\s", re.I
)
MODULE_VAR = re.compile(
    r"^(?:Public|Private)\s+(?:WithEvents\s+)?[A-Za-z_][A-Za-z0-9_]*\s+As\s", re.I
)


def proc_names(path):
    """モジュール内の手続き名を集める"""
    names = set()
    for raw in open(path, encoding="utf-8"):
        code = strip_literals(raw.strip())
        if re.search(r"\bDeclare\b", code, re.I):
            continue
        m = PROC_NAME.match(code)
        if m:
            names.add(m.group(1).lower())
    return names


def check(path):
    problems = []
    depth = 0
    seen_proc = False
    procs = proc_names(path)
    for ln, raw in enumerate(open(path, encoding="utf-8"), 1):
        line = raw.rstrip("\n")
        stripped = line.strip()
        if stripped.startswith("'"):
            continue

        # 文字列リテラルを潰してから調べる。シート名の "　印刷　" などを
        # 全角スペースとして誤検出しないため。VBA の "" は文字列中の " 1個。
        code = strip_literals(stripped)

        for name in VAR_AS.findall(code):
            # "as" も除外しない。VBA は大文字小文字を区別しないので
            # aS のような名前は予約語 As と衝突する。
            if name.lower() in RESERVED:
                problems.append((ln, f"予約語を名前に使用: {name}", code[:78]))
            # 手続き名と同じ名前の変数は、その手続きを呼べなくなる。
            # 宣言している行そのものは除く（Function LastRow() As Long）。
            elif name.lower() in procs and not PROC_NAME.match(code):
                problems.append(
                    (ln, f"手続きと同じ名前の変数: {name}"
                         "（VBA は大文字小文字を区別しないため衝突する）", code[:78])
                )
            elif name.lower() in RISKY:
                problems.append(
                    (ln, f"VBA / Excel の関数と同じ名前の変数: {name}"
                         "（そのあと {0}() を呼べなくなる）".format(name), code[:78])
                )
        for name in FOR_EACH.findall(code):
            if name.lower() in RESERVED:
                problems.append((ln, f"For Each の変数が予約語: {name}", code[:78]))
            elif name.lower() in procs:
                problems.append(
                    (ln, f"For Each の変数が手続きと同じ名前: {name}", code[:78])
                )

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
    sys.exit(main(sys.argv[1:] or sorted(
        glob.glob("excel/vba/src/*.bas") + glob.glob("excel/vba/dist/*.bas"))))
