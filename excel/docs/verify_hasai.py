#!/usr/bin/env python3
"""M_Hasai が書き込む数式を、VBA と同じ手順で再現する。

作成環境に Excel が無いので、マクロの判定をそのまま Python に写して
「どのセルに何が入るか」を出す。VBA を直したらこちらも直すこと。

    python3 excel/docs/verify_hasai.py
"""
import collections
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as gl
from openpyxl.utils import column_index_from_string as ci

HERE = os.path.dirname(os.path.abspath(__file__))
BOOK = os.path.join(HERE, "..", "original", "06_dokou_hosou.xlsx")

# ---- M_Hasai の先頭にある設定と同じもの -----------------------------------
TARGET_SHEET = "総括表（土工事）"
COL_MAP = ("J=試掘（舗50|K=試掘（舗300|L=試掘（舗75|M=試掘（舗400|N=試掘（舗600|"
           "P=仮配（舗|Q=給水(舗|"
           "R=管工（舗50|S=管工（舗75|T=管工（舗400|U=管工（舗600")
SECTION_LABEL = "舗装版破砕"
BLOCK_LABEL = "□舗装版破砕"
INPUT_COLOR = "FFFF00"


def norm(v):
    """VBA の Norm と同じ。全角英数を半角に直し、空白を落として大文字に"""
    if v is None:
        return ""
    out = []
    for ch in str(v):
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            o -= 0xFEE0
        if o not in (32, 0x3000, 9, 10, 13):
            out.append(chr(o))
    return "".join(out).upper()


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def merged_value(ws, r, c):
    """VBA の MergeArea.Cells(1,1).Value と同じ"""
    for m in ws.merged_cells.ranges:
        if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
            return ws.cell(m.min_row, m.min_col).value
    return ws.cell(r, c).value


def merged_top(ws, r, c):
    for m in ws.merged_cells.ranges:
        if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
            return m.min_row, m.min_col
    return r, c


def is_input_cell(ws, r, c):
    """黄色く塗ってある入力セルか（VBA の Interior.Color = 65535 と同じ）"""
    f = ws.cell(r, c).fill
    if f is None or f.patternType is None:
        return False
    fg = f.fgColor
    if fg.type == "rgb":
        return isinstance(fg.rgb, str) and fg.rgb.endswith(INPUT_COLOR)
    if fg.type == "indexed":
        return fg.indexed == 13        # 既定パレットの黄色 = FFFF00
    return False


# ---- 転記元シートのブロック ------------------------------------------------
_cache = {}


def find_block(wb, sn):
    if sn in _cache:
        return _cache[sn]
    _cache[sn] = None
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]
    sect = 0
    for r in range(1, 61):
        for c in range(1, 61):
            if norm(BLOCK_LABEL) in norm(ws.cell(r, c).value):
                sect = r
                break
        if sect:
            break
    if not sect:
        return None

    hdr, kinds, totals = 0, [], []
    for r in range(sect + 1, sect + 4):
        kinds, totals = [], []
        for c in range(1, 61):
            v = norm(ws.cell(r, c).value)
            if v == norm("種別・舗装厚"):
                kinds.append(c)
            if v == norm("合 計"):
                totals.append(c)
        if kinds and totals:
            hdr = r
            break
    if not hdr:
        return None

    pairs = []
    for kc in kinds:
        nxt = [t for t in totals if t > kc]
        if nxt:
            pairs.append((kc + 1, min(nxt)))
    if not pairs:
        return None
    a_thk, a_sum = pairs[0]
    c_thk, c_sum = pairs[1] if len(pairs) > 1 else (0, 0)

    r0, r1 = hdr + 1, hdr
    for r in range(r0, r0 + 41):
        if norm(ws.cell(r, a_thk - 1).value) not in ("AS", "CO"):
            break
        r1 = r
    if r1 < r0:
        return None
    _cache[sn] = (r0, r1 + 1, a_thk, a_sum, c_thk, c_sum)   # 予備を1行
    return _cache[sn]


def rng(col, r0, r1):
    L = gl(col)
    return f"${L}${r0}:${L}${r1}"


def build_sumif(wb, sn, tname, thk_ref, kind):
    b = find_block(wb, sn)
    if not b:
        return "", f"{sn} に {BLOCK_LABEL} のブロックがありません"
    r0, r1, a_thk, a_sum, c_thk, c_sum = b
    crit = f"'{tname}'!{thk_ref}"
    q = f"'{sn}'!"
    out = ""
    if "As" in kind and a_sum:
        out = f"SUMIF({q}{rng(a_thk, r0, r1)},{crit},{q}{rng(a_sum, r0, r1)})"
    if "Co" in kind and c_sum:
        if out:
            out += "+"
        out += f"SUMIF({q}{rng(c_thk, r0, r1)},{crit},{q}{rng(c_sum, r0, r1)})"
    if not out:
        return "", f"{sn} に {kind} 側の欄がありません"
    note = ""
    if "Co" in kind and not c_sum:
        note = f"{sn} に Co 側の欄が無いため As だけを合計しました"
    return "=" + out, note


# ---- 総括表側の読み取り ----------------------------------------------------
def is_section_row(ws, r, first_col):
    for c in range(1, first_col):
        v = merged_value(ws, r, c)
        if v is None:
            continue
        t = str(v)
        if t.strip() and not t.startswith("="):
            return norm(SECTION_LABEL) in norm(t)
    return False


def thick_ref(ws, r, first_col):
    for c in range(first_col - 1, 0, -1):
        tr, tc = merged_top(ws, r, c)
        if is_num(ws.cell(tr, tc).value):
            return f"${gl(tc)}{tr}"
    return ""


def kind_of_row(ws, r, first_col):
    out = ""
    for c in range(1, first_col):
        t = norm(merged_value(ws, r, c))
        if "AS" in t and "As" not in out:
            out += "As"
        if "CO" in t and "Co" not in out:
            out += "Co"
    return out


def main():
    wb = openpyxl.load_workbook(BOOK)
    if TARGET_SHEET not in wb.sheetnames:
        print("シートがありません:", TARGET_SHEET)
        return 1
    ws = wb[TARGET_SHEET]

    pairs = [p.split("=", 1) for p in COL_MAP.split("|") if "=" in p]
    first_col = min(ci(c) for c, _ in pairs)

    print("=== 転記元の対応（確認画面に出るもの）===")
    for cl, sn in pairs:
        b = find_block(wb, sn)
        if sn not in wb.sheetnames:
            note = "★シートがありません"
        elif not b:
            note = "★破砕のブロックがありません"
        else:
            r0, r1, at, a_s, ct, cs = b
            note = (f"({r0}-{r1}行 As={gl(at)}/{gl(a_s)} "
                    + (f"Co={gl(ct)}/{gl(cs)})" if cs else "Coなし)"))
        print(f"  {cl}列 → {sn:<12} {note}")

    last = min(ws.max_row, 500)
    rows = [r for r in range(1, last + 1) if is_section_row(ws, r, first_col)]
    print(f"\n{SECTION_LABEL} の行: {len(rows)} 行 "
          f"({rows[0]}〜{rows[-1]})" if rows else "対象行なし")

    written, skipped, notes = {}, [], []
    for r in rows:
        tr = thick_ref(ws, r, first_col)
        kd = kind_of_row(ws, r, first_col)
        if not tr or not kd:
            continue
        for cl, sn in pairs:
            if not is_input_cell(ws, r, ci(cl)):
                continue
            f, note = build_sumif(wb, sn, ws.title, tr, kd)
            if not f:
                skipped.append((r, cl, note))
            else:
                written[(r, cl)] = f
                if note:
                    notes.append((r, cl, note))

    print(f"\n=== 書き込むセル {len(written)} 個 ===")
    by_row = collections.defaultdict(list)
    for (r, cl) in written:
        by_row[r].append(cl)
    for r in sorted(by_row):
        thk = thick_ref(ws, r, first_col)
        print(f"  {r:3}行 厚さ{thk:<6} {kind_of_row(ws, r, first_col):<5} "
              f"→ {','.join(sorted(by_row[r], key=ci))}")

    print("\n=== 13〜25行 J〜N の数式 ===")
    for r in range(13, 26):
        for cl in ("J", "K", "L", "M", "N"):
            if (r, cl) in written:
                print(f"  {cl}{r}: {written[(r, cl)]}")

    want = ("=SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,"
            "'試掘（舗50'!$P$11:$P$17)")
    got = written.get((14, "J"), "")
    ok = got == want
    print("\nJ14 が指示どおりか:", "一致" if ok else f"違う\n  期待 {want}\n  実際 {got}")

    if skipped:
        print(f"\n見送り {len(skipped)} 個")
        for r, cl, w in skipped[:10]:
            print(f"   {cl}{r}: {w}")
    if notes:
        print(f"\n注意 {len(notes)} 個")
        for r, cl, w in notes[:5]:
            print(f"   {cl}{r}: {w}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
