#!/usr/bin/env python3
"""M_Link が生成する INDIRECT 数式の検証

VBA の「リンク設定を取り込む」→「数式を再生成」と同じ処理を再現し、
生成した数式が今の直接リンクと同じ値になることを確かめる。

    cd excel/original && python3 ../docs/verify_link_formulas.py

VBA を書き換えたら、こちらのロジックも合わせて更新すること。
"""
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as gl

BOOK = "06_dokou_hosou.xlsx"
TARGET = "総括表（土工事）"
LNK = "リンク設定"

TERM = re.compile(r"'([^']+)'!(\$?[A-Z]{1,3}\$?[0-9]+)")
GEN = re.compile(
    r'IFERROR\(INDIRECT\("\'([^"]+)"&\'' + LNK + r'\'!\$C\$(\d+)&"\'!(\$?[A-Z]{1,3}\$?[0-9]+)"\),0\)'
)


def trail_digits(s):
    """末尾に続く数字。VBA の TrailDigits と同じ"""
    m = re.search(r"(\d+)$", s)
    return m.group(1) if m else ""


def make_template(f, dia):
    """数式中の該当口径を # にする。VBA の MakeTemplate と同じ"""
    return re.sub(r"'([^']+)" + re.escape(dia) + r"'!", r"'\1#'!", f)


def build_indirect(tmpl, dia_cell):
    """テンプレート → INDIRECT 数式。VBA の BuildIndirect と同じ"""

    def rep(m):
        return f"IFERROR(INDIRECT(\"'{m.group(1)}\"&'{LNK}'!{dia_cell}&\"'!{m.group(2)}\"),0)"

    return re.sub(r"'([^']+)#'!(\$?[A-Z]{1,3}\$?[0-9]+)", rep, tmpl)


def collect(ws):
    """リンクを拾って、列ごとの口径を集める。VBA の取り込みと同じ"""
    links, col_dia = [], {}
    for row in ws.iter_rows():
        for c in row:
            if not (isinstance(c.value, str) and c.value.startswith("=")):
                continue
            f = c.value
            if "!" not in f or "表紙" in f:
                continue
            terms = TERM.findall(f)
            if not terms:
                continue
            cl = gl(c.column)
            if len(terms) != f.count("!"):
                # 引用符なしのシート名など想定外の書き方。手を出さない
                links.append((c.coordinate, cl, f, "?"))
                continue
            ds, bad = [], False
            for pre, _addr in terms:
                d = trail_digits(pre)
                if not d:
                    bad = True
                    break
                if d not in ds:
                    ds.append(d)
            if bad:
                continue
            col_dia.setdefault(cl, [])
            for d in ds:
                if d not in col_dia[cl]:
                    col_dia[cl].append(d)
            links.append((c.coordinate, cl, f, ",".join(ds)))
    return links, col_dia


def main():
    wbf = openpyxl.load_workbook(BOOK)
    wbv = openpyxl.load_workbook(BOOK, data_only=True)
    ws, wsv = wbf[TARGET], wbv[TARGET]

    links, col_dia = collect(ws)
    cols = sorted(col_dia, key=lambda x: (len(x), x))
    dia_cell = {cl: f"$C${6 + i}" for i, cl in enumerate(cols)}
    dia_val = {dia_cell[cl]: col_dia[cl][0] for cl in cols if len(col_dia[cl]) == 1}

    print("口径対応表:")
    for cl in cols:
        print(f"   {dia_cell[cl]}  {cl}列 → {col_dia[cl]}")

    ok = ng = skip = 0
    for co, cl, f, ds in links:
        if ds == "?" or "," in ds or len(col_dia.get(cl, [])) != 1:
            skip += 1
            continue
        gen = build_indirect(make_template(f, ds), dia_cell[cl])
        body = gen[1:]
        parts = GEN.findall(body)
        total, err = 0.0, ""
        if len(parts) != body.count("IFERROR"):
            err = "生成式の形が不正"
        else:
            for pre, drow, addr in parts:
                sn = pre + dia_val[f"$C${drow}"]
                if sn not in wbv.sheetnames:
                    continue  # IFERROR → 0
                v = wbv[sn][addr.replace("$", "")].value
                total += float(v) if isinstance(v, (int, float)) else 0.0
        cur = wsv[co].value
        cur = float(cur) if isinstance(cur, (int, float)) else None
        good = not err and cur is not None and abs(total - cur) < 1e-9
        if good:
            ok += 1
        else:
            ng += 1
            print(f"   NG {co}: {f}\n      → {gen}\n      got={total} cur={cur} {err}")

    print(f"\n生成式の評価: 一致 {ok} / {ok + ng}   （手動のまま {skip} 本）")
    return 0 if ng == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
