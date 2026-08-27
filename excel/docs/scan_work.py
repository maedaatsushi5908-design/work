#!/usr/bin/env python3
"""工事フォルダを調べて、M_Hasai の COL_MAP の下書きを出す

新しい工事の数量計算書一式を excel/works/<名前>/ に入れたら、これを流す。
総括表の見出し・シート名・舗装版破砕ブロックの有無を突き合わせ、
マクロの先頭に貼る COL_MAP を作る。

    python3 excel/docs/scan_work.py excel/works/02
"""
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import verify_hasai as H            # noqa: E402  判定はマクロと同じものを使う


def header_of(ws, col):
    for r in range(1, 13):
        v = ws.cell(r, col).value
        if v not in (None, "") and H.norm(v) and not str(v).startswith("="):
            return " ".join(str(v).split())
    return ""


def kei_of(ws, col):
    """系統＝見出しの括弧より前"""
    h = header_of(ws, col)
    if not h:
        return ""
    for br in ("（", "("):
        p = h.find(br)
        if p > 0:
            h = h[:p]
            break
    return h.replace(" ", "").replace("　", "").strip()


def lead(sn):
    """シート名の括弧より前  仮配（舗 → 仮配"""
    t = H.norm(sn)
    for ch in ("(", "（"):
        p = t.find(ch)
        if p > 0:
            return t[:p]
    return t


def find_target(wb):
    """いちばん左の工種名が「舗装版破砕」の行があるシートを探す"""
    hits = []
    for sn in wb.sheetnames:
        ws = wb[sn]
        for r in range(1, min(ws.max_row, 250) + 1):
            for c in range(1, 10):
                v = H.merged_value(ws, r, c)
                if v in (None, "") or str(v).startswith("="):
                    continue
                if str(v).strip():
                    if H.norm(H.SECTION_LABEL) in H.norm(v):
                        hits.append((sn, r))
                    break            # その行のいちばん左だけ見る
            if hits and hits[-1][0] == sn:
                break
    return hits


def scan(folder):
    names = os.listdir(folder)
    books = sorted(f for f in names
                   if f.lower().endswith((".xlsx", ".xlsm")) and not f.startswith("~"))
    xls = sorted(f for f in names if f.lower().endswith(".xls"))
    if xls:
        print(f"※ .xls（97-2003 形式）は飛ばします: {', '.join(xls)}\n")

    found = False
    for bk in books:
        try:
            wb = openpyxl.load_workbook(os.path.join(folder, bk))
        except Exception as e:                                  # noqa: BLE001
            print(f"--- {bk}  読めません: {e}\n")
            continue

        targets = find_target(wb)
        if not targets:
            continue
        found = True

        print("=" * 74)
        print(f"■ {bk}　（シート {len(wb.sheetnames)} 枚）")

        H._cache.clear()
        srcs = [(sn, H.find_block(wb, sn)) for sn in wb.sheetnames]
        srcs = [(sn, b) for sn, b in srcs if b]
        block_of_all = {sn for sn, _ in srcs}
        print(f"\n  舗装版破砕のブロックを持つシート {len(srcs)} 枚")
        for sn, b in srcs:
            r0, r1, at, a_s, ct, cs = b
            co = f"Co={gl(ct)}/{gl(cs)}" if cs else "Co なし"
            print(f"    {sn:<16} {r0:>3}-{r1:<3}行  As={gl(at)}/{gl(a_s):<3} {co}")

        for tsn, trow in targets:
            # 自分が転記元のシートは総括表ではない
            if tsn in block_of_all:
                continue
            ws = wb[tsn]
            head = f"\n  ▼ 総括表候補「{tsn}」　{H.SECTION_LABEL} は {trow} 行目から"

            fc = 0
            for c in range(1, 40):
                if kei_of(ws, c):
                    fc = c
                    break
            if not fc:
                continue

            block_of = dict(srcs)
            cands, info = {}, {}
            for c in range(fc, min(ws.max_column, 60) + 1):
                k = kei_of(ws, c)
                if not k:
                    continue
                yellow = sum(1 for r in range(trow, min(trow + 90, ws.max_row + 1))
                             if H.is_input_cell(ws, r, c))
                if yellow == 0:
                    continue
                cl = gl(c)
                info[cl] = (header_of(ws, c), k, yellow)

                # 1. 名前がそのまま一致するシート（給水2度 → 給水2度）
                exact = [sn for sn in wb.sheetnames if H.norm(sn) == H.norm(k)]
                if exact:
                    cands[cl] = exact if exact[0] in block_of else []
                    continue
                # 2. 括弧より前が一致し、口径が見出しの数字に含まれるシート
                nums = re.findall(r"\d+", header_of(ws, c))
                lst = []
                for sn, _b in srcs:
                    if H.norm(k).startswith(lead(sn)):
                        tail = re.search(r"(\d+)$", sn)
                        if tail is None or tail.group(1) in nums:
                            lst.append(sn)
                cands[cl] = lst

            # 候補が少ない列から順に決め、使ったシートは他の列から外す。
            # 見出しの「PE50-300」だけでは絞れなくても、他の列が先に
            # 決まることで1つに残ることが多い。
            chosen, used = {}, set()
            todo = {c: list(v) for c, v in cands.items() if v}
            while todo:
                best = min(todo, key=lambda c: (len([x for x in todo[c] if x not in used]),
                                                list(cands).index(c)))
                left = [x for x in todo[best] if x not in used]
                if not left:
                    del todo[best]
                    continue
                chosen[best] = (left[0], len(left) == 1)
                used.add(left[0])
                del todo[best]

            if len([c for c in cands if cands[c]]) < 2:
                continue                       # 数量欄の無いシートは飛ばす
            print(head)
            print(f"    {'列':<3} {'見出し':<32} {'系統':<9} {'黄':>3}  候補シート")
            print("    " + "-" * 78)

            proposal, unsure = [], []
            for cl in cands:
                hdr, k, yellow = info[cl]
                if cl in chosen:
                    sn, sure = chosen[cl]
                    mark = "" if sure else "  ★候補が複数（他の列から決めました）"
                    if not sure:
                        unsure.append(cl)
                    proposal.append((cl, sn))
                    shown = sn + (f"（候補 {' / '.join(cands[cl])}）" if not sure else "")
                elif cands.get(cl) == [] and any(H.norm(sn) == H.norm(k)
                                                for sn in wb.sheetnames):
                    shown, mark = "同名シートに破砕ブロックが無い → 対象外", ""
                else:
                    shown, mark = "(候補なし)", "  ★要確認"
                    unsure.append(cl)
                print(f"    {cl:<3} {hdr[:32]:<32} {k:<9} {yellow:>3}  {shown}{mark}")

            if proposal:
                print("\n    --- マクロ先頭に貼る下書き ---")
                print(f'    Private Const TARGET_SHEET As String = "{tsn}"')
                print('    Private Const COL_MAP As String = _')
                for i, (cl, sn) in enumerate(proposal):
                    last = (i == len(proposal) - 1)
                    print(f'        "{cl}={sn}{"" if last else "|"}"{"" if last else " & _"}')
                if unsure:
                    print(f"\n    ★ {', '.join(unsure)} 列は自動で決め切れていません。"
                          "見出しとシート名を見比べて確かめてください。")
        print()

    if not found:
        print("舗装版破砕の区間がある総括表は見つかりませんでした。")
        print("シート名や工種名の書き方が違うかもしれません。")


def main(argv):
    if not argv:
        print(__doc__)
        return 1
    folder = argv[0]
    if not os.path.isdir(folder):
        print("フォルダがありません:", folder)
        return 1
    if not [f for f in os.listdir(folder) if f.lower().endswith((".xls", ".xlsx", ".xlsm"))]:
        print(f"{folder} に Excel ファイルがありません。")
        return 1
    scan(folder)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
