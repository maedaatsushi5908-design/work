#!/usr/bin/env python3
"""M_Hasai が書き込む数式を、VBA と同じ手順で再現する。

作成環境に Excel が無いので、マクロの判定をそのまま Python に写して
「どのセルに何が入るか」を出す。VBA を直したらこちらも直すこと。

    python3 excel/docs/verify_hasai.py
    WORK=02_nagata python3 excel/docs/verify_hasai.py
"""
import collections
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter as gl
from openpyxl.utils import column_index_from_string as ci

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get("WORK", "01_higashishirakawa")
FOLDER = os.path.join(HERE, "..", "works", WORK)
BOOK = os.environ.get("BOOK", "06_dokou_hosou.xlsx")

# ---- M_Hasai の先頭にある設定と同じもの -----------------------------------
TARGET_SHEET = "総括表（土工事）"
COL_MAP = os.environ.get("COL_MAP") or (
    "J=試掘（舗50|K=試掘（舗300|L=試掘（舗75|M=試掘（舗400|N=試掘（舗600|"
    "O=給水2度|P=仮配（舗|Q=給水(舗|"
    "R=管工（舗50|S=管工（舗75|T=管工（舗400|U=管工（舗600")
SECTION_LABEL = "舗装版破砕"
BLOCK_LABEL = "□舗装版破砕"
CUT_LABEL = "舗装切断工"
CUT_BLOCK = "□舗装切断工"
TORI_BLOCK = "□舗装版取壊工"      # 給水2度だけ並びが違う
TORI_KIND = "As"
KARA_LABEL = "殻運搬"
# 工種名 > 副見出しに含む語（＋区切り） > 転記元のブロック見出し
KARA_MAP = ("殻運搬>現場+仮置場>□As殻Co殻運搬（|"
            "殻運搬>現場+処分地>●殻運搬処理|"
            "殻運搬>積込み>●仮置土積込工")
# 仮配（舗・給水(舗・管工（舗50・75・400・600 の殻運搬（現場→処分地）だけは
# 並びが違う。転記元に「＝」の並びが無く、その1に似た並び（種別・舗装厚｜
# 合計）で、以下/超の順番も総括表の行の並びと単純に対応しない（15cm以下の
# 予備行が車道・歩道・Co・舗装で挟まる位置が違う）ため、kara_rows では
# 読めない。マクロには推測させず、この工事で確かめたシートごとの行の対応
# （総括表の行→転記元のセル）をそのまま書く。管工（舗50・75・400・600 は
# R83・S83・T84 など、この工事ではすでに手入力で正しい式が入っていたので、
# その式から行の対応を確かめて他の行にも広げた。
# 「15cm以下」の予備行は2つあり（83／84行、87／88行、91／92行）、
# 管工（舗50・75 は前の行（83・87・91）、管工（舗400・600 は後ろの行
# （84・88・92）を使う。「15cm超」（85・89・93行）は1行だけで全シート共通。
KARA_DIRECT_MAP = ("仮配（舗>83=T20|85=T21|87=T22|89=T23|91=Y20|93=Y21;"
                    "給水(舗>83=T19|85=T20|87=T21|89=T22|91=Y19|93=Y20;"
                    "管工（舗50>83=T27|85=T28|87=T29|89=T30|91=X27|93=X28;"
                    "管工（舗75>83=T27|85=T28|87=T29|89=T30|91=X27|93=X28;"
                    "管工（舗400>84=T27|85=T28|88=T29|89=T30|92=X27|93=X28;"
                    "管工（舗600>84=T27|85=T28|88=T29|89=T30|92=X27|93=X28")
GENDO_LABEL = "先行路盤（発生土）"    # 左端は「舗装仮復旧」で共通。副見出しで見分ける
GENDO_BLOCK = "□先行路盤（発生土）"
GENDO_BLOCK2 = "□先行路盤（再使用）"    # 試掘（舗50 などの、もう1つの転記元
FUKKYU_LABEL = "仮復旧"    # 「舗装仮復旧」に含まれるので完全一致で見分ける
FUKKYU_SRC = "給水2度"
FUKKYU_SIDE_MAP = "4=歩道|5=車道|10=車道"
FUKKYU_BLOCK2 = "□仮復旧工"    # 仮配（舗 のもう1つの転記元（その1と同じ並び）
ZENKOURO_LABEL = "先行路盤"    # 「先行路盤（発生土）」に含まれるので完全一致で見分ける
ZENKOURO_BLOCK = "□先行路盤"    # 材料名（再生／粒調）で選ぶ、もう1つの先行路盤
ZENKOURO_MATERIALS = "再生砕石|粒調砕石"    # 式に書く材料名（総括表側の書き方）
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


def is_sub_cell(ws, r, c):
    """殻運搬の欄は薄い橙（テーマ色アクセント6）。黄色と同じく入力セル"""
    if is_input_cell(ws, r, c):
        return True
    f = ws.cell(r, c).fill
    if f is None or f.patternType is None:
        return False
    fg = f.fgColor
    return fg.type == "theme" and fg.theme == 9      # accent6


# ---- 転記元シートのブロック ------------------------------------------------
_cache = {}


def sheet_ref(sn):
    """数式に書くシート名。囲む必要のある名前だけ ' で囲む（VBA の SheetRef）"""
    def safe(ch):
        o = ord(ch)
        if ch.isascii() and (ch.isalnum() or ch == "_"):
            return True
        if o in (0x3000, 0x30FB):
            return False
        return (0x3041 <= o <= 0x30FF) or (0x4E00 <= o <= 0x9FFF) or (0xFF66 <= o <= 0xFF9F)

    if not sn or sn[0].isdigit() or not all(safe(c) for c in sn):
        return "'" + sn.replace("'", "''") + "'!"
    return sn + "!"


def merge_wide(ws, r, c):
    for m in ws.merged_cells.ranges:
        if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
            return m.max_col - m.min_col + 1
    return 1


def find_block(wb, sn, anchor=BLOCK_LABEL):
    """転記元のブロックを探し、(r0, r1, 枠の一覧) を返す（VBA の FindBlock）

    枠の一覧は [(種別, 厚さ列, 合計列, 合計の幅), ...]
    """
    key = (anchor, sn)
    if key in _cache:
        return _cache[key]
    _cache[key] = None
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]
    sect = 0
    for r in range(1, 61):
        for c in range(1, 61):
            if norm(anchor) in norm(ws.cell(r, c).value):
                sect = r
                break
        if sect:
            break
    if not sect:
        return None

    out = scan_kind_header(ws, sect) or scan_gokou_header(ws, sect)
    _cache[key] = out
    return out


def scan_kind_header(ws, sect):
    """並び その1 －「種別・舗装厚」と「合 計」が並ぶ形"""
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

    r0, r1 = hdr + 1, hdr
    for r in range(r0, r0 + 41):
        if norm(ws.cell(r, kinds[0]).value) not in ("AS", "CO"):
            break
        r1 = r
    if r1 < r0:
        return None

    pairs = []
    for kc in kinds:
        nxt = [t for t in totals if t > kc]
        if not nxt:
            continue
        sc = min(nxt)
        k = norm(ws.cell(r0, kc).value)
        k = "As" if k == "AS" else ("Co" if k == "CO" else ("As" if not pairs else "Co"))
        pairs.append((k, kc + 1, sc, merge_wide(ws, r0, sc)))
        if len(pairs) == 2:
            break
    if not pairs:
        return None
    return (r0, r1 + 1, pairs)          # 予備を1行


def scan_gokou_header(ws, sect):
    """並び その2 －「車道 5号工」「歩道 5号工」が並ぶ形（給水2度）"""
    hdr, cols = 0, []
    for r in range(sect, sect + 4):
        cols = [c for c in range(1, 61) if norm("号工") in norm(ws.cell(r, c).value)]
        if cols:
            hdr = r
            break
    if not hdr:
        return None

    r0, r1 = hdr + 1, hdr
    for r in range(r0, r0 + 41):
        if ws.cell(r, cols[0]).value is None:
            break
        r1 = r
    if r1 < r0:
        return None

    pairs = []
    for c in cols:
        start, wide = c, merge_wide(ws, hdr, c)
        for m in ws.merged_cells.ranges:
            if m.min_row <= hdr <= m.max_row and m.min_col <= c <= m.max_col:
                start = m.min_col
                break
        pairs.append((TORI_KIND, start, start + 1, max(1, wide - 1)))
    return (r0, r1, pairs) if pairs else None      # 予備は付けない


def rng(col, wide, r0, r1):
    return f"${gl(col)}${r0}:${gl(col + wide - 1)}${r1}"



def hasai_block(wb, sn):
    """舗装版破砕の転記元。給水2度だけ見出しが違うので順に試す"""
    return find_block(wb, sn, BLOCK_LABEL) or find_block(wb, sn, TORI_BLOCK)


def build_sumif(wb, sn, tname, thk_ref, kind):
    b = hasai_block(wb, sn)
    if not b:
        return "", f"{sn} に {BLOCK_LABEL} のブロックがありません"
    r0, r1, pairs = b
    crit = sheet_ref(tname) + thk_ref
    q = sheet_ref(sn)
    terms = []
    for k, thk, tot, w in pairs:
        if k in kind:
            terms.append(f"SUMIF({q}{rng(thk, 1, r0, r1)},{crit},{q}{rng(tot, w, r0, r1)})")
    if not terms:
        return "", f"{sn} に {kind} 側の欄がありません"
    note = ""
    if "Co" in kind and not any(k == "Co" for k, *_ in pairs):
        note = f"{sn} に Co 側の欄が無いため As だけを合計しました"
    return "=" + "+".join(terms), note


def scan_gendo_header(ws, sect):
    """並び その3 －「路盤厚｜面 積」が並ぶ形（給水2度・先行路盤（発生土））

    厚さ欄（路盤厚）が結合で広く、合計欄（面積）は1列。舗装版破砕とは逆。
    見出しが複数回現れても最初の1枠だけを使う。
    """
    hdr, tc, sc = 0, 0, 0
    for r in range(sect, sect + 4):
        tc = sc = 0
        for c in range(1, 61):
            v = norm(ws.cell(r, c).value)
            if not tc and v == norm("路盤厚"):
                tc = c
            if not sc and v == norm("面 積"):
                sc = c
        if tc and sc:
            hdr = r
            break
    if not hdr:
        return None

    r0, r1 = hdr + 1, hdr
    for r in range(r0, r0 + 41):
        if ws.cell(r, tc).value is None:
            break
        r1 = r
    if r1 < r0:
        return None

    mr, mc = merged_top(ws, r0, tc)
    thk_wide = merge_wide(ws, mr, mc)
    return (r0, r1, mc, thk_wide, sc, merge_wide(ws, r0, sc))


def scan_saiyou_header(ws, sect, kind_header="種別・路盤厚"):
    """その1と同じ形だが種別が As/Co の2種ではなく1種類だけ続くもの
    （VBA の ScanSaiyouHeader）。見出しの文字は kind_header で渡す
    （試掘（舗50 などの先行路盤（再使用）は「種別・路盤厚」、仮配（舗 の
    仮復旧工は「種別・舗装厚」）。種別の欄が先頭行と同じ文字で続くところ
    までを行の範囲とする。厚さ・合計とも幅は転記元の結合セルに合わせる
    （仮配（舗 は2列結合、試掘は結合なしの1列）。2つ目の枠があっても
    使わない。
    """
    hdr, kinds, totals = 0, [], []
    for r in range(sect + 1, sect + 4):
        kinds, totals = [], []
        for c in range(1, 61):
            v = norm(ws.cell(r, c).value)
            if v == norm(kind_header):
                kinds.append(c)
            if v == norm("合 計"):
                totals.append(c)
        if kinds and totals:
            hdr = r
            break
    if not hdr:
        return None

    kc = kinds[0]
    nxt = [t for t in totals if t > kc]
    if not nxt:
        return None
    sc = min(nxt)
    thk_col = kc + 1

    r0 = hdr + 1
    want = norm(ws.cell(r0, kc).value)
    if not want:
        return None
    r1 = r0 - 1
    for r in range(r0, r0 + 41):
        if norm(ws.cell(r, kc).value) != want:
            break
        r1 = r
    if r1 < r0:
        return None

    thk_wide = merge_wide(ws, r0, thk_col)
    sum_wide = merge_wide(ws, r0, sc)
    return (r0, r1, thk_col, thk_wide, sc, sum_wide)


def gendo_block(wb, sn):
    """先行路盤（発生土）の転記元。(r0, r1, 厚さ列, 厚さ幅, 合計列, 合計幅)（VBA の GendoBlock）

    見出しの文字（□先行路盤（発生土）／□先行路盤（再使用））と並びの形
    （その3／その7）の組み合わせは工事・シートによって違う（仮配（舗 は
    見出しが「発生土」なのに並びは「その7」）ので、見出し2種 × 並び2種を
    順に試す。
    """
    key = ("GENDO", sn)
    if key in _cache:
        return _cache[key]
    _cache[key] = None
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]

    # 仮配（舗 には見出しがもう1つ（…No.2）あるので、そちらは飛ばす。
    # No.2 は別ブック構成用の予備で、この工事は No.1（無印）側だけを使う。
    for anchor in (GENDO_BLOCK, GENDO_BLOCK2):
        sect = 0
        for r in range(1, 61):
            for c in range(1, 61):
                v = norm(ws.cell(r, c).value)
                if norm(anchor) in v and "NO.2" not in v:
                    sect = r
                    break
            if sect:
                break
        if not sect:
            continue
        out = scan_gendo_header(ws, sect) or scan_saiyou_header(ws, sect)
        if out:
            _cache[key] = out
            return out
    return None


def gendo_ref(wb, sn, tname, thk_ref):
    """先行路盤（発生土）の1セル分（VBA の GendoRef）

    ブロックの無いシート（試掘・管工など）は、この工種がまだ未対応という
    だけなので黙って飛ばす（殻運搬の kara_ref と同じ扱い。見送りには数えない）
    """
    b = gendo_block(wb, sn)
    if not b:
        return "", ""
    r0, r1, thk_col, thk_wide, sum_col, sum_wide = b
    crit = sheet_ref(tname) + thk_ref
    q = sheet_ref(sn)
    return ("=" + f"SUMIF({q}{rng(thk_col, thk_wide, r0, r1)},{crit},"
            f"{q}{rng(sum_col, sum_wide, r0, r1)})"), ""


def fukkyu_side(thick):
    """厚さ(cm) → 車道/歩道（VBA の FukkyuSide）。マップに無ければ空文字

    厚さは数値で比べる（VBA と同じ）。セルの表示形式（.Text 相当）だと
    "4cm" や "4.0" のように化けて一致しなくなる不具合が実際に起きた。
    """
    for p in FUKKYU_SIDE_MAP.split("|"):
        k, _, v = p.partition("=")
        if float(k.strip()) == float(thick):
            return v.strip()
    return ""


def fukkyu_block2(wb, sn):
    """仮配（舗 の仮復旧の転記元。(r0, r1, 厚さ列, 厚さ幅, 合計列, 合計幅)
    （VBA の FukkyuBlock2）"""
    key = ("FUKKYU2", sn)
    if key in _cache:
        return _cache[key]
    _cache[key] = None
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]
    # 仮配（舗 には見出しがもう1つ（…No.2）あるので、そちらは飛ばす
    # （gendo_block と同じ理由）
    sect = 0
    for r in range(1, 61):
        for c in range(1, 61):
            v = norm(ws.cell(r, c).value)
            if norm(FUKKYU_BLOCK2) in v and "NO.2" not in v:
                sect = r
                break
        if sect:
            break
    if not sect:
        return None
    out = scan_saiyou_header(ws, sect, "種別・舗装厚")
    _cache[key] = out
    return out


def fukkyu_ref(wb, sn, tname, thk_ref, thick):
    """仮復旧（再生As）の1セル分（VBA の FukkyuRef）

    給水2度は舗装版破砕と同じ「車道/歩道 5号工」の枠を使うが、車道＋歩道を
    足さずどちらか片方だけを使う。どちらの側かは工事ごとに違うので
    FUKKYU_SIDE_MAP（厚さ→車道/歩道）で決める。仮配（舗 は「□仮復旧工」
    （その1と同じ並び、枠が2つあっても最初の1枠だけ）を使う。それ以外は
    まだ対応が無いので黙って見送る。
    """
    if sn == FUKKYU_SRC:
        side = fukkyu_side(thick)
        if not side:
            return "", f"厚さ {thick} の車道/歩道が FUKKYU_SIDE_MAP にありません"
        b = hasai_block(wb, sn)
        if not b:
            return "", f"{sn} に {BLOCK_LABEL} のブロックがありません"
        r0, r1, pairs = b
        idx = 0 if side == "車道" else 1
        if idx >= len(pairs):
            return "", f"{sn} に {side} 側の欄がありません"
        _, thk, tot, w = pairs[idx]
        crit = sheet_ref(tname) + thk_ref
        q = sheet_ref(sn)
        return "=" + f"SUMIF({q}{rng(thk, 1, r0, r1)},{crit},{q}{rng(tot, w, r0, r1)})", ""

    b = fukkyu_block2(wb, sn)
    if not b:
        return "", ""
    r0, r1, thk_col, thk_wide, sum_col, sum_wide = b
    crit = sheet_ref(tname) + thk_ref
    q = sheet_ref(sn)
    return ("=" + f"SUMIF({q}{rng(thk_col, thk_wide, r0, r1)},{crit},"
            f"{q}{rng(sum_col, sum_wide, r0, r1)})"), ""


def scan_zenkouro_header(ws, sect):
    """並び －「種別・路盤厚｜合 計」が2枠あり、種別がラベル名そのもの
    （給水(舗 は「再生」「粒調」、管工は「再生砕石」「粒調砕石」）（VBA の
    ScanZenkouroHeader）。その1と似ているが種別が As/Co ではなく資材名
    なので、枠ごとにラベルをそのまま持ち帰る。
    """
    hdr, kinds, totals = 0, [], []
    for r in range(sect + 1, sect + 4):
        kinds, totals = [], []
        for c in range(1, 61):
            v = norm(ws.cell(r, c).value)
            if v == norm("種別・路盤厚"):
                kinds.append(c)
            if v == norm("合 計"):
                totals.append(c)
        if kinds and totals:
            hdr = r
            break
    if not hdr:
        return None

    r0 = hdr + 1
    want = norm(ws.cell(r0, kinds[0]).value)
    if not want:
        return None
    r1 = r0 - 1
    for r in range(r0, r0 + 41):
        if norm(ws.cell(r, kinds[0]).value) != want:
            break
        r1 = r
    if r1 < r0:
        return None

    pairs = []
    for kc in kinds:
        nxt = [t for t in totals if t > kc]
        if not nxt:
            continue
        sc = min(nxt)
        kind_label = str(ws.cell(r0, kc).value or "").strip()
        if not kind_label:
            continue
        thk_col = kc + 1
        thk_wide = merge_wide(ws, r0, thk_col)
        sum_wide = merge_wide(ws, r0, sc)
        pairs.append((kind_label, thk_col, thk_wide, sc, sum_wide))
    if not pairs:
        return None
    return (r0, r1, pairs)


def zenkouro_block(wb, sn):
    """先行路盤（材料名）の転記元。(r0, r1, [(材料名, 厚さ列, 厚さ幅, 合計列,
    合計幅), ...])（VBA の ZenkouroBlock）"""
    key = ("ZENKOURO", sn)
    if key in _cache:
        return _cache[key]
    _cache[key] = None
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]
    # 「先行路盤（発生土）」「先行路盤（再使用）」は括弧が続くので除く。
    # 「…No.2」も避ける（gendo_block と同じ理由）
    sect = 0
    for r in range(1, 61):
        for c in range(1, 61):
            v = norm(ws.cell(r, c).value)
            if norm(ZENKOURO_BLOCK) in v and norm("先行路盤（") not in v and "NO.2" not in v:
                sect = r
                break
        if sect:
            break
    if not sect:
        return None
    out = scan_zenkouro_header(ws, sect)
    _cache[key] = out
    return out


def zenkouro_ref(wb, sn, tname, thk_ref, r):
    """先行路盤（材料名）の1セル分（VBA の ZenkouroRef）

    厚さでは選べないので、材料名（E列）が枠のラベルと一致するかを式
    そのものに書く。式に書く比較文字列は総括表の材料名の書き方
    （ZENKOURO_MATERIALS）にそろえ、転記元の枠はラベルがその材料名に
    含まれるかどうかで選ぶ（給水(舗 は「再生」、管工は「再生砕石」と
    ラベルの書き方が違うが、どちらも「再生砕石」に含まれる）。
    マクロ側では材料名を見ず、どの行にも同じ式（行番号だけ違う）を
    入れる。最初の条件だけ総括表のシート名を付け、2つ目以降は付けない
    （指示された式のとおり）。枠に無い材料は自動化できる転記元が無いので
    そのまま文字列「手入力」を返す式にする。
    """
    b = zenkouro_block(wb, sn)
    if not b:
        return "", ""
    r0, r1, pairs = b

    mat_ref = sheet_ref(tname) + f"E{r}"
    mat_bare = f"E{r}"
    crit = sheet_ref(tname) + thk_ref
    q = sheet_ref(sn)

    mats = ZENKOURO_MATERIALS.split("|")
    terms = [""] * len(mats)
    first_idx = None
    for j, mat in enumerate(mats):
        for kind_label, thk_col, thk_wide, sum_col, sum_wide in pairs:
            if kind_label in mat:
                terms[j] = (f"SUMIF({q}{rng(thk_col, thk_wide, r0, r1)},{crit},"
                            f"{q}{rng(sum_col, sum_wide, r0, r1)})")
                if first_idx is None:
                    first_idx = j
                break
    if first_idx is None:
        return "", ""

    out = '"手入力"'
    for j in range(len(mats) - 1, -1, -1):
        if terms[j]:
            cond = mat_ref if j == first_idx else mat_bare
            out = f'IF({cond}="{mats[j]}",{terms[j]},{out})'
    return "=" + out, ""


def has_exact_label(ws, r, first_col, label):
    """左の欄のどれかが label とちょうど一致するか（部分一致ではなく）。
    「仮復旧」が左端の「舗装仮復旧」に含まれてしまうケースを区別するため
    （VBA の HasExactLabel）"""
    want = norm(label)
    for c in range(1, first_col):
        v = merged_value(ws, r, c)
        if v is None:
            continue
        t = str(v)
        if not t.startswith("=") and norm(t) == want:
            return True
    return False


def thick_cell_for(ws, r, tc, sect):
    """厚さの入っているセル (row, col) を返す。先行路盤（発生土）・仮復旧・
    先行路盤（材料名）は I ではなく H にあるので、厚さ列の1つ左で数値が
    見つかった列を厚さ欄とみなす（VBA と同じ）"""
    if sect in (GENDO_LABEL, FUKKYU_LABEL, ZENKOURO_LABEL):
        for cc in range(tc - 1, 0, -1):
            tr0, tc0 = merged_top(ws, r, cc)
            if is_num(ws.cell(tr0, tc0).value):
                return tr0, tc0
        return None
    return merged_top(ws, r, tc)


def cut_ref(wb, sn, label, kind):
    """舗装切断工の1セル分。行を探して直接参照にする（VBA の CutRef）"""
    b = find_block(wb, sn, CUT_BLOCK)
    if not b:
        return "", f"{sn} に {CUT_BLOCK} のブロックがありません"
    r0, r1, pairs = b
    hit = [(thk, tot) for k, thk, tot, w in pairs if k in kind]
    if not hit:
        return "", f"{sn} に {kind} 側の欄がありません"
    thk, tot = hit[0]
    want = norm_label(label)
    if not want:
        return "", "総括表の厚さ区分が読めません"
    ws = wb[sn]
    for r in range(r0, r1 + 1):
        if norm_label(ws.cell(r, thk).value) == want:
            return "=" + sheet_ref(sn) + f"{gl(tot)}{r}", ""
    return "", f"{sn} に「{label}」の行がありません"


# ---- 総括表側の読み取り ----------------------------------------------------
def section_of(ws, r, first_col):
    """いちばん左の工種名で、どの工種の行かを返す（VBA の SectionOf）"""
    head, whole = "", ""
    for c in range(1, first_col):
        v = merged_value(ws, r, c)
        if v is None:
            continue
        t = str(v)
        if t.strip() and not t.startswith("="):
            if not head:
                head = norm(t)
            whole += norm(t)
    if not head:
        return ""
    if norm(SECTION_LABEL) in head:
        return SECTION_LABEL
    if norm(CUT_LABEL) in head:
        return CUT_LABEL
    if norm(GENDO_LABEL) in whole:
        # 左端（B列）は「舗装仮復旧」で他の資材とも共通。副見出し
        # （C列）に「先行路盤（発生土）」があるかで見分ける
        return GENDO_LABEL
    if has_exact_label(ws, r, first_col, FUKKYU_LABEL):
        # 「仮復旧」は左端の「舗装仮復旧」自体にも含まれるので、部分一致
        # ではなく完全一致で見分ける
        return FUKKYU_LABEL
    if has_exact_label(ws, r, first_col, ZENKOURO_LABEL):
        # 「先行路盤」も同じ理由で完全一致で見分ける
        return ZENKOURO_LABEL
    if kara_anchor_of(head, whole):
        return KARA_LABEL
    return ""


def kara_anchor_of(head, whole):
    """殻運搬の行なら転記元のブロック見出しを返す（VBA の KaraAnchorOf）"""
    for e in KARA_MAP.split("|"):
        f = e.split(">")
        if len(f) < 3:
            continue
        if norm(f[0]) in head and all(norm(w) in whole for w in f[1].split("+")):
            return f[2]
    return ""


def kara_anchor(ws, r, first_col):
    head, whole = "", ""
    for c in range(1, first_col):
        v = merged_value(ws, r, c)
        if v is None:
            continue
        t = str(v)
        if t.strip() and not t.startswith("="):
            if not head:
                head = norm(t)
            whole += norm(t)
    return kara_anchor_of(head, whole)


def group_label(ws, r, first_col):
    """殻運搬の行の摘要（As・車道 など）（VBA の GroupLabel）"""
    for c in range(1, first_col):
        v = merged_value(ws, r, c)
        t = norm(v)
        if "AS" in t or "CO" in t:
            return str(v)
    return ""


def norm_group(v):
    """摘要をそろえる。「As・車道」も「As車道」も同じ（VBA の NormGroup）"""
    return norm(v).replace("\u30fb", "")


def kara_rows(wb, sn, anchor):
    """見出しの下の「＝」の行を集める。[(摘要, 数量の列, 行), …]（VBA の KaraRows）"""
    key = ("KARA", anchor, sn)
    if key in _cache:
        return _cache[key]
    _cache[key] = None
    if sn not in wb.sheetnames:
        return None
    ws = wb[sn]
    sect = 0
    for r in range(1, 61):
        for c in range(1, 41):
            if norm(anchor) in norm(ws.cell(r, c).value):
                sect = r
                break
        if sect:
            break
    if not sect:
        return None

    out = []
    for r in range(sect + 1, sect + 13):
        # 次のブロックの見出し（□ や ●）が出たらそこまで
        if any("\u25a1" in norm(ws.cell(r, c).value) or "\u25cf" in norm(ws.cell(r, c).value)
               for c in range(1, 41)):
            break
        eq = 0
        for c in range(1, 41):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == "\uff1d":
                eq = c
                break
        if not eq:
            continue
        # 摘要は「＝」から左へ戻って最初に出てくる文字（VBA と同じ）
        for c in range(eq - 1, 0, -1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("="):
                continue
            if is_num(v):
                continue
            t = norm_group(v)
            if t and t not in ("\uff08", "\uff09", "(", ")", "-"):
                out.append((t, eq + 1, r))
                break
    _cache[key] = out or None
    return _cache[key]


def kara_direct_rows(sn):
    """KARA_DIRECT_MAP からシート名 sn の「行=セル|…」を取り出す
    （VBA の KaraDirectRows）。載っていなければ空文字列"""
    for e in KARA_DIRECT_MAP.split(";"):
        parts = e.split(">")
        if len(parts) == 2 and parts[0] == sn:
            return parts[1]
    return ""


def kara_direct_ref(sn, rows_, r):
    """kara_direct_rows で取り出した並びから、総括表の行番号 r に対応する
    セルの数式を組み立てる（VBA の KaraDirectRef）"""
    for p in rows_.split("|"):
        kv = p.split("=")
        if len(kv) == 2 and int(kv[0]) == r:
            return "=" + sheet_ref(sn) + kv[1]
    return ""


def kara_ref(wb, sn, anchor, label, r):
    """殻運搬の1セル分。行を探して直接参照にする（VBA の KaraRef）"""
    direct_rows = kara_direct_rows(sn)
    if direct_rows:
        return kara_direct_ref(sn, direct_rows, r), ""
    rows = kara_rows(wb, sn, anchor)
    if not rows:
        return "", ""            # 未対応のシートは黙って飛ばす
    want = norm_group(label)
    # 1. そのまま一致
    for lab, col, r in rows:
        if lab == want:
            return "=" + sheet_ref(sn) + f"{gl(col)}{r}", ""
    # 2. ラベルが摘要の頭に一致
    for lab, col, r in rows:
        if lab and want.startswith(lab):
            return "=" + sheet_ref(sn) + f"{gl(col)}{r}", ""
    # 3. 摘要に As / Co が無いなら先頭行
    if want[:2] not in ("AS", "CO"):
        lab, col, r = rows[0]
        return "=" + sheet_ref(sn) + f"{gl(col)}{r}", ""
    flat = " ".join(str(label).split())
    return "", f"{sn} に「{flat}」の行がありません"


def is_section_row(ws, r, first_col):
    return bool(section_of(ws, r, first_col))


def norm_label(v):
    """厚さ区分の文字をそろえる。t≦15㎝ も t≦15 も同じ（VBA の NormLabel）"""
    return norm(v).replace("\u339d", "").replace("CM", "")


def thick_col(ws, rows, first_col):
    """舗装厚の列。舗装版破砕の行で数値が現れた回数がいちばん多い列（VBA の ThickCol）

    舗装切断工は厚さが区分の文字なので、列を数えるときは見ない。
    """
    hits = collections.Counter()
    for r in rows:
        if section_of(ws, r, first_col) != SECTION_LABEL:
            continue
        for c in range(first_col - 1, 0, -1):
            tr, tc = merged_top(ws, r, c)
            if is_num(ws.cell(tr, tc).value):
                hits[tc] += 1
                break
    return hits.most_common(1)[0][0] if hits else 0


def is_text_cell(ws, r, c):
    """「計」などの文字が入っているか。空欄と数値は False"""
    tr, tc = merged_top(ws, r, c)
    v = ws.cell(tr, tc).value
    if v is None:
        return False
    return isinstance(v, str) and bool(v.strip())


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
    wb = openpyxl.load_workbook(os.path.join(FOLDER, BOOK))
    if TARGET_SHEET not in wb.sheetnames:
        print("シートがありません:", TARGET_SHEET)
        return 1
    ws = wb[TARGET_SHEET]
    print(f"ブック {WORK}/{BOOK}")

    pairs = [p.split("=", 1) for p in COL_MAP.split("|") if "=" in p]
    first_col = min(ci(c) for c, _ in pairs)

    print("=== 転記元の対応（確認画面に出るもの）===")
    for cl, sn in pairs:
        b = hasai_block(wb, sn)
        if sn not in wb.sheetnames:
            note = "★シートがありません"
        elif not b:
            note = "★破砕のブロックがありません"
        else:
            r0, r1, blk = b
            note = f"破砕 {r0}-{r1}行 " + " ".join(
                f"{k}={gl(t)}/{gl(u)}" + (f"×{w}" if w > 1 else "") for k, t, u, w in blk)
        c = find_block(wb, sn, CUT_BLOCK)
        note += f"  切断 {c[0]}-{c[1]}行" if c else "  切断なし"
        print(f"  {cl}列 → {sn:<12} {note}")

    last = min(ws.max_row, 500)
    rows = [r for r in range(1, last + 1) if is_section_row(ws, r, first_col)]
    print(f"\n{SECTION_LABEL} の行: {len(rows)} 行 "
          f"({rows[0]}〜{rows[-1]})" if rows else "対象行なし")

    tc = thick_col(ws, rows, first_col)
    print(f"舗装厚の列: {gl(tc)} 列")

    written, skipped, notes, sect_of = {}, [], [], {}
    for r in rows:
        sect = section_of(ws, r, first_col)
        istext = is_text_cell(ws, r, tc)
        grp = ""
        # 破砕は厚さが数値（「計」の行は触らない）、切断は区分の文字、
        # 殻運搬は厚さを使わず摘要で照合する
        if sect == SECTION_LABEL and istext:
            continue
        if sect == CUT_LABEL and not istext:
            continue
        anchor = ""
        if sect == KARA_LABEL:
            grp = group_label(ws, r, first_col)
            anchor = kara_anchor(ws, r, first_col)
            if not anchor:
                continue
        cell = thick_cell_for(ws, r, tc, sect)
        if cell is None:
            continue
        mr, mc = cell
        tr = f"${gl(mc)}{mr}"
        kd = kind_of_row(ws, r, first_col)
        # 殻運搬・先行路盤（発生土）・仮復旧・先行路盤（材料名）は種別
        # （As/Co）の欄が無いので、そこは問わない
        if not kd and sect not in (KARA_LABEL, GENDO_LABEL, FUKKYU_LABEL, ZENKOURO_LABEL):
            continue
        for cl, sn in pairs:
            ok = is_sub_cell(ws, r, ci(cl)) if sect == KARA_LABEL \
                else is_input_cell(ws, r, ci(cl))
            if not ok:
                continue
            if sect == SECTION_LABEL:
                f, note = build_sumif(wb, sn, ws.title, tr, kd)
            elif sect == CUT_LABEL:
                f, note = cut_ref(wb, sn, str(ws.cell(mr, mc).value), kd)
            elif sect == GENDO_LABEL:
                f, note = gendo_ref(wb, sn, ws.title, tr)
            elif sect == FUKKYU_LABEL:
                f, note = fukkyu_ref(wb, sn, ws.title, tr, ws.cell(mr, mc).value)
            elif sect == ZENKOURO_LABEL:
                f, note = zenkouro_ref(wb, sn, ws.title, tr, r)
            else:
                f, note = kara_ref(wb, sn, anchor, grp, r)
            if not f:
                if note:
                    skipped.append((r, cl, note))
            else:
                written[(r, cl)] = f
                sect_of[(r, cl)] = sect
                if note:
                    notes.append((r, cl, note))

    print(f"\n=== 書き込むセル {len(written)} 個 ===")
    by_row = collections.defaultdict(list)
    for (r, cl) in written:
        by_row[r].append(cl)
    for r in sorted(by_row):
        sect = section_of(ws, r, first_col)
        mr, mc = thick_cell_for(ws, r, tc, sect)
        thk = ws.cell(mr, mc).value
        if sect == KARA_LABEL:
            thk = " ".join(str(group_label(ws, r, first_col)).split())
        print(f"  {r:3}行 {sect:<6} {str(thk):<11} {kind_of_row(ws, r, first_col):<5} "
              f"→ {','.join(sorted(by_row[r], key=ci))}")

    print("\n=== 舗装切断工（9〜12行）の数式 ===")
    for r in range(9, 13):
        for cl in ("J", "M", "R"):
            if (r, cl) in written:
                print(f"  {cl}{r}: {written[(r, cl)]}")

    print("\n=== 13〜25行 J〜N の数式 ===")
    for r in range(13, 26):
        for cl in ("J", "K", "L", "M", "N"):
            if (r, cl) in written:
                print(f"  {cl}{r}: {written[(r, cl)]}")

    print("\n=== 先行路盤（発生土）162〜171行 O の数式 ===")
    for r in range(162, 172):
        if (r, "O") in written:
            print(f"  O{r}: {written[(r, 'O')]}")

    want = ("=SUMIF('試掘（舗50'!$O$11:$O$17,'総括表（土工事）'!$I14,"
            "'試掘（舗50'!$P$11:$P$17)")
    got = written.get((14, "J"), "")
    ok = got == want
    print("J14:", "一致" if ok else f"違う（{got}）")
    for cell, expect in (
            ("O77", "=給水2度!AA10"), ("O96", "=給水2度!AA25"),
            ("O83", "=給水2度!AA28"), ("O87", "=給水2度!AA30"),
            ("O31", "=SUMIF(給水2度!$K$4:$K$6,'総括表（土工事）'!$I31,給水2度!$L$4:$N$6)"
                    "+SUMIF(給水2度!$O$4:$O$6,'総括表（土工事）'!$I31,給水2度!$P$4:$R$6)"),
            ("J10", "='試掘（舗50'!P5"), ("J12", "='試掘（舗50'!S5"),
                         ("J9", "='試掘（舗50'!P4"), ("M11", "='試掘（舗400'!S4"),
                         ("R9", "='管工（舗50'!T10"), ("S11", "='管工（舗75'!X10"),
            ("O162", "=SUMIF(給水2度!$U$4:$W$6,'総括表（土工事）'!$H162,給水2度!$X$4:$X$6)"),
            ("O163", "=SUMIF(給水2度!$U$4:$W$6,'総括表（土工事）'!$H163,給水2度!$X$4:$X$6)"),
            ("O164", "=SUMIF(給水2度!$U$4:$W$6,'総括表（土工事）'!$H164,給水2度!$X$4:$X$6)"),
            ("O165", "=SUMIF(給水2度!$U$4:$W$6,'総括表（土工事）'!$H165,給水2度!$X$4:$X$6)"),
            ("O166", "=SUMIF(給水2度!$U$4:$W$6,'総括表（土工事）'!$H166,給水2度!$X$4:$X$6)"),
            ("O167", "=SUMIF(給水2度!$O$4:$O$6,'総括表（土工事）'!$H167,給水2度!$P$4:$R$6)"),
            ("O168", "=SUMIF(給水2度!$K$4:$K$6,'総括表（土工事）'!$H168,給水2度!$L$4:$N$6)"),
            ("O169", "=SUMIF(給水2度!$K$4:$K$6,'総括表（土工事）'!$H169,給水2度!$L$4:$N$6)"),
            ("J162", "=SUMIF('試掘（舗50'!$O$26:$O$31,'総括表（土工事）'!$H162,"
                     "'試掘（舗50'!$P$26:$P$31)"),
            ("P162", "=SUMIF('仮配（舗'!$R$27:$S$32,'総括表（土工事）'!$H162,"
                     "'仮配（舗'!$T$27:$U$32)"),
            ("P167", "=SUMIF('仮配（舗'!$R$36:$S$39,'総括表（土工事）'!$H167,"
                     "'仮配（舗'!$T$36:$U$39)"),
            ("Q156", "=IF('総括表（土工事）'!E156=\"再生砕石\","
                     "SUMIF('給水(舗'!$R$26:$S$30,'総括表（土工事）'!$H156,'給水(舗'!$T$26:$U$30),"
                     "IF(E156=\"粒調砕石\","
                     "SUMIF('給水(舗'!$W$26:$X$30,'総括表（土工事）'!$H156,'給水(舗'!$Y$26:$Z$30),"
                     "\"手入力\"))"),
            ("P83", "='仮配（舗'!T20"), ("P85", "='仮配（舗'!T21"),
            ("P87", "='仮配（舗'!T22"), ("P89", "='仮配（舗'!T23"),
            ("P91", "='仮配（舗'!Y20"), ("P93", "='仮配（舗'!Y21"),
            ("Q83", "='給水(舗'!T19"), ("Q85", "='給水(舗'!T20"),
            ("Q87", "='給水(舗'!T21"), ("Q89", "='給水(舗'!T22"),
            ("Q91", "='給水(舗'!Y19"), ("Q93", "='給水(舗'!Y20"),
            ("R83", "='管工（舗50'!T27"), ("S83", "='管工（舗75'!T27"),
            ("R85", "='管工（舗50'!T28"), ("S85", "='管工（舗75'!T28"),
            ("R87", "='管工（舗50'!T29"), ("S87", "='管工（舗75'!T29"),
            ("R89", "='管工（舗50'!T30"), ("S89", "='管工（舗75'!T30"),
            ("R91", "='管工（舗50'!X27"), ("S91", "='管工（舗75'!X27"),
            ("R93", "='管工（舗50'!X28"), ("S93", "='管工（舗75'!X28"),
            ("T84", "='管工（舗400'!T27"), ("U84", "='管工（舗600'!T27"),
            ("T85", "='管工（舗400'!T28"), ("U85", "='管工（舗600'!T28"),
            ("T88", "='管工（舗400'!T29"), ("U88", "='管工（舗600'!T29"),
            ("T89", "='管工（舗400'!T30"), ("U89", "='管工（舗600'!T30"),
            ("T92", "='管工（舗400'!X27"), ("U92", "='管工（舗600'!X27"),
            ("T93", "='管工（舗400'!X28"), ("U93", "='管工（舗600'!X28")):
        r = int(cell[1:])
        g = written.get((r, cell[0]), "")
        mark = "一致" if g == expect else f"違う（{g}）"
        print(f"{cell} = {expect}  … {mark}")
        ok = ok and g == expect
    print("\n指示された式がすべて一致したか:", "はい" if ok else "いいえ")

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
