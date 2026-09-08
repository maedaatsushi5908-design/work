#!/usr/bin/env python3
"""M_HasaiHosou が書き込む数式を、VBA と同じ手順で再現する。

作成環境に Excel が無いので、マクロの判定をそのまま Python に写して
「どのセルに何が入るか」を出す。VBA を直したらこちらも直すこと。

    python3 excel/docs/verify_hasai_hosou.py
"""
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.environ.get("WORK", "01_higashishirakawa")
FOLDER = os.path.join(HERE, "..", "works", WORK)
BOOK = os.environ.get("BOOK", "06_dokou_hosou.xlsx")

# ---- M_HasaiHosou の先頭にある設定と同じもの -------------------------------
TARGET_SHEET = "総括表（舗装工事）"
PAVE_SRC = "舗装（集計）"
CELL_MAP = "I7=M4|I8=M5|I9=P4|I10=P5"
INPUT_COLOR = "FFFF00"


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


def is_input_cell(ws, addr):
    """黄色く塗ってある入力セルか（VBA の Interior.Color = 65535 と同じ）"""
    f = ws[addr].fill
    if f is None or f.patternType is None:
        return False
    fg = f.fgColor
    if fg.type == "rgb":
        return isinstance(fg.rgb, str) and fg.rgb.endswith(INPUT_COLOR)
    if fg.type == "indexed":
        return fg.indexed == 13        # 既定パレットの黄色 = FFFF00
    return False


def main():
    path = os.path.join(FOLDER, BOOK)
    wb = openpyxl.load_workbook(path, data_only=False)
    if TARGET_SHEET not in wb.sheetnames:
        print(f"シートが見つかりません: {TARGET_SHEET}")
        return 1
    if PAVE_SRC not in wb.sheetnames:
        print(f"転記元シートが見つかりません: {PAVE_SRC}")
        return 1
    ws = wb[TARGET_SHEET]

    written = {}
    skipped = []
    for p in CELL_MAP.split("|"):
        addr, cell_ref = p.split("=")
        if is_input_cell(ws, addr):
            written[addr] = "=" + sheet_ref(PAVE_SRC) + cell_ref
        else:
            skipped.append(addr)

    ok = True
    for addr, expect in (
            ("I7", "='舗装（集計）'!M4"),
            ("I8", "='舗装（集計）'!M5"),
            ("I9", "='舗装（集計）'!P4"),
            ("I10", "='舗装（集計）'!P5")):
        g = written.get(addr, "")
        mark = "一致" if g == expect else f"違う（{g}）"
        print(f"{addr} = {expect}  … {mark}")
        ok = ok and g == expect

    print("\n指示された式がすべて一致したか:", "はい" if ok else "いいえ")
    if skipped:
        print(f"\n見送り {len(skipped)} 個（黄色でないセル）: {', '.join(skipped)}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
