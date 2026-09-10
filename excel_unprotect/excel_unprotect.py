#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_unprotect - Excel の「ブックの保護」「シートの保護」を解除するツール
=========================================================================

自分が所有する Excel ファイルで、シート／ブック構造の保護パスワードを
忘れてしまったときに、その保護を解除する（外す）ためのツールです。

■ 何ができるか
--------------
以下の「編集ロック系」の保護を解除できます。パスワードを解読するのではなく、
保護そのものを取り除きます（これらは暗号化ではなく、単なる編集ロックのため
安全に外せます）。

  - ブックの保護        (workbookProtection : シート追加/削除/移動の禁止など)
  - シートの保護        (sheetProtection    : セル編集の禁止など)
  - グラフシートの保護  (chartsheet の sheetProtection)
  - 書き込み予約        (fileSharing        : 読み取り推奨/書き込みパスワード)

対応形式: .xlsx / .xlsm / .xltx / .xltm など、
          ZIP ベースの新しい Office 形式 (OOXML)。

■ 何ができないか（重要）
------------------------
「ファイルを開くためのパスワード」（暗号化されたファイル）は、本物の AES 暗号
なので解除できません。この場合は正しいパスワードが必要です。
本ツールは暗号化ファイルを検出したら、その旨を明示して中断します。

古い .xls 形式（BIFF）は本ツールの対象外です（別途メッセージを表示します）。

■ 使い方（コマンドライン）
--------------------------
    python3 excel_unprotect.py 保護されたファイル.xlsx
        -> 保護されたファイル_unprotected.xlsx を生成

    python3 excel_unprotect.py in.xlsx -o out.xlsx
    python3 excel_unprotect.py in.xlsx --inplace       # 元ファイルを直接書き換え(バックアップ付き)
    python3 excel_unprotect.py in.xlsx --dry-run       # 解除せず、保護の有無だけ調べる

作者ではなく所有者による、自分のファイルへの利用を想定しています。
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import List


# --- ファイルシグネチャ（先頭バイト）--------------------------------------
ZIP_MAGIC = b"PK\x03\x04"              # OOXML (.xlsx など) は実体が ZIP
OLE_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"  # 旧 .xls / 暗号化 OOXML の器

# --- 除去対象の保護要素（XMLの空要素）------------------------------------
# 名前空間プレフィックス（例: x:sheetProtection）にも一応対応する。
_PROTECTION_PATTERNS = {
    "workbookProtection": re.compile(rb"<(?:\w+:)?workbookProtection\b[^>]*?/>", re.IGNORECASE),
    "sheetProtection":    re.compile(rb"<(?:\w+:)?sheetProtection\b[^>]*?/>", re.IGNORECASE),
    "fileSharing":        re.compile(rb"<(?:\w+:)?fileSharing\b[^>]*?/>", re.IGNORECASE),
}


@dataclass
class RemovalResult:
    """解除処理の結果をまとめて返すためのデータ。"""
    input_path: str
    output_path: str = ""
    removed: List[str] = field(default_factory=list)   # 例: ["xl/workbook.xml: workbookProtection", ...]
    changed_files: int = 0

    @property
    def had_protection(self) -> bool:
        return bool(self.removed)


class UnsupportedFileError(Exception):
    """暗号化ファイルや旧 .xls など、本ツールで処理できないファイル。"""


def _sniff(path: str) -> bytes:
    """ファイル先頭のマジックバイトを読む。"""
    with open(path, "rb") as f:
        return f.read(8)


def _classify(path: str) -> str:
    """ファイル種別を判定して 'ooxml' / 'ole' / 'unknown' を返す。"""
    head = _sniff(path)
    if head.startswith(ZIP_MAGIC):
        return "ooxml"
    if head.startswith(OLE_MAGIC):
        return "ole"
    return "unknown"


def _is_target_xml(name: str) -> bool:
    """保護要素が含まれ得る XML パートかどうか。"""
    lname = name.lower()
    return (
        lname == "xl/workbook.xml"
        or lname.startswith("xl/worksheets/") and lname.endswith(".xml")
        or lname.startswith("xl/chartsheets/") and lname.endswith(".xml")
    )


def _strip_protection(data: bytes, part_name: str, removed: List[str]) -> bytes:
    """1つのXMLパートから保護要素を取り除く。"""
    for label, pattern in _PROTECTION_PATTERNS.items():
        new_data, count = pattern.subn(b"", data)
        if count:
            data = new_data
            for _ in range(count):
                removed.append(f"{part_name}: {label}")
    return data


def inspect(path: str) -> RemovalResult:
    """解除は行わず、どんな保護が掛かっているかだけを調べる（dry-run 用）。"""
    kind = _classify(path)
    if kind == "ole":
        raise UnsupportedFileError(_ole_message(path))
    if kind != "ooxml":
        raise UnsupportedFileError(
            "Excel ファイル (.xlsx/.xlsm など) として認識できませんでした。"
        )

    result = RemovalResult(input_path=path)
    with zipfile.ZipFile(path, "r") as zin:
        for name in zin.namelist():
            if not _is_target_xml(name):
                continue
            data = zin.read(name)
            before = len(result.removed)
            _strip_protection(data, name, result.removed)
            if len(result.removed) > before:
                result.changed_files += 1
    return result


def remove_protection(path: str, output_path: str) -> RemovalResult:
    """
    保護を解除した新しいファイルを output_path に書き出す。
    元の ZIP の内容は保護要素以外そのまま維持する。
    """
    kind = _classify(path)
    if kind == "ole":
        raise UnsupportedFileError(_ole_message(path))
    if kind != "ooxml":
        raise UnsupportedFileError(
            "Excel ファイル (.xlsx/.xlsm など) として認識できませんでした。"
        )

    result = RemovalResult(input_path=path, output_path=output_path)

    # 同一ディレクトリ内の一時ファイルに書いてから原子的に置き換える。
    out_dir = os.path.dirname(os.path.abspath(output_path)) or "."
    fd, tmp_path = tempfile.mkstemp(prefix=".xlunprotect_", dir=out_dir)
    os.close(fd)

    try:
        with zipfile.ZipFile(path, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if _is_target_xml(item.filename):
                    before = len(result.removed)
                    data = _strip_protection(data, item.filename, result.removed)
                    if len(result.removed) > before:
                        result.changed_files += 1

                # 元エントリの圧縮方式・メタ情報をなるべく維持して書き戻す。
                new_info = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                new_info.compress_type = item.compress_type
                new_info.external_attr = item.external_attr
                new_info.internal_attr = item.internal_attr
                new_info.create_system = item.create_system
                zout.writestr(new_info, data)

        os.replace(tmp_path, output_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

    return result


def _ole_message(path: str) -> str:
    """OLE ファイル（暗号化 OOXML か旧 .xls）向けの説明メッセージ。"""
    return (
        "このファイルは OLE 形式（旧 .xls、または『ファイルを開く』パスワードで\n"
        "暗号化された Excel ファイル）です。\n\n"
        "・『ファイルを開く』パスワードは本物の暗号化のため、本ツールでは解除\n"
        "  できません。正しいパスワードが必要です。\n"
        "・旧 .xls 形式のブック/シート保護は本ツールの対象外です。いったん\n"
        "  正しいパスワードで開いて .xlsx として保存し直してからご利用ください。"
    )


def _default_output(path: str) -> str:
    root, ext = os.path.splitext(path)
    return f"{root}_unprotected{ext or '.xlsx'}"


def _run_cli(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(
        prog="excel_unprotect",
        description="Excel のブック/シート保護（編集ロック）を解除します。"
                    "『ファイルを開く』パスワード（暗号化）は対象外です。",
    )
    parser.add_argument("input", help="対象の Excel ファイル (.xlsx/.xlsm など)")
    parser.add_argument("-o", "--output", help="出力先ファイル名（省略時は *_unprotected.xlsx）")
    parser.add_argument("--inplace", action="store_true",
                        help="元ファイルを直接書き換える（.bak バックアップを作成）")
    parser.add_argument("--dry-run", action="store_true",
                        help="解除は行わず、保護の有無だけを表示する")
    args = parser.parse_args(argv)

    if not os.path.isfile(args.input):
        print(f"エラー: ファイルが見つかりません: {args.input}", file=sys.stderr)
        return 2

    try:
        if args.dry_run:
            res = inspect(args.input)
            if res.had_protection:
                print("次の保護が見つかりました:")
                for item in res.removed:
                    print(f"  - {item}")
            else:
                print("保護は見つかりませんでした（すでに解除済みか、掛かっていません）。")
            return 0

        if args.inplace:
            backup = args.input + ".bak"
            shutil.copy2(args.input, backup)
            output = args.input
        else:
            output = args.output or _default_output(args.input)

        res = remove_protection(args.input, output)

    except UnsupportedFileError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 3
    except zipfile.BadZipFile:
        print("エラー: ファイルが壊れているか、Excel 形式ではありません。", file=sys.stderr)
        return 3

    if res.had_protection:
        print(f"保護を解除しました（{res.changed_files} 個のパートを変更）:")
        for item in res.removed:
            print(f"  - {item}")
        print(f"\n出力: {res.output_path}")
        if args.inplace:
            print(f"バックアップ: {args.input}.bak")
    else:
        print("保護は見つかりませんでした。ファイルはそのままコピーされました。")
        print(f"出力: {res.output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_run_cli(sys.argv[1:]))
