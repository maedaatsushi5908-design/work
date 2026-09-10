#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel_unprotect_gui - Excel の保護解除ツール（GUI 版）
======================================================

ダブルクリック感覚で使える簡単な画面版です。ファイルを選ぶと、
掛かっている保護（ブックの保護／シートの保護など）を調べて表示し、
[保護を解除する] ボタンで解除済みファイルを書き出します。

起動方法:
    python3 excel_unprotect_gui.py

※ 本ツールが解除できるのは「編集ロック系」の保護のみです。
   「ファイルを開くパスワード」（暗号化）は解除できません。
   （詳しくは excel_unprotect.py の説明を参照）
"""

from __future__ import annotations

import os
import sys

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
except Exception:  # pragma: no cover - tkinter が無い環境向け
    sys.stderr.write(
        "この GUI には tkinter が必要です。\n"
        "コマンドライン版 (excel_unprotect.py) をお使いください。\n"
    )
    raise SystemExit(1)

# 同じフォルダにあるコア機能を読み込む。
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import excel_unprotect as core  # noqa: E402


class App:
    def __init__(self, root: "tk.Tk") -> None:
        self.root = root
        self.selected_path: str = ""
        root.title("Excel 保護解除ツール")
        root.geometry("620x460")
        root.minsize(520, 400)

        pad = {"padx": 12, "pady": 6}

        # --- 1. ファイル選択 -------------------------------------------
        top = tk.Frame(root)
        top.pack(fill="x", **pad)
        tk.Button(top, text="① Excel ファイルを選ぶ…",
                  command=self.choose_file, height=2).pack(side="left")
        self.path_label = tk.Label(top, text="（未選択）", anchor="w", fg="#555")
        self.path_label.pack(side="left", fill="x", expand=True, padx=8)

        # --- 2. 状態表示 -----------------------------------------------
        self.log = scrolledtext.ScrolledText(root, height=12, wrap="word",
                                             state="disabled")
        self.log.pack(fill="both", expand=True, padx=12, pady=6)

        # --- 3. 実行ボタン ---------------------------------------------
        bottom = tk.Frame(root)
        bottom.pack(fill="x", **pad)
        self.run_btn = tk.Button(bottom, text="② 保護を解除する",
                                 command=self.run_remove, height=2,
                                 state="disabled")
        self.run_btn.pack(side="left")
        tk.Button(bottom, text="終了", command=root.destroy,
                  height=2).pack(side="right")

        self._write(
            "使い方:\n"
            "  1) 上の［① Excel ファイルを選ぶ…］でファイルを選択します。\n"
            "  2) 掛かっている保護がここに表示されます。\n"
            "  3) ［② 保護を解除する］を押すと、解除済みファイルを保存します。\n\n"
            "解除できるのは『ブックの保護』『シートの保護』などの編集ロックです。\n"
            "『ファイルを開くパスワード』（暗号化）は解除できません。\n"
        )

    # -- ユーティリティ ------------------------------------------------
    def _write(self, text: str, clear: bool = False) -> None:
        self.log.configure(state="normal")
        if clear:
            self.log.delete("1.0", "end")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    # -- ① ファイル選択 ------------------------------------------------
    def choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Excel ファイルを選択",
            filetypes=[("Excel ブック", "*.xlsx *.xlsm *.xltx *.xltm"),
                       ("すべてのファイル", "*.*")],
        )
        if not path:
            return
        self.selected_path = path
        self.path_label.config(text=os.path.basename(path))
        self.run_btn.config(state="disabled")
        self._inspect()

    def _inspect(self) -> None:
        self._write(f"選択: {self.selected_path}", clear=True)
        try:
            res = core.inspect(self.selected_path)
        except core.UnsupportedFileError as exc:
            self._write("\n⚠ このファイルは解除できません:\n" + str(exc))
            return
        except Exception as exc:
            self._write(f"\n⚠ ファイルを読めませんでした: {exc}")
            return

        if res.had_protection:
            self._write("\n次の保護が見つかりました:")
            for item in res.removed:
                self._write("  ・ " + item)
            self._write("\n［② 保護を解除する］を押してください。")
            self.run_btn.config(state="normal")
        else:
            self._write("\n保護は見つかりませんでした。"
                        "（すでに解除済みか、掛かっていません）")

    # -- ② 解除実行 ----------------------------------------------------
    def run_remove(self) -> None:
        if not self.selected_path:
            return
        default_out = core._default_output(self.selected_path)
        out = filedialog.asksaveasfilename(
            title="解除後のファイルの保存先",
            initialfile=os.path.basename(default_out),
            initialdir=os.path.dirname(default_out),
            defaultextension=os.path.splitext(self.selected_path)[1] or ".xlsx",
            filetypes=[("Excel ブック", "*.xlsx *.xlsm *.xltx *.xltm"),
                       ("すべてのファイル", "*.*")],
        )
        if not out:
            return
        try:
            res = core.remove_protection(self.selected_path, out)
        except core.UnsupportedFileError as exc:
            messagebox.showerror("解除できません", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("エラー", f"処理に失敗しました:\n{exc}")
            return

        self._write(f"\n✔ 完了: {res.changed_files} 個のパートから保護を削除しました。")
        self._write(f"保存先: {res.output_path}")
        messagebox.showinfo(
            "完了",
            "保護を解除しました。\n\n保存先:\n" + res.output_path,
        )


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
