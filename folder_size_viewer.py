# -*- coding: utf-8 -*-
"""
フォルダ容量ビューア
====================
指定したフォルダ(ローカル or ネットワークサーバーの共有フォルダ)の
直下にあるフォルダ・ファイルを、容量の大きい順に一覧表示するツールです。

【使い方】
  python folder_size_viewer.py
  を実行し、上部の「参照...」ボタンでフォルダを選ぶか、
  パス欄にネットワークパス(例: \\\\server\\share\\フォルダ)を直接入力して
  Enter または「開く」を押してください。

  一覧でフォルダをダブルクリックすると、その中へさらに掘り進めます。
  「戻る」で1階層戻れます。「再スキャン」で最新の状態を取り直します。
  「CSVで保存」で現在の一覧をCSVファイルとして保存できます(結果の記録・バックアップ用)。

【安全のための方針】
  このツールに削除機能はありません。容量の大きいフォルダ・ファイルを
  見つけたら、通常のエクスプローラーから確認のうえ削除してください。

【動作環境】
  Python 3.8以降の標準ライブラリのみで動作します(tkinter含む)。
  追加インストールは不要です。

【引き継ぎ・保守メモ】
  - このファイル1つで完結しています。改造する場合はこのファイルを直接編集してください。
  - スキャンは別スレッドで実行しており、ネットワーク越しの大きいフォルダでも
    ウィンドウが固まらないようにしています。
  - 一度スキャンした階層は self.cache にパスをキーとして保持し、
    「戻る」で戻ったときは再スキャンしません。「再スキャン」ボタンで強制的に取り直せます。
  - Python環境のない同僚に配布したい場合は、下記コマンドで実行ファイル(.exe)化できます。
        pip install pyinstaller
        pyinstaller --onefile --windowed folder_size_viewer.py
    生成された dist フォルダ内の .exe を配布すればPython不要で動きます。
"""

import os
import csv
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def format_bytes(n):
    """バイト数を人が読みやすい単位(B/KB/MB/GB/TB)に変換する"""
    if n == 0:
        return "0 B"
    units = ["B", "KB", "MB", "GB", "TB"]
    i = 0
    n = float(n)
    while n >= 1024 and i < len(units) - 1:
        n /= 1024
        i += 1
    if i == 0:
        return f"{int(n)} {units[i]}"
    return f"{n:.1f} {units[i]}"


class FolderSizeViewer:
    def __init__(self, root):
        self.root = root
        root.title("フォルダ容量ビューア")
        root.geometry("860x580")
        root.minsize(600, 400)

        self.history = []      # 現在地までのパスの履歴(ドリルダウン用)
        self.cache = {}        # path -> (entries, error_count)
        self.entry_names = {}  # treeitem_id -> フォルダ/ファイル名
        self.cancel_flag = threading.Event()
        self.msg_queue = queue.Queue()
        # スキャンごとに番号を振り、中止した古いスキャンの結果が
        # あとから届いても無視できるようにする
        self.scan_id = 0

        self._build_ui()
        self.root.after(100, self._poll_queue)

    # ---------- UI構築 ----------
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="パス:").pack(side="left")
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(top, textvariable=self.path_var)
        self.path_entry.pack(side="left", padx=4, fill="x", expand=True)
        self.path_entry.bind("<Return>", lambda e: self.go_to_path(self.path_var.get()))

        ttk.Button(top, text="参照...", command=self.browse_folder).pack(side="left", padx=2)
        ttk.Button(top, text="開く", command=lambda: self.go_to_path(self.path_var.get())).pack(side="left", padx=2)

        toolbar = ttk.Frame(self.root, padding=(8, 0, 8, 8))
        toolbar.pack(fill="x")
        self.back_btn = ttk.Button(toolbar, text="← 戻る", command=self.go_back, state="disabled")
        self.back_btn.pack(side="left")
        self.rescan_btn = ttk.Button(toolbar, text="再スキャン", command=self.rescan, state="disabled")
        self.rescan_btn.pack(side="left", padx=4)
        self.csv_btn = ttk.Button(toolbar, text="CSVで保存", command=self.export_csv, state="disabled")
        self.csv_btn.pack(side="left", padx=4)
        self.cancel_btn = ttk.Button(toolbar, text="スキャン中止", command=self.cancel_scan)

        self.status_var = tk.StringVar(value="パスを入力するか「参照...」でフォルダを選んでください")
        ttk.Label(self.root, textvariable=self.status_var, padding=(8, 0, 8, 4), foreground="#555").pack(fill="x")

        columns = ("rank", "type", "name", "size", "bar")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("rank", text="順位")
        self.tree.heading("type", text="種類")
        self.tree.heading("name", text="名前")
        self.tree.heading("size", text="サイズ")
        self.tree.heading("bar", text="割合")
        self.tree.column("rank", width=45, anchor="e")
        self.tree.column("type", width=55, anchor="center")
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("size", width=90, anchor="e")
        self.tree.column("bar", width=190, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.tree.bind("<Double-1>", self.on_double_click)

        scrollbar = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

    # ---------- 操作 ----------
    def browse_folder(self):
        path = filedialog.askdirectory(title="フォルダを選択(ネットワークサーバーも指定可)")
        if path:
            self.go_to_path(path)

    def go_to_path(self, path):
        path = path.strip()
        if not path:
            return
        if not os.path.isdir(path):
            messagebox.showerror("エラー", f"フォルダが見つかりません:\n{path}")
            return
        self.history = [path]
        self.cache.clear()
        self._scan(path)

    def go_back(self):
        if len(self.history) <= 1:
            return
        self.history.pop()
        self._scan(self.history[-1], use_cache=True)

    def rescan(self):
        if not self.history:
            return
        path = self.history[-1]
        self.cache.pop(path, None)
        self._scan(path)

    def cancel_scan(self):
        self.cancel_flag.set()

    def on_double_click(self, _event):
        item = self.tree.focus()
        if not item:
            return
        tags = self.tree.item(item, "tags")
        if "dir" not in tags:
            return
        name = self.entry_names.get(item)
        if not name or not self.history:
            return
        new_path = os.path.join(self.history[-1], name)
        self.history.append(new_path)
        self._scan(new_path, use_cache=True)

    # ---------- スキャン処理 ----------
    def _scan(self, path, use_cache=False):
        self.path_var.set(path)
        if use_cache and path in self.cache:
            self._render(path, self.cache[path])
            return

        # 前のスキャンが動いていれば止め、新しいスキャン用の中止フラグを作る
        self.cancel_flag.set()
        self.cancel_flag = threading.Event()
        self.scan_id += 1

        # 前の一覧を残すと、中止・エラー後にダブルクリックで
        # 違う場所のパスを組み立ててしまうため、いったん空にする
        self.tree.delete(*self.tree.get_children())
        self.entry_names = {}

        self._set_scanning(True)
        self.status_var.set("スキャン中...")

        thread = threading.Thread(
            target=self._scan_worker,
            args=(path, self.scan_id, self.cancel_flag),
            daemon=True,
        )
        thread.start()

    def _scan_worker(self, path, scan_id, cancel_flag):
        """別スレッドで動く。結果は msg_queue 経由でメインスレッドに渡す"""
        def send(kind, payload):
            self.msg_queue.put((scan_id, kind, payload))

        entries = []
        error_count = 0
        scanned = 0
        try:
            with os.scandir(path) as it:
                dir_entries = list(it)
        except Exception as e:
            send("error", str(e))
            return

        for entry in dir_entries:
            if cancel_flag.is_set():
                send("cancelled", None)
                return
            try:
                if entry.is_dir(follow_symlinks=False):
                    size, ec, sc = self._dir_size(entry.path, cancel_flag)
                    error_count += ec
                    scanned += sc
                    entries.append((entry.name, True, size))
                else:
                    try:
                        size = entry.stat(follow_symlinks=False).st_size
                    except Exception:
                        size = 0
                        error_count += 1
                    entries.append((entry.name, False, size))
                    scanned += 1
            except Exception:
                error_count += 1
            send("progress", scanned)

        if cancel_flag.is_set():
            send("cancelled", None)
            return
        entries.sort(key=lambda x: x[2], reverse=True)
        send("done", (path, entries, error_count))

    def _dir_size(self, path, cancel_flag):
        """フォルダ以下の合計サイズを再帰的に計算する(アクセス不可はスキップ)"""
        total = 0
        error_count = 0
        scanned = 0
        for dirpath, _dirnames, filenames in os.walk(path, onerror=lambda e: None):
            if cancel_flag.is_set():
                break
            for fname in filenames:
                fp = os.path.join(dirpath, fname)
                try:
                    total += os.path.getsize(fp)
                    scanned += 1
                except Exception:
                    error_count += 1
        return total, error_count, scanned

    # ---------- UI更新(メインスレッド) ----------
    def _set_scanning(self, scanning):
        self.back_btn.config(state="disabled" if scanning or len(self.history) <= 1 else "normal")
        self.rescan_btn.config(state="disabled" if scanning or not self.history else "normal")
        self.csv_btn.config(state="disabled" if scanning or not self.history else "normal")
        if scanning:
            self.cancel_btn.pack(side="left", padx=4)
        else:
            self.cancel_btn.pack_forget()

    def _poll_queue(self):
        try:
            while True:
                scan_id, kind, payload = self.msg_queue.get_nowait()
                if scan_id != self.scan_id:
                    continue  # 中止済みの古いスキャンからの結果は捨てる
                if kind == "progress":
                    self.status_var.set(f"スキャン中... {payload}件処理済み")
                elif kind == "done":
                    path, entries, error_count = payload
                    self.cache[path] = (entries, error_count)
                    self._render(path, (entries, error_count))
                    self._set_scanning(False)
                elif kind == "error":
                    self.status_var.set("フォルダを開けませんでした")
                    messagebox.showerror("エラー", f"フォルダを開けません:\n{payload}")
                    self._set_scanning(False)
                elif kind == "cancelled":
                    self.status_var.set("スキャンを中止しました")
                    self._set_scanning(False)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _render(self, path, data):
        entries, error_count = data
        self.tree.delete(*self.tree.get_children())
        self.entry_names = {}
        total = sum(e[2] for e in entries)

        for i, (name, is_dir, size) in enumerate(entries, start=1):
            pct = (size / total * 100) if total > 0 else 0
            bar_len = 20
            filled = int(pct / 100 * bar_len)
            bar = "█" * filled + "░" * (bar_len - filled) + f" {pct:.0f}%"
            item_id = self.tree.insert(
                "", "end",
                values=(i, "フォルダ" if is_dir else "ファイル", name, format_bytes(size), bar),
                tags=("dir" if is_dir else "file",)
            )
            self.entry_names[item_id] = name

        err_text = f" / アクセス不可: {error_count}件" if error_count else ""
        self.status_var.set(f"{path}   項目数: {len(entries)} / 合計: {format_bytes(total)}{err_text}")
        self.back_btn.config(state="normal" if len(self.history) > 1 else "disabled")

    # ---------- CSV出力 ----------
    def export_csv(self):
        if not self.history:
            return
        path = self.history[-1]
        data = self.cache.get(path)
        if not data:
            return
        entries, _ = data
        save_path = filedialog.asksaveasfilename(
            title="CSVとして保存",
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv")],
            initialfile="容量一覧.csv"
        )
        if not save_path:
            return
        try:
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["名前", "種類", "サイズ(バイト)"])
                for name, is_dir, size in entries:
                    writer.writerow([name, "フォルダ" if is_dir else "ファイル", size])
            messagebox.showinfo("保存完了", f"保存しました:\n{save_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました:\n{e}")


def main():
    root = tk.Tk()
    FolderSizeViewer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
