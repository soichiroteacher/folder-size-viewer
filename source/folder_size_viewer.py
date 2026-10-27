# -*- coding: utf-8 -*-
"""
フォルダ容量ビューア
====================
指定したフォルダ(ローカル or ネットワークサーバーの共有フォルダ)の中身を
エクスプローラーのように表示し、あわせて容量を表示するツールです。

【画面の構成】
  ・「フォルダの中身」タブ
      エクスプローラーと同じように、今いるフォルダの中のフォルダ・ファイルを
      表示します(フォルダが先、名前順)。フォルダにも中身の合計容量を表示します。
      列の見出しをクリックすると並べ替えられます(「サイズ」で容量の大きい順)。
      フォルダをダブルクリックすると中に入り、「↑ 上のフォルダへ」で戻ります。
  ・「容量ランキング」タブ
      指定したパス以下にあるすべてのファイル(またはフォルダ)を、
      容量の大きい順に表示します。ダブルクリックすると、その場所を
      「フォルダの中身」タブで開きます。
      「更新日」で「3年以上前」などを選ぶと、長く更新されていない
      大きいファイル(フォルダ)だけに絞り込めます。

【使い方】
  上部の「参照...」ボタンでフォルダを選ぶか、パス欄にネットワークパス
  (例: \\\\server\\share\\フォルダ)を直接入力して Enter または「開く」を押します。
  そのパス以下を1回だけまとめてスキャンし、両方のタブに結果を表示します。
  「再スキャン」で最新の状態を取り直します。
  「CSVで保存」で、表示中のタブの一覧をCSVファイルとして保存できます。
  「エクスプローラーで表示」で、選んだ項目をエクスプローラーで開けます。

【安全のための方針】
  このツールに削除機能はありません。容量の大きいフォルダ・ファイルを
  見つけたら、エクスプローラーで中身を確認のうえ削除してください。

【動作環境】
  Python 3.8以降の標準ライブラリのみで動作します(tkinter含む)。

【引き継ぎ・保守メモ】
  - このファイル1つで完結しています。
  - スキャン処理は scan_tree() 関数にまとめてあり、画面とは独立しています。
    別スレッドで実行するので、ネットワーク越しの大きいフォルダでも
    ウィンドウは固まりません。
  - スキャン結果として持つのは「全フォルダの合計容量」と「大きいファイル上位
    TOP_FILES 件」だけです。全ファイルを覚えるとメモリを使いすぎるためです。
    「フォルダの中身」タブは、表示のたびにそのフォルダだけを読み直し、
    フォルダの容量はスキャン結果から取り出して表示しています。
  - フォルダの読み直しや、入力されたパスの確認も別スレッドで行います。
    つながらないネットワークパスは、確認だけで数十秒かかることがあるためです。
    別スレッドの結果はすべて msg_queue で画面側(_poll_queue)に届けます。
  - .exe の作り方は build_exe.bat と README.md を参照してください。
"""

import csv
import heapq
import os
import queue
import re
import subprocess
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 容量ランキングで覚えておくファイルの最大件数
TOP_FILES = 1000
# ランキングの表示件数の選択肢
RANK_LIMITS = ("100", "500", "1000")
# 「更新日」で絞り込むときの選択肢(〇年以上前)。変えたいときはここを直す
OLD_YEARS = (1, 3, 5)
NO_FILTER = "指定しない"
YEAR_SECONDS = 365.25 * 24 * 60 * 60


# ====================================================================
# 共通の小さな関数
# ====================================================================
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


def format_time(ts):
    """更新日時(エポック秒)を「2026/09/26 22:40」の形にする"""
    if ts is None:
        return ""
    try:
        return time.strftime("%Y/%m/%d %H:%M", time.localtime(ts))
    except (OSError, OverflowError, ValueError):
        return ""


def type_text(name, is_dir):
    """エクスプローラーの「種類」列に近い文字列を返す"""
    if is_dir:
        return "ファイル フォルダー"
    ext = os.path.splitext(name)[1]
    if not ext and name.startswith(".") and len(name) > 1:
        # 「.gitignore」のように点で始まる名前は、エクスプローラーでは
        # 名前全体を拡張子として扱う(「GITIGNORE ファイル」と表示される)
        ext = name
    return f"{ext[1:].upper()} ファイル" if ext else "ファイル"


def natural_key(name):
    """エクスプローラーと同じく「2」が「10」より先に来るように並べるためのキー"""
    parts = re.split(r"(\d+)", name.lower())
    return [(0, int(p), "") if p.isdigit() else (1, 0, p) for p in parts]


def is_real_dir(entry):
    """中に入ってよいフォルダかどうか。
    シンボリックリンクやジャンクション(別の場所への近道)はたどらない。
    たどると同じ中身を二重に数えたり、無限ループになったりするため。"""
    try:
        if not entry.is_dir(follow_symlinks=False):
            return False
        is_junction = getattr(entry, "is_junction", None)  # Python 3.12 以降
        return not (is_junction and is_junction())
    except OSError:
        return False


# ====================================================================
# スキャン処理(画面とは独立。別スレッドから呼ばれる)
# ====================================================================
class ScanResult:
    """1回のスキャン結果"""

    def __init__(self, root):
        self.root = root
        self.dir_sizes = {}    # フォルダのパス -> 中身の合計容量(バイト)
        self.dir_counts = {}   # フォルダのパス -> 中にあるファイル数(下の階層も含む)
        self.dir_newest = {}   # フォルダのパス -> 中のファイルで一番新しい更新日時(ファイルがなければ None)
        self.top_files = []    # (容量, パス, 更新日時) 大きい順
        self.old_files = {}    # 年数 -> その年数以上更新されていないファイルの top_files
        self.scanned_at = time.time()
        self.error_count = 0   # 読めなかったフォルダ・ファイルの数
        self.file_count = 0


def scan_tree(root, cancel_flag, progress=None, top_n=TOP_FILES, old_years=OLD_YEARS):
    """root 以下をすべて調べ、ScanResult を返す。中止されたら None を返す。
    progress には (ファイル数, フォルダ数) を受け取る関数を渡せる。"""
    result = ScanResult(root)
    order = []                 # 調べた順のフォルダ一覧(親が必ず子より先)
    parent = {root: None}
    own_size = {}              # そのフォルダ直下のファイルだけの合計
    own_count = {}
    own_newest = {}            # そのフォルダ直下のファイルで一番新しい更新日時
    heap = []                  # 大きいファイル上位 top_n 件(最小ヒープ)
    # 「〇年以上更新されていないファイル」の上位も別に覚えておく。
    # 全体の上位 top_n 件から絞り込むだけだと、古いファイルがほとんど残らないため
    cutoffs = {y: result.scanned_at - y * YEAR_SECONDS for y in old_years}
    old_heaps = {y: [] for y in old_years}
    stack = [root]
    last_report = 0.0

    def keep_top(h, item):
        if len(h) < top_n:
            heapq.heappush(h, item)
        elif item[0] > h[0][0]:
            heapq.heapreplace(h, item)

    while stack:
        if cancel_flag.is_set():
            return None
        d = stack.pop()
        order.append(d)
        size = 0
        count = 0
        newest = None
        try:
            with os.scandir(d) as it:
                for e in it:
                    try:
                        if is_real_dir(e):
                            parent[e.path] = d
                            stack.append(e.path)
                        else:
                            st = e.stat(follow_symlinks=False)
                            size += st.st_size
                            count += 1
                            item = (st.st_size, e.path, st.st_mtime)
                            keep_top(heap, item)
                            for y, cut in cutoffs.items():
                                if st.st_mtime < cut:
                                    keep_top(old_heaps[y], item)
                            if newest is None or st.st_mtime > newest:
                                newest = st.st_mtime
                    except OSError:
                        result.error_count += 1
        except OSError:
            # アクセス権がない・ネットワークが切れた などで読めないフォルダ
            result.error_count += 1
        own_size[d] = size
        own_count[d] = count
        own_newest[d] = newest
        result.file_count += count

        now = time.monotonic()
        if progress and now - last_report > 0.2:
            progress(result.file_count, len(order))
            last_report = now

    # 下の階層から順に、子フォルダの合計を親フォルダへ足し上げる
    sizes = dict(own_size)
    counts = dict(own_count)
    newests = dict(own_newest)
    for d in reversed(order):
        p = parent[d]
        if p is not None:
            sizes[p] += sizes[d]
            counts[p] += counts[d]
            if newests[d] is not None and (newests[p] is None or newests[d] > newests[p]):
                newests[p] = newests[d]
    result.dir_sizes = sizes
    result.dir_counts = counts
    result.dir_newest = newests
    result.top_files = sorted(heap, reverse=True)
    result.old_files = {y: sorted(h, reverse=True) for y, h in old_heaps.items()}
    return result


def list_folder(path):
    """フォルダ直下の項目を一覧にする(エクスプローラー表示用)。
    戻り値: [(名前, フルパス, フォルダか, ファイル容量 or None, 更新日時), ...]"""
    items = []
    with os.scandir(path) as it:
        for e in it:
            try:
                is_dir = is_real_dir(e)
                st = e.stat(follow_symlinks=False)
                items.append((e.name, e.path, is_dir, None if is_dir else st.st_size, st.st_mtime))
            except OSError:
                items.append((e.name, e.path, False, None, None))
    return items


# ====================================================================
# 画面
# ====================================================================
class FolderSizeViewer:
    def __init__(self, root):
        self.root = root
        root.title("フォルダ容量ビューア")
        root.geometry("1000x640")
        root.minsize(720, 420)

        self.root_path = None   # スキャンしたパス(一番上)
        self.current = None     # 「フォルダの中身」タブで今見ているフォルダ
        self.result = None      # ScanResult(スキャン中・中止時は None)
        self.scanning = False
        self.folder_items = []  # 今のフォルダの中身(list_folder の戻り値)
        self.folder_rows = []   # 上を容量付き・画面の並び順にしたもの(CSV 出力用)
        self.row_paths = {}     # 「フォルダの中身」の行ID -> (フルパス, フォルダか)
        self.rank_paths = {}    # 「容量ランキング」の行ID -> (フルパス, フォルダか, 容量, ファイル数)
        self.sort_col = "name"  # 「フォルダの中身」タブの並べ替え列
        self.sort_desc = False

        self.cancel_flag = threading.Event()
        # 別スレッドからの知らせ。(種類, 番号, 中身) の形で入る
        self.msg_queue = queue.Queue()
        # スキャン・フォルダ読み込み・パス確認のそれぞれに番号を振り、
        # 古い(もう要らなくなった)結果があとから届いても無視できるようにする
        self.scan_id = 0
        self.folder_req = 0
        self.check_id = 0
        self.loading_path = None  # 読み込み中のフォルダ(読み込み中でなければ None)

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
        self.path_entry.bind("<Return>", lambda e: self.open_path(self.path_var.get()))
        ttk.Button(top, text="参照...", command=self.browse_folder).pack(side="left", padx=2)
        ttk.Button(top, text="開く", command=lambda: self.open_path(self.path_var.get())).pack(side="left", padx=2)

        toolbar = ttk.Frame(self.root, padding=(8, 0, 8, 6))
        toolbar.pack(fill="x")
        self.up_btn = ttk.Button(toolbar, text="↑ 上のフォルダへ", command=self.go_up, state="disabled")
        self.up_btn.pack(side="left")
        self.rescan_btn = ttk.Button(toolbar, text="再スキャン", command=self.rescan, state="disabled")
        self.rescan_btn.pack(side="left", padx=4)
        self.csv_btn = ttk.Button(toolbar, text="CSVで保存", command=self.export_csv, state="disabled")
        self.csv_btn.pack(side="left", padx=4)
        self.explorer_btn = ttk.Button(toolbar, text="エクスプローラーで表示",
                                       command=self.show_in_explorer, state="disabled")
        self.explorer_btn.pack(side="left", padx=4)
        self.cancel_btn = ttk.Button(toolbar, text="スキャン中止", command=self.cancel_scan)

        self.status_var = tk.StringVar(value="パスを入力するか「参照...」でフォルダを選んでください")
        ttk.Label(self.root, textvariable=self.status_var, padding=(8, 0, 8, 4),
                  foreground="#555").pack(fill="x")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._build_folder_tab()
        self._build_rank_tab()

    def _make_tree(self, parent, columns):
        """スクロールバー付きの一覧(Treeview)を作る。columns: [(ID, 見出し, 幅, 寄せ), ...]"""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", selectmode="browse")
        for cid, text, width, anchor in columns:
            tree.heading(cid, text=text)
            tree.column(cid, width=width, anchor=anchor, stretch=(cid in ("name", "location")))
        ysb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=ysb.set)
        ysb.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)
        return tree

    def _build_folder_tab(self):
        tab = ttk.Frame(self.notebook, padding=(0, 6, 0, 0))
        self.notebook.add(tab, text="フォルダの中身")
        self.folder_label = tk.StringVar()
        ttk.Label(tab, textvariable=self.folder_label, padding=(2, 0, 2, 4)).pack(fill="x")
        self.folder_cols = [
            ("name", "名前", 340, "w"),
            ("mtime", "更新日時", 130, "w"),
            ("type", "種類", 120, "w"),
            ("size", "サイズ", 90, "e"),
            ("bar", "割合", 190, "w"),
        ]
        self.folder_tree = self._make_tree(tab, self.folder_cols)
        for cid, _t, _w, _a in self.folder_cols:
            if cid != "bar":
                self.folder_tree.heading(cid, command=lambda c=cid: self.sort_folder(c))
        self.folder_tree.bind("<Double-1>", self.on_folder_double_click)
        self.folder_tree.bind("<Return>", self.on_folder_double_click)
        self.folder_tree.bind("<BackSpace>", lambda e: self.go_up())
        self._update_sort_headings()

    def _build_rank_tab(self):
        tab = ttk.Frame(self.notebook, padding=(0, 6, 0, 0))
        self.notebook.add(tab, text="容量ランキング")
        opts = ttk.Frame(tab, padding=(2, 0, 2, 4))
        opts.pack(fill="x")
        ttk.Label(opts, text="対象:").pack(side="left")
        self.rank_kind = tk.StringVar(value="file")
        ttk.Radiobutton(opts, text="ファイル", value="file", variable=self.rank_kind,
                        command=self.render_rank).pack(side="left", padx=(4, 0))
        ttk.Radiobutton(opts, text="フォルダ", value="dir", variable=self.rank_kind,
                        command=self.render_rank).pack(side="left", padx=(4, 12))
        ttk.Label(opts, text="表示件数:").pack(side="left")
        self.rank_limit = tk.StringVar(value=RANK_LIMITS[0])
        combo = ttk.Combobox(opts, textvariable=self.rank_limit, values=RANK_LIMITS,
                             width=6, state="readonly")
        combo.pack(side="left", padx=4)
        combo.bind("<<ComboboxSelected>>", lambda e: self.render_rank())
        ttk.Label(opts, text="更新日:").pack(side="left", padx=(12, 0))
        self.rank_old = tk.StringVar(value=NO_FILTER)
        old_combo = ttk.Combobox(opts, textvariable=self.rank_old, width=10, state="readonly",
                                 values=[NO_FILTER] + [f"{y}年以上前" for y in OLD_YEARS])
        old_combo.pack(side="left", padx=4)
        old_combo.bind("<<ComboboxSelected>>", lambda e: self.render_rank())
        self.rank_note = tk.StringVar()
        # 説明は長くなるので、選択欄の下に 1 行とって表示する
        ttk.Label(tab, textvariable=self.rank_note, foreground="#777", padding=(2, 0, 2, 4)).pack(fill="x")

        self.rank_tree = self._make_tree(tab, [
            ("rank", "順位", 45, "e"),
            ("name", "名前", 260, "w"),
            ("size", "サイズ", 90, "e"),
            ("count", "ファイル数", 80, "e"),
            ("mtime", "更新日時", 130, "w"),
            ("location", "場所", 320, "w"),
        ])
        self.rank_tree.bind("<Double-1>", self.on_rank_double_click)
        self.rank_tree.bind("<Return>", self.on_rank_double_click)

    # ---------- パス指定・スキャン ----------
    def browse_folder(self):
        path = filedialog.askdirectory(title="フォルダを選択(ネットワークサーバーも指定可)")
        if path:
            self.open_path(path)

    def open_path(self, path):
        path = path.strip().strip('"')
        if not path:
            return
        path = os.path.normpath(path)
        # つながらないネットワークパスは確認だけで数十秒かかることがあるので、
        # 画面が固まらないよう別スレッドで確かめる。結果は _on_path_checked に届く
        self.check_id += 1
        self.status_var.set(f"フォルダを確認しています... {path}")
        threading.Thread(target=self._check_worker, args=(path, self.check_id), daemon=True).start()

    def _check_worker(self, path, check_id):
        """別スレッドで動く。パスがフォルダかどうかを確かめる"""
        self.msg_queue.put(("checked", check_id, (path, os.path.isdir(path))))

    def _on_path_checked(self, check_id, payload):
        if check_id != self.check_id:
            return  # そのあとに別のパスが指定された
        path, ok = payload
        if not ok:
            if self.result:
                self._update_status()
            else:
                self.status_var.set("フォルダが見つかりませんでした")
            messagebox.showerror(
                "エラー",
                f"フォルダが見つかりません:\n{path}\n\n"
                "パスが正しいか、ネットワークにつながっているかを確かめてください。")
            return
        self.root_path = path
        self.path_var.set(path)
        self._start_scan(keep_folder=False)

    def rescan(self):
        if self.root_path:
            self._start_scan(keep_folder=True)

    def cancel_scan(self):
        self.cancel_flag.set()

    def _start_scan(self, keep_folder):
        # 前のスキャンが動いていれば止め、新しいスキャン用の中止フラグを作る
        self.cancel_flag.set()
        self.cancel_flag = threading.Event()
        self.scan_id += 1
        self.result = None
        self.scanning = True
        self._update_buttons()
        self.status_var.set("スキャン中...")

        # スキャンを待たずに中身を表示する(フォルダの容量は「計算中…」)。
        # 再スキャンのときは今いるフォルダと選んでいる項目をそのまま保つ
        # (今いたフォルダが消えていたら、読み込みの失敗後に一番上へ戻る)
        if not keep_folder or not self.current:
            self.current = self.root_path
        self.show_folder(self.current, select_path=self._folder_selection() if keep_folder else None)
        self.render_rank()

        threading.Thread(
            target=self._scan_worker,
            args=(self.root_path, self.scan_id, self.cancel_flag),
            daemon=True,
        ).start()

    def _scan_worker(self, path, scan_id, cancel_flag):
        """別スレッドで動く。結果は msg_queue 経由でメインスレッドに渡す"""
        def progress(files, dirs):
            self.msg_queue.put(("progress", scan_id, (files, dirs)))
        try:
            result = scan_tree(path, cancel_flag, progress)
        except Exception as e:  # 想定外のエラーでも画面が固まらないようにする
            self.msg_queue.put(("error", scan_id, str(e)))
            return
        if result is None:
            self.msg_queue.put(("cancelled", scan_id, None))
        else:
            self.msg_queue.put(("done", scan_id, result))

    def _poll_queue(self):
        """別スレッドからの知らせを受け取って画面に反映する(0.1 秒ごと)"""
        try:
            while True:
                kind, msg_id, payload = self.msg_queue.get_nowait()
                if kind == "folder":
                    self._on_folder_loaded(msg_id, payload)
                    continue
                if kind == "checked":
                    self._on_path_checked(msg_id, payload)
                    continue
                if msg_id != self.scan_id:
                    continue  # 中止済みの古いスキャンからの結果は捨てる
                if kind == "progress":
                    files, dirs = payload
                    self.status_var.set(f"スキャン中... ファイル {files:,} 件 / フォルダ {dirs:,} 件")
                elif kind == "done":
                    self.result = payload
                    self.scanning = False
                    self._refresh_folder_sizes()
                    self.render_rank()
                    self._update_buttons()
                    self._update_status()
                elif kind == "error":
                    self.scanning = False
                    self._update_buttons()
                    self.status_var.set("スキャンに失敗しました")
                    messagebox.showerror("エラー", f"スキャンに失敗しました:\n{payload}")
                elif kind == "cancelled":
                    self.scanning = False
                    self._update_buttons()
                    self._refresh_folder_sizes()
                    self.render_rank()
                    self.status_var.set("スキャンを中止しました(フォルダの容量は未計算です。「再スキャン」で取り直せます)")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _update_buttons(self):
        has_root = self.root_path is not None
        self.up_btn.config(state="normal" if has_root and self.current != self.root_path else "disabled")
        self.rescan_btn.config(state="normal" if has_root and not self.scanning else "disabled")
        self.csv_btn.config(state="normal" if has_root else "disabled")
        self.explorer_btn.config(state="normal" if has_root else "disabled")
        if self.scanning:
            self.cancel_btn.pack(side="left", padx=4)
        else:
            self.cancel_btn.pack_forget()

    def _update_status(self):
        r = self.result
        if not r:
            return
        err = f" / 読めなかった項目: {r.error_count:,}件" if r.error_count else ""
        self.status_var.set(
            f"{r.root}   合計: {format_bytes(r.dir_sizes.get(r.root, 0))}"
            f" / ファイル {r.file_count:,} 件 / フォルダ {len(r.dir_sizes) - 1:,} 件{err}"
        )

    # ---------- 「フォルダの中身」タブ ----------
    def folder_size(self, path):
        """スキャン結果からフォルダの合計容量を取り出す(未計算なら None)"""
        return self.result.dir_sizes.get(path) if self.result else None

    def _folder_selection(self):
        """「フォルダの中身」タブで選ばれている項目のパス(なければ None)"""
        sel = self.folder_tree.selection()
        return self.row_paths[sel[0]][0] if sel and sel[0] in self.row_paths else None

    def show_folder(self, path, select_path=None):
        """フォルダの中身を別スレッドで読み込む。読み終わると _on_folder_loaded で表示する。
        ネットワークが遅いと 1 つのフォルダを読むだけでも時間がかかり、
        画面が固まってしまうため、別スレッドにしている。"""
        self.folder_req += 1
        self.loading_path = path
        self.folder_label.set(f"読み込み中... {path}")
        threading.Thread(target=self._folder_worker, args=(path, select_path, self.folder_req),
                         daemon=True).start()

    def _folder_worker(self, path, select_path, req):
        """別スレッドで動く。フォルダ直下の一覧を読む"""
        try:
            items, err = list_folder(path), None
        except OSError as e:
            items, err = None, e
        self.msg_queue.put(("folder", req, (path, select_path, items, err)))

    def _on_folder_loaded(self, req, payload):
        if req != self.folder_req:
            return  # 読み込み中に別のフォルダを開いたので、この結果は使わない
        path, select_path, items, err = payload
        self.loading_path = None
        if err is not None:
            messagebox.showerror(
                "エラー",
                f"フォルダを開けませんでした:\n{path}\n\n{err}\n\n"
                "アクセス権があるか、ネットワークにつながっているかを確かめてください。")
            if path == self.current and path != self.root_path:
                # 再スキャンのときに今いたフォルダが消えていた場合は、一番上に戻る
                self.show_folder(self.root_path)
            else:
                if path == self.current:
                    self.folder_items = []
                self._render_folder()
            return
        self.current = path
        self.path_var.set(path)
        self.folder_items = items
        self._render_folder(select_path)
        self._update_buttons()

    def _refresh_folder_sizes(self):
        """スキャンが終わったとき、今の一覧にフォルダの容量を反映する(読み直しはしない)"""
        if self.loading_path is None and self.current:
            self._render_folder(self._folder_selection())

    def _render_folder(self, select_path=None):
        tree = self.folder_tree
        tree.delete(*tree.get_children())
        self.row_paths = {}

        # 各項目の容量(フォルダはスキャン結果から)
        rows = []
        for name, path, is_dir, fsize, mtime in self.folder_items:
            size = self.folder_size(path) if is_dir else fsize
            rows.append((name, path, is_dir, size, mtime))
        total = self.folder_size(self.current)
        if total is None:
            total = sum(r[3] or 0 for r in rows)

        # 並べ替え。名前・種類・日時のときはエクスプローラーと同じくフォルダを先にする
        col, desc = self.sort_col, self.sort_desc
        if col == "size":
            rows.sort(key=lambda r: (r[3] if r[3] is not None else -1, natural_key(r[0])), reverse=desc)
        else:
            if col == "name":
                keyf = lambda r: natural_key(r[0])
            elif col == "type":
                keyf = lambda r: (type_text(r[0], r[2]), natural_key(r[0]))
            else:  # mtime
                keyf = lambda r: (r[4] or 0, natural_key(r[0]))
            rows.sort(key=keyf, reverse=desc)
            rows.sort(key=lambda r: not r[2])  # 安定ソートなのでフォルダが先に集まる

        self.folder_rows = rows
        pending = "計算中…" if self.scanning else "未計算"
        select_iid = None
        for name, path, is_dir, size, mtime in rows:
            if size is None:
                size_text, bar = (pending if is_dir else ""), ""
            else:
                size_text = format_bytes(size)
                pct = (size / total * 100) if total else 0
                filled = round(pct / 10)
                bar = f"{pct:3.0f}%  " + "■" * filled + "□" * (10 - filled)
            label = ("📁 " if is_dir else "📄 ") + name
            iid = tree.insert("", "end", values=(label, format_time(mtime), type_text(name, is_dir), size_text, bar),
                              tags=("dir" if is_dir else "file",))
            self.row_paths[iid] = (path, is_dir)
            if select_path and path == select_path:
                select_iid = iid
        if select_iid:
            tree.selection_set(select_iid)
            tree.focus(select_iid)
            tree.see(select_iid)

        # 今いる場所を、スキャンしたパスからの相対で表示
        rel = os.path.relpath(self.current, self.root_path) if self.root_path else self.current
        place = "(一番上)" if rel == "." else rel
        size_text = format_bytes(total) if self.folder_size(self.current) is not None else pending
        self.folder_label.set(f"場所: {place}   このフォルダの合計: {size_text}   項目数: {len(rows):,}")

    def sort_folder(self, col):
        if self.sort_col == col:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = col
            # サイズと更新日時は大きい順・新しい順から始めるほうが便利
            self.sort_desc = col in ("size", "mtime")
        self._update_sort_headings()
        if self.current:
            self._render_folder()

    def _update_sort_headings(self):
        for cid, text, _w, _a in self.folder_cols:
            mark = (" ▼" if self.sort_desc else " ▲") if cid == self.sort_col else ""
            self.folder_tree.heading(cid, text=text + mark)

    def on_folder_double_click(self, _event=None):
        item = self.folder_tree.focus()
        if not item or item not in self.row_paths:
            return
        path, is_dir = self.row_paths[item]
        if is_dir:
            self.show_folder(path)

    def go_up(self):
        # 読み込み中なら、その読み込み中のフォルダから 1 つ上へ
        child = self.loading_path or self.current
        if not child or child == self.root_path:
            return
        self.show_folder(os.path.dirname(child), select_path=child)

    # ---------- 「容量ランキング」タブ ----------
    def render_rank(self):
        tree = self.rank_tree
        tree.delete(*tree.get_children())
        self.rank_paths = {}
        r = self.result
        if not r:
            self.rank_note.set("スキャン中です…" if self.scanning else "スキャンが終わると表示されます")
            return
        limit = int(self.rank_limit.get())
        is_dir = self.rank_kind.get() == "dir"
        # 「更新日」の絞り込み(「3年以上前」なら 3、「指定しない」なら None)
        old = self.rank_old.get()
        years = None if old == NO_FILTER else int(old.split("年")[0])
        old_note = f"更新日が{years}年以上前のものだけを表示しています。" if years else ""
        if is_dir:
            # フォルダの「更新日時」は、中のファイルで一番新しいものを表示する。
            # フォルダ自体の更新日時は直下の出し入れでしか変わらず、整理の目安にならないため
            cutoff = r.scanned_at - years * YEAR_SECONDS if years else None
            items = []
            for path, size in r.dir_sizes.items():
                if path == r.root:
                    continue  # 一番上のフォルダ自身はランキングから外す
                newest = r.dir_newest.get(path)
                if cutoff is not None and (newest is None or newest >= cutoff):
                    continue
                items.append((size, path, newest))
            items = heapq.nlargest(limit, items)
            self.rank_tree.heading("mtime", text="中の最終更新")
            self.rank_note.set("※ フォルダの容量には、その中のフォルダの分も含まれます。"
                               "「中の最終更新」は、中のファイルで一番新しい更新日時です。" + old_note)
        else:
            items = (r.old_files.get(years, []) if years else r.top_files)[:limit]
            self.rank_tree.heading("mtime", text="更新日時")
            self.rank_note.set(old_note)

        for i, (size, path, mtime) in enumerate(items, start=1):
            parent_rel = os.path.relpath(os.path.dirname(path), r.root)
            location = "(一番上)" if parent_rel == "." else parent_rel
            count = r.dir_counts.get(path, 0) if is_dir else None
            iid = tree.insert("", "end", values=(i, os.path.basename(path), format_bytes(size),
                                                 "" if count is None else f"{count:,}",
                                                 format_time(mtime), location))
            self.rank_paths[iid] = (path, is_dir, size, count)

    def on_rank_double_click(self, _event=None):
        """ランキングの項目を「フォルダの中身」タブで開く"""
        item = self.rank_tree.focus()
        if not item or item not in self.rank_paths:
            return
        path, is_dir = self.rank_paths[item][:2]
        if is_dir:
            self.show_folder(path)
        else:
            self.show_folder(os.path.dirname(path), select_path=path)
        self.notebook.select(0)
        self.folder_tree.focus_set()

    # ---------- エクスプローラー連携 ----------
    def _selected_path(self):
        """表示中のタブで選ばれている項目のパス(なければ None)"""
        if self.notebook.index("current") == 0:
            tree, mapping = self.folder_tree, self.row_paths
        else:
            tree, mapping = self.rank_tree, self.rank_paths
        sel = tree.selection()
        return mapping[sel[0]][0] if sel and sel[0] in mapping else None

    def show_in_explorer(self):
        """選んだ項目をエクスプローラーで表示する(何も選んでいなければ今のフォルダを開く)"""
        path = self._selected_path()
        try:
            if path:
                subprocess.Popen(f'explorer /select,"{path}"')
            elif self.current:
                os.startfile(self.current)
        except Exception as e:
            messagebox.showerror("エラー", f"エクスプローラーを開けませんでした:\n{e}")

    # ---------- CSV出力 ----------
    def export_csv(self):
        """表示中のタブの一覧をCSVに保存する"""
        on_folder_tab = self.notebook.index("current") == 0
        default = "フォルダの中身.csv" if on_folder_tab else "容量ランキング.csv"
        save_path = filedialog.asksaveasfilename(
            title="CSVとして保存", defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv")], initialfile=default,
        )
        if not save_path:
            return
        try:
            # UTF-8 BOM 付きにすると Excel で開いても文字化けしない
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                if on_folder_tab:
                    writer.writerow(["名前", "種類", "サイズ(バイト)", "更新日時", "フルパス"])
                    for name, path, is_dir, size, mtime in self.folder_rows:
                        writer.writerow([name, type_text(name, is_dir), "" if size is None else size,
                                         format_time(mtime), path])
                else:
                    mtime_head = "中の最終更新日時" if self.rank_kind.get() == "dir" else "更新日時"
                    writer.writerow(["順位", "名前", "種類", "サイズ(バイト)", "ファイル数", mtime_head, "フルパス"])
                    for iid in self.rank_tree.get_children():
                        path, is_dir, size, count = self.rank_paths[iid]
                        rank, name, _s, _c, mtime, _loc = self.rank_tree.item(iid, "values")
                        writer.writerow([rank, name, "フォルダ" if is_dir else "ファイル", size,
                                         "" if count is None else count, mtime, path])
            messagebox.showinfo("保存完了", f"保存しました:\n{save_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました:\n{e}")


def main():
    root = tk.Tk()
    FolderSizeViewer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
