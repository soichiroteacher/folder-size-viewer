# -*- coding: utf-8 -*-
"""
フォルダ容量ビューア
====================
指定したフォルダ(ローカル or ネットワークサーバーの共有フォルダ)の中身を
エクスプローラーのように表示し、あわせて容量を表示するツールです。

【画面の構成】
  ・上の段 … パス欄・「参照...」「開く」・「見本で試す」・「❓ 使い方」
  ・ボタンの段 … 「↑ 上のフォルダへ」「再スキャン」「CSVで保存」
                 「エクスプローラーで表示」「スキャン中止」
  ・はじめの画面 … まだフォルダを開いていないときに「はじめての方へ」を表示
  ・「フォルダの中身」タブ
      エクスプローラーと同じように、今いるフォルダの中のフォルダ・ファイルを
      表示します(フォルダが先、名前順)。フォルダにも中身の合計容量を表示します。
      列の見出しをクリックすると並べ替えられます(「サイズ」で容量の大きい順)。
      フォルダをダブルクリックすると中に入り、「↑ 上のフォルダへ」で戻ります。
  ・「ツリー表示」タブ
      指定したパス以下を、フォルダの階層(ツリー)で表示します。
      ▶ をクリックするとフォルダが開きます。中身は容量の大きい順です。
  ・「容量ランキング」タブ
      指定したパス以下にあるすべてのファイル(またはフォルダ)を、
      容量の大きい順に表示します。ダブルクリックすると、その場所を
      「フォルダの中身」タブで開きます。
      「更新日」で「3年以上前」などを選ぶと、長く更新されていない
      大きいファイル(フォルダ)だけに絞り込めます。

【CSVで保存】
  「フォルダの中身」タブでは今いるフォルダ、「ツリー表示」タブでは一番上の
  フォルダについて、その下のすべてのフォルダ・ファイルをツリーの順に保存します。
  「容量ランキング」タブでは、表示中のランキングを保存します。

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
  - スキャン結果には、全フォルダの合計容量に加えて、全フォルダ・全ファイルの
    一覧(名前・容量・更新日時)も持ちます。「ツリー表示」と「CSVで保存」で
    使うためです。ネットワークを 2 回読みに行くと遅いので、1 回のスキャンで
    覚えておく方式にしました。そのぶんメモリを使います
    (ファイル 10 万件でおよそ 20〜30MB)。
  - 「フォルダの中身」タブは、表示のたびにそのフォルダだけを読み直し、
    フォルダの容量はスキャン結果から取り出して表示しています。
  - フォルダの読み直しや、入力されたパスの確認も別スレッドで行います。
    つながらないネットワークパスは、確認だけで数十秒かかることがあるためです。
    別スレッドの結果はすべて msg_queue で画面側(_poll_queue)に届けます。
  - 「見本で試す」は、一時フォルダ(%TEMP%)に架空のフォルダを作って開きます。
    中身は SAMPLE_FILES で決めています。毎回同じ内容に上書きするだけで、
    消すことはしません(このツールには削除の処理を入れない方針のため)。
  - 画面ガイド(「❓ 使い方」)の手順は _guide_steps() にあります。
    画面の部品(ボタンの名前や位置)を変えたら、ここの説明も必ず直してください。
  - .exe の作り方は build_exe.bat と HANDOFF.md を参照してください。
"""

import csv
import heapq
import os
import queue
import re
import subprocess
import tempfile
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

# ====================================================================
# 設定(変えたいときはここを直す)
# ====================================================================
APP_TITLE = "フォルダ容量ビューア"
# 文字の大きさ(ポイント)。パソコンが苦手な人にも読みやすいよう大きめにしている
FONT_SIZE = 11
# 容量ランキングで覚えておくファイルの最大件数
TOP_FILES = 1000
# ランキングの表示件数の選択肢
RANK_LIMITS = ("100", "500", "1000")
# 「更新日」で絞り込むときの選択肢(〇年以上前)
OLD_YEARS = (1, 3, 5)
NO_FILTER = "指定しない"
YEAR_SECONDS = 365.25 * 24 * 60 * 60
# Excel で開ける行数の上限(これを超える CSV は Excel では最後まで開けない)
EXCEL_MAX_ROWS = 1048576

# 「見本で試す」で一時フォルダに作る架空のフォルダ
SAMPLE_PARENT = "フォルダ容量ビューア_見本"
SAMPLE_ROOT = "見本_共有フォルダ"
# (場所, 容量KB, 何年前に更新したことにするか)。実在の学校名・人名は使わない
SAMPLE_FILES = [
    ("はじめにお読みください.txt", 2, 0.1),
    ("01_教務/年間行事予定.xlsx", 48, 0.3),
    ("01_教務/時間割_前期.xlsx", 62, 0.5),
    ("01_教務/会議資料/職員会議_4月.docx", 35, 0.5),
    ("01_教務/会議資料/職員会議_5月.docx", 38, 0.4),
    ("01_教務/会議資料/古い会議資料/職員会議_2020年度.pdf", 900, 5.5),
    ("02_学年/1学年/学年通信_第1号.docx", 120, 0.4),
    ("02_学年/1学年/学年通信_第2号.docx", 130, 0.2),
    ("02_学年/2学年/学年通信_第1号.docx", 110, 1.3),
    ("03_行事/体育祭_2020/動画_リレー.mp4", 2400, 6.0),
    ("03_行事/合唱コンクール_2022/動画_合唱.mp4", 1800, 3.8),
    ("03_行事/文化祭_2025/ポスター.pdf", 700, 1.0),
    ("04_研修/研修資料_情報モラル.pdf", 520, 2.2),
    ("04_研修/研修資料_授業づくり.pptx", 1500, 0.8),
    ("整理前の古いデータ/バックアップ_2018.zip", 3000, 7.5),
    ("整理前の古いデータ/旧ホームページ/index.html", 12, 8.0),
    ("整理前の古いデータ/旧ホームページ/images/top.jpg", 400, 8.0),
] + [
    (f"02_学年/3学年/修学旅行/写真_{i:03}.jpg", 120 + i * 3, 1.1) for i in range(1, 25)
] + [
    (f"03_行事/体育祭_2020/写真_{i:03}.jpg", 150 + i * 2, 6.0) for i in range(1, 16)
]


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


def percent_bar(size, total):
    """割合の列に出す「 35%  ■■■■□□□□□□」のような文字列"""
    pct = (size / total * 100) if total else 0
    filled = round(pct / 10)
    return f"{pct:3.0f}%  " + "■" * filled + "□" * (10 - filled)


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


def make_sample_folder():
    """「見本で試す」用の架空のフォルダを一時フォルダに作り、そのパスを返す。
    すでにあれば同じ内容で上書きする(消すことはしない)。
    ファイルの中身は空白(0 の並び)で、容量と更新日時だけを見本らしくしている。"""
    base = os.path.join(tempfile.gettempdir(), SAMPLE_PARENT, SAMPLE_ROOT)
    now = time.time()
    for rel, kb, years in SAMPLE_FILES:
        path = os.path.join(base, *rel.split("/"))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.truncate(kb * 1024)
        t = now - years * YEAR_SECONDS
        os.utime(path, (t, t))
    return base


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
        self.subdirs = {}      # フォルダのパス -> 直下のフォルダのパスの一覧
        self.files = {}        # フォルダのパス -> 直下のファイルの一覧 [(名前, 容量, 更新日時), ...]
        self.top_files = []    # (容量, パス, 更新日時) 大きい順
        self.old_files = {}    # 年数 -> その年数以上更新されていないファイルの top_files
        self.error_count = 0   # 読めなかったフォルダ・ファイルの数
        self.file_count = 0
        self.scanned_at = time.time()


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
        subs = []
        files = []
        try:
            with os.scandir(d) as it:
                for e in it:
                    try:
                        if is_real_dir(e):
                            parent[e.path] = d
                            subs.append(e.path)
                            stack.append(e.path)
                        else:
                            st = e.stat(follow_symlinks=False)
                            size += st.st_size
                            count += 1
                            files.append((e.name, st.st_size, st.st_mtime))
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
        result.subdirs[d] = subs
        result.files[d] = files
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


def tree_children(result, path):
    """ツリー表示・CSV用に、フォルダ直下の項目を「フォルダ(容量の大きい順)→ファイル(容量の大きい順)」で返す。
    戻り値: [(名前, フルパス, フォルダか, 容量, 更新日時), ...]"""
    dirs = [(os.path.basename(p), p, True, result.dir_sizes.get(p, 0), result.dir_newest.get(p))
            for p in result.subdirs.get(path, [])]
    files = [(name, os.path.join(path, name), False, size, mtime)
             for name, size, mtime in result.files.get(path, [])]
    dirs.sort(key=lambda x: (-x[3], natural_key(x[0])))
    files.sort(key=lambda x: (-x[3], natural_key(x[0])))
    return dirs + files


def write_tree_csv(f, result, start):
    """start 以下のすべてのフォルダ・ファイルを、ツリーの順に CSV へ書く。書いた行数(見出しを除く)を返す"""
    writer = csv.writer(f)
    writer.writerow(["階層", "名前", "種類", "サイズ(バイト)", "サイズ", "割合(親フォルダに対して)",
                     "ファイル数", "更新日時", "フルパス"])
    rows = 0
    # 再帰を使わず、積み上げ(スタック)でたどる。階層がとても深くても止まらないようにするため
    stack = [(os.path.basename(start) or start, start, True, result.dir_sizes.get(start, 0),
              result.dir_newest.get(start), 0, None)]
    while stack:
        name, path, is_dir, size, mtime, depth, parent_size = stack.pop()
        pct = "" if parent_size is None else (f"{size / parent_size * 100:.1f}%" if parent_size else "0.0%")
        writer.writerow([depth, "　" * depth + name, "フォルダ" if is_dir else type_text(name, False),
                         size, format_bytes(size), pct,
                         result.dir_counts.get(path, 0) if is_dir else "", format_time(mtime), path])
        rows += 1
        if is_dir:
            # 後に積んだものが先に出るので、逆順に積む
            for child in reversed(tree_children(result, path)):
                stack.append(child + (depth + 1, size))
    return rows


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
# 画面ガイド(「❓ 使い方」)
# ====================================================================
class ScreenGuide:
    """画面の部品を 1 つずつ黄色の枠で囲み、そばに説明の小窓を出す案内。
    steps は (対象の部品を返す関数 or None, 見出し, 説明, 表示前に行う処理 or None) の並び。"""

    BORDER = 4
    COLOR = "#FFC107"

    def __init__(self, root, steps, on_end=None):
        self.root = root
        self.steps = steps
        self.on_end = on_end
        self.index = 0
        # 枠は、線の形をした小さな窓を 4 つ(上下左右)並べて作る
        self.edges = []
        for _ in range(4):
            w = tk.Toplevel(root)
            w.overrideredirect(True)
            w.configure(bg=self.COLOR)
            w.attributes("-topmost", True)
            self.edges.append(w)
        # 説明の小窓も題名の帯なしの窓にする。題名の帯がある普通の窓だと、
        # Windows が位置を勝手に戻してしまい、説明する部品のそばに置けないことがあるため。
        # (閉じるときは「終わる」ボタンか Esc キー)
        self.card = tk.Toplevel(root)
        self.card.overrideredirect(True)
        self.card.attributes("-topmost", True)
        self.card.configure(bg=self.COLOR, padx=3, pady=3)  # 黄色の縁取り
        body = ttk.Frame(self.card, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="画面ガイド", foreground="#777").pack(anchor="w")
        self.title_var = tk.StringVar()
        self.text_var = tk.StringVar()
        ttk.Label(body, textvariable=self.title_var, font=("", FONT_SIZE + 3, "bold")).pack(anchor="w")
        ttk.Label(body, textvariable=self.text_var, wraplength=480, justify="left").pack(anchor="w", pady=(6, 12))
        foot = ttk.Frame(body)
        foot.pack(fill="x")
        self.count_var = tk.StringVar()
        ttk.Label(foot, textvariable=self.count_var, foreground="#777").pack(side="left")
        ttk.Button(foot, text="終わる", command=self.end).pack(side="right")
        self.next_btn = ttk.Button(foot, text="次へ ▶", command=lambda: self.show(self.index + 1))
        self.next_btn.pack(side="right", padx=4)
        self.back_btn = ttk.Button(foot, text="◀ 戻る", command=lambda: self.show(self.index - 1))
        self.back_btn.pack(side="right")
        self.card.bind("<Right>", lambda e: self.show(self.index + 1))
        self.card.bind("<Left>", lambda e: self.show(self.index - 1))
        self.card.bind("<Escape>", lambda e: self.end())
        # 画面を動かしたり大きさを変えたりしたら、枠と小窓もついていく
        # (部品の大きさが変わるたびに届くので、まとめて 1 回だけ置き直す)
        self.place_pending = False
        self.bind_id = root.bind("<Configure>", self._on_configure, add="+")
        self.alive = True
        self.show(0)

    def show(self, i):
        if not self.alive:
            return
        if i >= len(self.steps):
            self.end()
            return
        self.index = max(0, i)
        _target, title, text, before = self.steps[self.index]
        if before:
            before()
        self.title_var.set(title)
        self.text_var.set(text)
        self.count_var.set(f"{self.index + 1} / {len(self.steps)}")
        self.back_btn.config(state="normal" if self.index > 0 else "disabled")
        self.next_btn.config(text="終わる" if self.index == len(self.steps) - 1 else "次へ ▶")
        # 部品の位置が決まってから枠を置く
        self.root.update_idletasks()
        self.place()
        self.card.focus_force()

    def _on_configure(self, _event=None):
        if not self.place_pending:
            self.place_pending = True
            self.root.after_idle(self.place)

    def place(self):
        self.place_pending = False
        if not self.alive:
            return
        target_fn = self.steps[self.index][0]
        target = target_fn() if target_fn else None
        self.card.update_idletasks()
        cw, ch = self.card.winfo_reqwidth(), self.card.winfo_reqheight()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        rects = None
        if target is not None and target.winfo_ismapped():
            x, y = target.winfo_rootx(), target.winfo_rooty()
            w, h = target.winfo_width(), target.winfo_height()
            b = self.BORDER
            rects = [(x - b, y - b, w + 2 * b, b), (x - b, y + h, w + 2 * b, b),
                     (x - b, y, b, h), (x + w, y, b, h)]
            # 小窓は部品の下に置く。下に入らなければ上、それも無理なら部品の中に重ねる
            cx = min(max(x, 0), sw - cw - 10)
            if h > 250:
                # タブ全体のような大きい部品は、上の行が隠れないよう右下に重ねる
                cx = min(x + w - cw - 30, sw - cw - 10)
                cy = y + h - ch - 30
            elif y + h + b + 10 + ch < sh - 40:
                cy = y + h + b + 10
            elif y - b - 10 - ch > 0:
                cy = y - b - 10 - ch
            else:
                cy = y + 20
        else:
            cx = self.root.winfo_rootx() + (self.root.winfo_width() - cw) // 2
            cy = self.root.winfo_rooty() + (self.root.winfo_height() - ch) // 3
        # 小窓を先に動かしてから枠を置く。枠を先に置くと、Windows では
        # 小窓の移動が効かなくなることがあるため(試して確かめた)
        self.card.geometry(f"{cw}x{ch}+{int(cx)}+{int(cy)}")
        self.card.update_idletasks()
        if rects:
            for edge, (ex, ey, ew, eh) in zip(self.edges, rects):
                edge.geometry(f"{max(ew, 1)}x{max(eh, 1)}+{ex}+{ey}")
                edge.deiconify()
                edge.lift()
        else:
            for edge in self.edges:
                edge.withdraw()
        self.card.lift()

    def end(self):
        if not self.alive:
            return
        self.alive = False
        try:
            self.root.unbind("<Configure>", self.bind_id)
        except tk.TclError:
            pass
        for w in self.edges + [self.card]:
            w.destroy()
        if self.on_end:
            self.on_end()


# ====================================================================
# 画面
# ====================================================================
class FolderSizeViewer:
    TAB_FOLDER, TAB_TREE, TAB_RANK = 0, 1, 2

    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("1100x700")
        root.minsize(820, 480)
        self._setup_fonts()

        self.root_path = None   # スキャンしたパス(一番上)
        self.current = None     # 「フォルダの中身」タブで今見ているフォルダ
        self.result = None      # ScanResult(スキャン中・中止時は None)
        self.scanning = False
        self.is_sample = False  # 「見本で試す」で開いているか
        self.guide = None       # 表示中の画面ガイド
        self.folder_items = []  # 今のフォルダの中身(list_folder の戻り値)
        self.folder_rows = []   # 上を容量付き・画面の並び順にしたもの
        self.row_paths = {}     # 「フォルダの中身」の行ID -> (フルパス, フォルダか)
        self.tree_paths = {}    # 「ツリー表示」の行ID -> (フルパス, フォルダか)
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

    def _setup_fonts(self):
        """文字を大きめにする(共通ルール「文字は大きめ」)"""
        for name in ("TkDefaultFont", "TkTextFont", "TkHeadingFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(size=FONT_SIZE)
            except tk.TclError:
                pass
        style = ttk.Style(self.root)
        # 行の高さも文字に合わせて広げる
        style.configure("Treeview", rowheight=int(FONT_SIZE * 2.4))
        style.configure("Big.TButton", padding=(14, 8))

    # ---------- UI構築 ----------
    def _build_ui(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="パス:").pack(side="left")
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(top, textvariable=self.path_var)
        self.path_entry.pack(side="left", padx=4, fill="x", expand=True)
        self.path_entry.bind("<Return>", lambda e: self.open_path(self.path_var.get()))
        self.browse_btn = ttk.Button(top, text="参照...", command=self.browse_folder)
        self.browse_btn.pack(side="left", padx=2)
        ttk.Button(top, text="開く", command=lambda: self.open_path(self.path_var.get())).pack(side="left", padx=2)
        self.help_btn = ttk.Button(top, text="❓ 使い方", command=self.start_guide)
        self.help_btn.pack(side="right", padx=(8, 0))
        self.sample_btn = ttk.Button(top, text="見本で試す", command=self.start_sample)
        self.sample_btn.pack(side="right", padx=(12, 0))

        self.toolbar = ttk.Frame(self.root, padding=(8, 0, 8, 6))
        self.toolbar.pack(fill="x")
        self.up_btn = ttk.Button(self.toolbar, text="↑ 上のフォルダへ", command=self.go_up, state="disabled")
        self.up_btn.pack(side="left")
        self.rescan_btn = ttk.Button(self.toolbar, text="再スキャン", command=self.rescan, state="disabled")
        self.rescan_btn.pack(side="left", padx=4)
        self.csv_btn = ttk.Button(self.toolbar, text="CSVで保存", command=self.export_csv, state="disabled")
        self.csv_btn.pack(side="left", padx=4)
        self.explorer_btn = ttk.Button(self.toolbar, text="エクスプローラーで表示",
                                       command=self.show_in_explorer, state="disabled")
        self.explorer_btn.pack(side="left", padx=4)
        self.cancel_btn = ttk.Button(self.toolbar, text="スキャン中止", command=self.cancel_scan)

        self.status_var = tk.StringVar(value="パスを入力するか「参照...」でフォルダを選んでください")
        ttk.Label(self.root, textvariable=self.status_var, padding=(8, 0, 8, 4),
                  foreground="#555").pack(fill="x")

        self.body = ttk.Frame(self.root)
        self.body.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.notebook = ttk.Notebook(self.body)
        self._build_folder_tab()
        self._build_tree_tab()
        self._build_rank_tab()
        self._build_welcome()
        self.welcome.pack(fill="both", expand=True)

    def _build_welcome(self):
        """まだフォルダを開いていないときに出す「はじめての方へ」の画面"""
        self.welcome = ttk.Frame(self.body, padding=30)
        ttk.Label(self.welcome, text="はじめての方へ", font=("", FONT_SIZE + 6, "bold")).pack(anchor="w")
        text = (
            "このアプリは、フォルダの中にあるフォルダやファイルが、どれくらいの容量を使っているかを調べます。\n"
            "容量の大きいものや、長く使っていないものを見つけて、整理に役立ててください。\n"
            "このアプリでファイルを消すことはできません(安心してお使いください)。\n\n"
            "使い方\n"
            "  1. 上の「参照...」で調べたいフォルダを選びます(パス欄に入力して Enter でも開けます)。\n"
            "  2. 「フォルダの中身」「ツリー表示」「容量ランキング」のタブで、容量の大きいものを探します。\n"
            "  3. 記録を残したいときは「CSVで保存」、中身を確かめたいときは「エクスプローラーで表示」を押します。\n\n"
            "まずは見本で試してみましょう。見本は架空のフォルダで、本物のデータには触れません。"
        )
        ttk.Label(self.welcome, text=text, justify="left").pack(anchor="w", pady=(10, 20))
        btns = ttk.Frame(self.welcome)
        btns.pack(anchor="w")
        ttk.Button(btns, text="見本で試す", style="Big.TButton", command=self.start_sample).pack(side="left")
        ttk.Button(btns, text="❓ 画面ガイドを見る", style="Big.TButton",
                   command=self.start_guide).pack(side="left", padx=12)

    def _show_tabs(self):
        """「はじめての方へ」を閉じて、タブの画面に切り替える"""
        if self.welcome.winfo_ismapped():
            self.welcome.pack_forget()
            self.notebook.pack(fill="both", expand=True)

    def _make_tree(self, parent, columns, tree_col=None):
        """スクロールバー付きの一覧(Treeview)を作る。columns: [(ID, 見出し, 幅, 寄せ), ...]
        tree_col に (見出し, 幅) を渡すと、先頭に階層(ツリー)の列を付ける"""
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns],
                            show="tree headings" if tree_col else "headings", selectmode="browse")
        if tree_col:
            tree.heading("#0", text=tree_col[0])
            tree.column("#0", width=tree_col[1], stretch=True)
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
            ("mtime", "更新日時", 140, "w"),
            ("type", "種類", 130, "w"),
            ("size", "サイズ", 95, "e"),
            ("bar", "割合", 200, "w"),
        ]
        self.folder_tree = self._make_tree(tab, self.folder_cols)
        for cid, _t, _w, _a in self.folder_cols:
            if cid != "bar":
                self.folder_tree.heading(cid, command=lambda c=cid: self.sort_folder(c))
        self.folder_tree.bind("<Double-1>", self.on_folder_double_click)
        self.folder_tree.bind("<Return>", self.on_folder_double_click)
        self.folder_tree.bind("<BackSpace>", lambda e: self.go_up())
        self._update_sort_headings()

    def _build_tree_tab(self):
        tab = ttk.Frame(self.notebook, padding=(0, 6, 0, 0))
        self.notebook.add(tab, text="ツリー表示")
        opts = ttk.Frame(tab, padding=(2, 0, 2, 4))
        opts.pack(fill="x")
        self.tree_show_files = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="ファイルも表示する", variable=self.tree_show_files,
                        command=self.render_tree).pack(side="left")
        ttk.Button(opts, text="すべて閉じる", command=self.render_tree).pack(side="left", padx=12)
        self.tree_note = tk.StringVar()
        ttk.Label(tab, textvariable=self.tree_note, foreground="#777", padding=(2, 0, 2, 4)).pack(fill="x")
        self.dir_tree = self._make_tree(tab, [
            ("size", "サイズ", 95, "e"),
            ("bar", "割合(親フォルダに対して)", 210, "w"),
            ("count", "ファイル数", 85, "e"),
            ("mtime", "更新日時", 140, "w"),
        ], tree_col=("名前", 420))
        self.dir_tree.bind("<<TreeviewOpen>>", self.on_tree_open)
        self.dir_tree.bind("<Double-1>", self.on_tree_double_click)

    def _build_rank_tab(self):
        tab = ttk.Frame(self.notebook, padding=(0, 6, 0, 0))
        self.notebook.add(tab, text="容量ランキング")
        opts = ttk.Frame(tab, padding=(2, 0, 2, 4))
        opts.pack(fill="x")
        self.rank_opts = opts
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
        self.old_combo = ttk.Combobox(opts, textvariable=self.rank_old, width=10, state="readonly",
                                      values=[NO_FILTER] + [f"{y}年以上前" for y in OLD_YEARS])
        self.old_combo.pack(side="left", padx=4)
        self.old_combo.bind("<<ComboboxSelected>>", lambda e: self.render_rank())
        self.rank_note = tk.StringVar()
        # 説明は長くなるので、選択欄の下に 1 行とって表示する
        ttk.Label(tab, textvariable=self.rank_note, foreground="#777", padding=(2, 0, 2, 4)).pack(fill="x")

        self.rank_tree = self._make_tree(tab, [
            ("rank", "順位", 50, "e"),
            ("name", "名前", 260, "w"),
            ("size", "サイズ", 95, "e"),
            ("count", "ファイル数", 85, "e"),
            ("mtime", "更新日時", 140, "w"),
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
        self._set_sample_mode(False)
        self.root_path = path
        self.path_var.set(path)
        self._start_scan(keep_folder=False)

    def start_sample(self):
        """「見本で試す」:一時フォルダに架空のフォルダを作って開く"""
        try:
            path = make_sample_folder()
        except OSError as e:
            messagebox.showerror(
                "エラー",
                f"見本のフォルダを作れませんでした:\n{e}\n\n"
                "パソコンの空き容量が少ないかもしれません。空きを作ってからもう一度お試しください。")
            return
        self.check_id += 1  # 確認中のパスがあれば、その結果は使わない
        self._set_sample_mode(True)
        self.root_path = path
        self.path_var.set(path)
        self._start_scan(keep_folder=False)

    def _set_sample_mode(self, on):
        self.is_sample = on
        self.root.title(f"{APP_TITLE} 【見本】" if on else APP_TITLE)

    def rescan(self):
        if self.root_path:
            self._start_scan(keep_folder=True)

    def cancel_scan(self):
        self.cancel_flag.set()

    def _start_scan(self, keep_folder):
        self._show_tabs()
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
        self.render_tree()
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
                    self.render_tree()
                    self.render_rank()
                    self._update_buttons()
                    self._update_status()
                elif kind == "error":
                    self.scanning = False
                    self._update_buttons()
                    self.status_var.set("スキャンに失敗しました")
                    messagebox.showerror(
                        "エラー",
                        f"スキャンに失敗しました:\n{payload}\n\n"
                        "ネットワークにつながっているかを確かめて、「再スキャン」を押してください。")
                elif kind == "cancelled":
                    self.scanning = False
                    self._update_buttons()
                    self._refresh_folder_sizes()
                    self.render_tree()
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
        sample = "【見本】 " if self.is_sample else ""
        self.status_var.set(
            f"{sample}{r.root}   合計: {format_bytes(r.dir_sizes.get(r.root, 0))}"
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
                size_text, bar = format_bytes(size), percent_bar(size, total)
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

    def open_in_folder_tab(self, path, is_dir):
        """指定した場所を「フォルダの中身」タブで開く(ファイルならそのフォルダを開いて選ぶ)"""
        if is_dir:
            self.show_folder(path)
        else:
            self.show_folder(os.path.dirname(path), select_path=path)
        self.notebook.select(self.TAB_FOLDER)
        self.folder_tree.focus_set()

    # ---------- 「ツリー表示」タブ ----------
    def render_tree(self):
        """ツリーを作り直す(一番上のフォルダだけ開いた状態にする)"""
        tree = self.dir_tree
        tree.delete(*tree.get_children())
        self.tree_paths = {}
        r = self.result
        if not r:
            self.tree_note.set("スキャン中です…" if self.scanning else "スキャンが終わると表示されます")
            return
        self.tree_note.set("▶ をクリックするとフォルダが開きます。中身は容量の大きい順です。"
                           "フォルダの更新日時は、中のファイルで一番新しいものです。")
        root_iid = self._insert_tree_node("", (r.root, r.root, True, r.dir_sizes.get(r.root, 0),
                                               r.dir_newest.get(r.root)), None)
        self._fill_tree_node(root_iid)
        tree.item(root_iid, open=True)

    def _insert_tree_node(self, parent_iid, item, parent_size):
        name, path, is_dir, size, mtime = item
        r = self.result
        bar = "" if parent_size is None else percent_bar(size, parent_size)
        count = f"{r.dir_counts.get(path, 0):,}" if is_dir else ""
        iid = self.dir_tree.insert(parent_iid, "end", text=("📁 " if is_dir else "📄 ") + name,
                                   values=(format_bytes(size), bar, count, format_time(mtime)))
        self.tree_paths[iid] = (path, is_dir)
        if is_dir and (r.subdirs.get(path) or (self.tree_show_files.get() and r.files.get(path))):
            # 中身は開いたときに入れる(全部を最初に入れると、大きなフォルダで時間がかかるため)。
            # ▶ を出すために、仮の行を 1 つ入れておく
            self.dir_tree.insert(iid, "end", text="読み込み中…", tags=("placeholder",))
        return iid

    def _fill_tree_node(self, iid):
        tree = self.dir_tree
        children = tree.get_children(iid)
        if not (len(children) == 1 and "placeholder" in tree.item(children[0], "tags")):
            return  # もう中身を入れてある
        tree.delete(children[0])
        path = self.tree_paths[iid][0]
        size = self.result.dir_sizes.get(path, 0)
        show_files = self.tree_show_files.get()
        for item in tree_children(self.result, path):
            if item[2] or show_files:
                self._insert_tree_node(iid, item, size)

    def on_tree_open(self, _event=None):
        iid = self.dir_tree.focus()
        if iid in self.tree_paths:
            self._fill_tree_node(iid)

    def on_tree_double_click(self, _event=None):
        """ファイルをダブルクリックしたら、その場所を「フォルダの中身」タブで開く
        (フォルダのダブルクリックは、開く・閉じるのまま)"""
        iid = self.dir_tree.focus()
        if iid in self.tree_paths and not self.tree_paths[iid][1]:
            self.open_in_folder_tab(*self.tree_paths[iid])

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
        self.open_in_folder_tab(path, is_dir)

    # ---------- エクスプローラー連携 ----------
    def _selected_path(self):
        """表示中のタブで選ばれている項目のパス(なければ None)"""
        tab = self.notebook.index("current")
        if tab == self.TAB_FOLDER:
            tree, mapping = self.folder_tree, self.row_paths
        elif tab == self.TAB_TREE:
            tree, mapping = self.dir_tree, self.tree_paths
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
        """表示中のタブに応じてCSVに保存する。
        「フォルダの中身」「ツリー表示」… その下のすべてのフォルダ・ファイル(ツリーの順)
        「容量ランキング」… 表示中のランキング"""
        tab = self.notebook.index("current")
        if tab != self.TAB_RANK and not self.result:
            messagebox.showinfo(
                "CSVで保存",
                "スキャンが終わっていないため、まだ保存できません。\n"
                "スキャンが終わってから(中止した場合は「再スキャン」をしてから)保存してください。")
            return
        today = time.strftime("%Y-%m-%d")
        if tab == self.TAB_RANK:
            default = f"容量ランキング_{today}.csv"
            start = None
        else:
            start = self.current if tab == self.TAB_FOLDER else self.root_path
            if start not in self.result.dir_sizes:
                start = self.root_path
            default = f"フォルダ容量_{os.path.basename(start) or 'フォルダ'}_{today}.csv"
        save_path = filedialog.asksaveasfilename(
            title="CSVとして保存" if start is None else f"CSVとして保存({os.path.basename(start) or start} 以下のすべて)",
            defaultextension=".csv", filetypes=[("CSVファイル", "*.csv")], initialfile=default,
        )
        if not save_path:
            return
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            # UTF-8 BOM 付きにすると Excel で開いても文字化けしない
            with open(save_path, "w", newline="", encoding="utf-8-sig") as f:
                if start is None:
                    self._write_rank_csv(f)
                    rows = len(self.rank_tree.get_children())
                else:
                    rows = write_tree_csv(f, self.result, start)
            note = ""
            if rows + 1 > EXCEL_MAX_ROWS:
                note = (f"\n\n行数が {rows:,} 行あり、Excel で開ける上限({EXCEL_MAX_ROWS:,} 行)を超えています。"
                        "Excel では途中までしか表示されません。もっと下のフォルダを開いてから保存してください。")
            messagebox.showinfo("保存完了", f"保存しました({rows:,} 行):\n{save_path}{note}")
        except Exception as e:
            messagebox.showerror(
                "エラー",
                f"保存に失敗しました:\n{e}\n\n"
                "同じ名前のファイルを Excel で開いていないか、保存先に書き込めるかを確かめてください。")
        finally:
            self.root.config(cursor="")

    def _write_rank_csv(self, f):
        writer = csv.writer(f)
        mtime_head = "中の最終更新日時" if self.rank_kind.get() == "dir" else "更新日時"
        writer.writerow(["順位", "名前", "種類", "サイズ(バイト)", "ファイル数", mtime_head, "フルパス"])
        for iid in self.rank_tree.get_children():
            path, is_dir, size, count = self.rank_paths[iid]
            rank, name, _s, _c, mtime, _loc = self.rank_tree.item(iid, "values")
            writer.writerow([rank, name, "フォルダ" if is_dir else "ファイル", size,
                             "" if count is None else count, mtime, path])

    # ---------- 画面ガイド ----------
    def start_guide(self):
        """「❓ 使い方」:画面ガイドを始める。フォルダを開いていなければ、まず見本を開く"""
        if self.guide and self.guide.alive:
            self.guide.card.lift()
            return
        if self.root_path is None:
            self.start_sample()
        self.guide = ScreenGuide(self.root, self._guide_steps())

    def _guide_steps(self):
        """画面ガイドの手順。(対象の部品, 見出し, 説明, 表示前に行う処理)
        ※ 画面の部品を変えたら、ここも必ず直すこと"""
        tab = lambda i: (lambda: self.notebook.select(i))
        return [
            (None, "フォルダ容量ビューアへようこそ",
             "フォルダの中で、どこが容量を使っているかを調べるアプリです。\n"
             "今は「見本」(架空のフォルダ)を開いています。本物のデータには触れないので、"
             "自由に触ってみてください。\n\nこのアプリでファイルを消すことはできません。", None),
            (lambda: self.path_entry, "① 調べるフォルダを指定する",
             "調べたいフォルダのパスを入力して Enter を押します。\n"
             "ネットワークの共有フォルダは「\\\\サーバー名\\共有名\\フォルダ」のように入力します。", None),
            (lambda: self.browse_btn, "「参照...」でも選べます",
             "パスが分からないときは、ここを押してフォルダを選びます。", None),
            (lambda: self.sample_btn, "見本で試す",
             "いつでも、ここを押すと見本のフォルダを開けます。", None),
            (lambda: self.notebook, "② フォルダの中身",
             "エクスプローラーと同じ表示です。フォルダにも合計の容量が出ます。\n"
             "フォルダをダブルクリックすると中に入ります。「サイズ」の見出しを押すと、容量の大きい順に並びます。",
             tab(self.TAB_FOLDER)),
            (lambda: self.up_btn, "上のフォルダへ戻る",
             "ひとつ上のフォルダに戻ります(最初に指定したフォルダより上には行きません)。", None),
            (lambda: self.notebook, "③ ツリー表示",
             "フォルダの階層を、容量の大きい順に表示します。▶ を押すと中が開きます。\n"
             "「割合」で、そのフォルダが親フォルダの何%を使っているかが分かります。",
             tab(self.TAB_TREE)),
            (lambda: self.notebook, "④ 容量ランキング",
             "指定したフォルダの下にあるすべてのファイル(またはフォルダ)を、容量の大きい順に並べます。\n"
             "ダブルクリックすると、その場所を「フォルダの中身」で開きます。",
             tab(self.TAB_RANK)),
            (lambda: self.old_combo, "長く使っていないものを探す",
             "「3年以上前」などを選ぶと、長く更新されていない大きいファイルだけになります。整理の候補探しに便利です。",
             tab(self.TAB_RANK)),
            (lambda: self.csv_btn, "⑤ CSVで保存",
             "記録を残せます(Excel で開けます)。\n"
             "「フォルダの中身」「ツリー表示」では、その下のすべてのフォルダ・ファイルを保存します。"
             "「容量ランキング」では、表示中のランキングを保存します。", None),
            (lambda: self.explorer_btn, "エクスプローラーで表示",
             "選んだファイルやフォルダを、エクスプローラーで開きます。\n"
             "消してよいかどうかは、エクスプローラーで中身を確かめてから判断してください。", None),
            (lambda: self.rescan_btn, "再スキャン",
             "フォルダの中身を整理したあとは、ここを押すと最新の容量に取り直せます。", None),
            (lambda: self.help_btn, "この案内をもう一度見るには",
             "「❓ 使い方」を押すと、いつでもこの案内を最初から見られます。\n"
             "本物のフォルダを調べるときは、パス欄か「参照...」でフォルダを指定してください。",
             tab(self.TAB_FOLDER)),
        ]


def main():
    root = tk.Tk()
    FolderSizeViewer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
