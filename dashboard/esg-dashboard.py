"""
esg-dashboard/dashboard.py
ESG 研究主控台 — 一覽下載進度與圖表萃取狀態。
可與 esg_downloader.py / esg_pdf_cuter.py 同時執行，自動偵測資料更新。
執行：python esg-dashboard/dashboard.py
"""
import threading
import subprocess
import tkinter as tk
import tkinter.ttk as ttk
from pathlib import Path
from datetime import datetime

import pandas as pd

# ============================================================
# 路徑設定（從此檔案往上一層找 data/）
# ============================================================
BASE_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = BASE_DIR / "data"

DOWNLOAD_STATUSES = ['成功', '未找到中文版報告', '已確認無報告', '下載失敗']
AUTO_REFRESH_MS   = 10_000   # 每 10 秒自動偵測一次

# ============================================================
# Apple 風格配色
# ============================================================
APPLE_BG     = '#f5f5f7'
APPLE_CARD   = '#ffffff'
APPLE_BLUE   = '#0071e3'
APPLE_GREEN  = '#34c759'
APPLE_RED    = '#ff3b30'
APPLE_ORANGE = '#ff9f0a'
APPLE_TEXT   = '#1d1d1f'
APPLE_GREY   = '#6e6e73'
APPLE_BORDER = '#d2d2d7'

FONT_TITLE  = ('Helvetica Neue', 13, 'bold')
FONT_MAIN   = ('Helvetica Neue', 10)
FONT_LABEL  = ('Helvetica Neue', 9)
FONT_SMALL  = ('Helvetica Neue', 8)
FONT_NUM    = ('Helvetica Neue', 11, 'bold')
FONT_HEADER = ('Helvetica Neue', 9, 'bold')

# ============================================================
# App Icon
# ============================================================
def set_app_icon(root: tk.Tk, emoji: str = "📊") -> None:
    try:
        from AppKit import NSApplication, NSImage, NSAttributedString, NSFont
        from Foundation import NSMakeSize
        size   = 256
        ns_img = NSImage.alloc().initWithSize_(NSMakeSize(size, size))
        ns_img.lockFocus()
        attrs  = {"NSFont": NSFont.systemFontOfSize_(200)}
        s      = NSAttributedString.alloc().initWithString_attributes_(emoji, attrs)
        s.drawAtPoint_((20, 20))
        ns_img.unlockFocus()
        NSApplication.sharedApplication().setApplicationIconImage_(ns_img)
        import base64
        from io import BytesIO
        from PIL import Image as PILImage
        tiff  = ns_img.TIFFRepresentation()
        pil   = PILImage.open(BytesIO(bytes(tiff)))
        buf   = BytesIO()
        pil.save(buf, format="PNG")
        photo = tk.PhotoImage(data=base64.b64encode(buf.getvalue()).decode())
        root.iconphoto(True, photo)
        root._icon_ref = photo
    except Exception:
        pass

# ============================================================
# 資料讀取
# ============================================================
def _file_fingerprint() -> str:
    """取得所有 Excel 與關鍵目錄的最後修改時間，用來判斷是否需要重整。"""
    parts = []
    if not DATA_DIR.is_dir():
        return ''
    for p in sorted(DATA_DIR.rglob("ESG_Download_Progress_*.xlsx")):
        parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    for p in sorted(DATA_DIR.rglob("ESG_Extract_Results_*.xlsx")):
        parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    # 用各年度 images/ 目錄的 mtime 偵測新增圖片
    for p in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]/*/images")):
        if p.is_dir():
            parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    return '|'.join(parts)


def load_download_stats() -> dict[str, dict]:
    stats = {}
    if not DATA_DIR.is_dir():
        return stats
    for year_dir in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        year = year_dir.name
        xls  = year_dir / f"ESG_Download_Progress_{year}.xlsx"
        if not xls.exists():
            stats[year] = {'_missing': True}
            continue
        try:
            df = pd.read_excel(xls)
            counts = df['status'].value_counts().to_dict()
            stats[year] = {s: counts.get(s, 0) for s in DOWNLOAD_STATUSES}
            stats[year]['_total'] = len(df)
            stats[year]['_df']    = df   # 供細節視窗使用
        except Exception as e:
            stats[year] = {'_error': str(e)}
    return stats


def load_cutter_stats() -> dict[str, dict]:
    stats = {}
    if not DATA_DIR.is_dir():
        return stats
    for year_dir in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        year = year_dir.name
        processed_dirs = [
            d for d in year_dir.iterdir()
            if d.is_dir()
            and (d / "images").is_dir()
            and any((d / "images").glob("*.jpg"))
        ]
        all_pdf_stems = {p.stem for p in year_dir.rglob("*.pdf")}
        processed_stems = {d.name for d in processed_dirs}
        pending = len(all_pdf_stems - processed_stems)

        total_images  = sum(
            len(list((d / "images").glob("*.jpg"))) for d in processed_dirs
        )
        garbled_files = list(year_dir.rglob("garbled_pages.txt"))

        stats[year] = {
            'processed':       len(processed_dirs),
            'pending':         pending,
            'images':          total_images,
            'garbled':         len(garbled_files),
            'garbled_files':   garbled_files,
            'processed_dirs':  processed_dirs,
        }
    return stats

# ============================================================
# 細節視窗（點年度列後彈出）
# ============================================================
class DetailWindow:
    """顯示單一年度的所有公司明細，可搜尋/篩選。"""
    _instances: dict[str, 'DetailWindow'] = {}

    @classmethod
    def open(cls, year: str, dl_row: dict, ct_row: dict):
        if year in cls._instances:
            try:
                cls._instances[year].win.lift()
                return
            except tk.TclError:
                pass
        inst = cls(year, dl_row, ct_row)
        cls._instances[year] = inst

    def __init__(self, year: str, dl_row: dict, ct_row: dict):
        self.year = year
        self.win  = tk.Toplevel()
        self.win.title(f"📋 {year} 年度明細")
        self.win.geometry("860x560")
        self.win.configure(bg=APPLE_BG)
        self.win.protocol("WM_DELETE_WINDOW",
                          lambda: (DetailWindow._instances.pop(year, None),
                                   self.win.destroy()))
        self._build(dl_row, ct_row)

    def _build(self, dl_row: dict, ct_row: dict):
        # ── Header ──
        hdr = tk.Frame(self.win, bg=APPLE_BLUE, pady=8)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"{self.year} 年度明細",
                 font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=16)

        # ── 搜尋列 ──
        sf = tk.Frame(self.win, bg=APPLE_BG, pady=8)
        sf.pack(fill=tk.X, padx=16)
        tk.Label(sf, text="搜尋：", font=FONT_MAIN, bg=APPLE_BG).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        entry = tk.Entry(sf, textvariable=self.search_var,
                         font=FONT_MAIN, width=28,
                         relief='solid', bd=1)
        entry.pack(side=tk.LEFT, padx=(4, 12))

        # 篩選下載狀態
        tk.Label(sf, text="狀態：", font=FONT_MAIN, bg=APPLE_BG).pack(side=tk.LEFT)
        self.filter_var = tk.StringVar(value='全部')
        filter_opts = ['全部'] + DOWNLOAD_STATUSES
        ttk.Combobox(sf, textvariable=self.filter_var,
                     values=filter_opts, width=16,
                     state='readonly').pack(side=tk.LEFT)

        self.search_var.trace_add('write', lambda *_: self._apply_filter())
        self.filter_var.trace_add('write', lambda *_: self._apply_filter())

        # ── 表格 ──
        cols = ('公司代碼', '公司名稱', '下載狀態', '已萃取', '圖片數', '亂碼頁')
        frame = tk.Frame(self.win, bg=APPLE_BG)
        frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))

        vsb = ttk.Scrollbar(frame, orient='vertical')
        hsb = ttk.Scrollbar(frame, orient='horizontal')
        self.tree = ttk.Treeview(
            frame, columns=cols, show='headings',
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            height=20
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        widths = [80, 180, 130, 70, 70, 70]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col,
                              command=lambda c=col: self._sort(c))
            self.tree.column(col, width=w, anchor='center')

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<MouseWheel>',
                       lambda e: (self.tree.yview_scroll(int(-1*(e.delta/120)), 'units'),
                                  'break')[1])

        # 整合下載 df + 萃取資訊
        self._build_rows(dl_row, ct_row)
        self._apply_filter()

        # 統計列
        tk.Label(self.win,
                 text=f"共 {len(self.all_rows)} 筆",
                 font=FONT_LABEL, fg=APPLE_GREY, bg=APPLE_BG
                 ).pack(pady=(0, 8))

    def _build_rows(self, dl_row: dict, ct_row: dict):
        """整合下載 Excel + 實際目錄掃描，建立完整列表。"""
        # 已萃取 stem set
        processed_stems = {d.name for d in ct_row.get('processed_dirs', [])}
        # stem → 圖片數
        img_counts = {
            d.name: len(list((d / "images").glob("*.jpg")))
            for d in ct_row.get('processed_dirs', [])
        }
        # stem → 亂碼頁
        garbled_map: dict[str, str] = {}
        for gf in ct_row.get('garbled_files', []):
            try:
                pages = gf.read_text(encoding='utf-8').strip()
                garbled_map[gf.parent.name] = pages
            except Exception:
                garbled_map[gf.parent.name] = '?'

        self.all_rows = []
        df = dl_row.get('_df')
        if df is not None:
            for _, r in df.iterrows():
                # 嘗試從檔名欄位取 stem
                raw = str(r.get('file_name', r.get('filename', r.get('檔名', ''))))
                stem = Path(raw).stem if raw else ''
                code   = str(r.get('stock_id', r.get('stock_code', r.get('代碼', ''))))
                name   = str(r.get('company_name', r.get('公司名稱', '')))
                status = str(r.get('status', r.get('狀態', '')))

                extracted = '✅' if stem in processed_stems else (
                    '—' if status != '成功' else '⏳'
                )
                imgs    = img_counts.get(stem, 0)
                garbled = garbled_map.get(stem, '')

                self.all_rows.append((
                    code, name, status,
                    extracted,
                    str(imgs) if imgs else '—',
                    garbled or '—',
                ))

    def _apply_filter(self):
        kw     = self.search_var.get().strip().lower()
        status = self.filter_var.get()
        self.tree.delete(*self.tree.get_children())
        for row in self.all_rows:
            if status != '全部' and row[2] != status:
                continue
            if kw and not any(kw in str(c).lower() for c in row):
                continue
            tag = 'ok' if row[2] == '成功' else (
                  'fail' if row[2] == '下載失敗' else 'other')
            self.tree.insert('', 'end', values=row, tags=(tag,))

        self.tree.tag_configure('ok',    foreground=APPLE_GREEN)
        self.tree.tag_configure('fail',  foreground=APPLE_RED)
        self.tree.tag_configure('other', foreground=APPLE_GREY)

    _sort_reverse: dict[str, bool] = {}
    def _sort(self, col: str):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        rev = self._sort_reverse.get(col, False)
        items.sort(reverse=rev)
        for i, (_, k) in enumerate(items):
            self.tree.move(k, '', i)
        self._sort_reverse[col] = not rev
        self.tree.heading(col, text=col + (' ↑' if not rev else ' ↓'))

# ============================================================
# 主控台視窗
# ============================================================
class Dashboard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("📊 ESG 研究主控台")
        self.root.geometry("960x720")
        self.root.configure(bg=APPLE_BG)
        self.root.resizable(True, True)
        set_app_icon(self.root)

        self._last_fingerprint = ''
        self._dl_stats: dict = {}
        self._ct_stats: dict = {}

        self._build_ui()
        self.refresh(force=True)
        self._schedule_auto_refresh()

    # ── Header ──────────────────────────────────────────────
    def _build_ui(self):
        header = tk.Frame(self.root, bg=APPLE_BLUE, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="📊  ESG 研究主控台",
                 font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=20)

        self.status_dot = tk.Label(header, text='●', font=FONT_MAIN,
                                   fg='#a8d4ff', bg=APPLE_BLUE)
        self.status_dot.pack(side=tk.RIGHT, padx=(0, 8))
        self.last_updated = tk.Label(header, text='', font=FONT_LABEL,
                                     fg='#a8d4ff', bg=APPLE_BLUE)
        self.last_updated.pack(side=tk.RIGHT, padx=(0, 4))

        tk.Button(header, text="↺  重新整理",
                  font=FONT_LABEL, bg='#005bb5', fg='white',
                  activebackground='#004a99', relief='flat',
                  padx=10, pady=4, cursor='hand2',
                  command=lambda: self.refresh(force=True)
                  ).pack(side=tk.RIGHT, padx=4)
        tk.Button(header, text="📁  開啟 data/",
                  font=FONT_LABEL, bg='#005bb5', fg='white',
                  activebackground='#004a99', relief='flat',
                  padx=10, pady=4, cursor='hand2',
                  command=lambda: subprocess.Popen(['open', str(DATA_DIR)])
                  ).pack(side=tk.RIGHT, padx=4)

        # ── Scrollable body ──
        body_outer = tk.Frame(self.root, bg=APPLE_BG)
        body_outer.pack(fill=tk.BOTH, expand=True)

        canvas    = tk.Canvas(body_outer, bg=APPLE_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.body = tk.Frame(canvas, bg=APPLE_BG)
        self.body_window = canvas.create_window((0, 0), window=self.body, anchor='nw')

        self.body.bind('<Configure>',
                       lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(self.body_window, width=e.width))
        self._canvas = canvas
        canvas.bind('<Enter>',
                    lambda _: canvas.bind_all('<MouseWheel>', self._on_canvas_scroll))
        canvas.bind('<Leave>',
                    lambda _: canvas.unbind_all('<MouseWheel>'))

    def _on_canvas_scroll(self, e):
        self._canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')

    # ── 自動重整排程 ────────────────────────────────────────
    def _schedule_auto_refresh(self):
        self.refresh(force=False)
        self.root.after(AUTO_REFRESH_MS, self._schedule_auto_refresh)

    # ── 重新整理邏輯 ────────────────────────────────────────
    def refresh(self, force: bool = True):
        self.status_dot.config(fg='#ffcc00')   # 黃色 = 讀取中

        def _load():
            fp = _file_fingerprint()
            if not force and fp == self._last_fingerprint:
                self.root.after(0, lambda: self.status_dot.config(fg='#a8d4ff'))
                return
            dl = load_download_stats()
            ct = load_cutter_stats()
            self.root.after(0, lambda: self._render(dl, ct, fp))

        threading.Thread(target=_load, daemon=True).start()

    def _render(self, dl_stats: dict, ct_stats: dict, fingerprint: str):
        self._last_fingerprint = fingerprint
        self._dl_stats = dl_stats
        self._ct_stats = ct_stats

        for w in self.body.winfo_children():
            w.destroy()

        self._build_summary(dl_stats, ct_stats)
        self._build_download_section(dl_stats)
        self._build_cutter_section(ct_stats)
        tk.Frame(self.body, bg=APPLE_BG, height=20).pack()

        self.last_updated.config(
            text=f"更新 {datetime.now().strftime('%H:%M:%S')}")
        self.status_dot.config(fg='#5cff7a')   # 綠色 = 已同步

    # ── 摘要數字卡 ──────────────────────────────────────────
    def _build_summary(self, dl_stats: dict, ct_stats: dict):
        frame = tk.Frame(self.body, bg=APPLE_BG, padx=20, pady=10)
        frame.pack(fill=tk.X)

        total_success   = sum(s.get('成功', 0) for s in dl_stats.values()
                              if not s.get('_missing') and not s.get('_error'))
        total_images    = sum(s['images']    for s in ct_stats.values())
        total_processed = sum(s['processed'] for s in ct_stats.values())
        total_garbled   = sum(s['garbled']   for s in ct_stats.values())

        for label, val, color in [
            ("已下載 PDF",  f"{total_success:,}",   APPLE_GREEN),
            ("已萃取公司",  f"{total_processed:,}", APPLE_BLUE),
            ("圖片總數",    f"{total_images:,}",    APPLE_TEXT),
            ("亂碼公司",    f"{total_garbled:,}",
             APPLE_ORANGE if total_garbled else APPLE_GREY),
        ]:
            card = tk.Frame(frame, bg=APPLE_CARD,
                            highlightthickness=1, highlightbackground=APPLE_BORDER)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)
            tk.Label(card, text=label, font=FONT_LABEL,
                     fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(8, 0))
            tk.Label(card, text=val,
                     font=('Helvetica Neue', 18, 'bold'),
                     fg=color, bg=APPLE_CARD).pack(pady=(2, 8))

    # ── 下載狀態表 ──────────────────────────────────────────
    def _section_title(self, text, subtext=''):
        f = tk.Frame(self.body, bg=APPLE_BG)
        f.pack(fill=tk.X, padx=20, pady=(20, 6))
        tk.Label(f, text=text, font=('Helvetica Neue', 12, 'bold'),
                 fg=APPLE_TEXT, bg=APPLE_BG).pack(side=tk.LEFT)
        if subtext:
            tk.Label(f, text=subtext, font=FONT_LABEL,
                     fg=APPLE_GREY, bg=APPLE_BG).pack(side=tk.LEFT, padx=(8, 0))
        tk.Frame(self.body, bg=APPLE_BORDER, height=1).pack(fill=tk.X, padx=20, pady=(0, 8))

    def _table_header(self, cols: list[tuple]):
        row = tk.Frame(self.body, bg='#e8e8ed')
        row.pack(fill=tk.X, padx=20)
        for text, w, anchor in cols:
            tk.Label(row, text=text, font=FONT_HEADER,
                     fg=APPLE_GREY, bg='#e8e8ed',
                     width=w, anchor=anchor).pack(side=tk.LEFT, padx=4, pady=4)

    def _table_row(self, cells: list[tuple], on_click=None):
        bg  = APPLE_CARD
        row = tk.Frame(self.body, bg=bg,
                       highlightthickness=1, highlightbackground=APPLE_BORDER,
                       cursor='hand2' if on_click else '')
        row.pack(fill=tk.X, padx=20, pady=1)
        if on_click:
            row.bind('<Button-1>', lambda e: on_click())
        for text, w, anchor, color in cells:
            lbl = tk.Label(row, text=text, font=FONT_MAIN,
                           fg=color or APPLE_TEXT, bg=bg,
                           width=w, anchor=anchor)
            lbl.pack(side=tk.LEFT, padx=4, pady=6)
            if on_click:
                lbl.bind('<Button-1>', lambda e: on_click())

    def _build_download_section(self, stats: dict):
        self._section_title("下載狀態",
                            "來源：ESG_Download_Progress_YYYY.xlsx  ·  點選列查看明細")
        self._table_header([
            ('年度',        8,  'center'),
            ('✅ 成功',      8,  'center'),
            ('⚠️ 未找到',    10, 'center'),
            ('🔒 已確認無',  10, 'center'),
            ('❌ 失敗',      8,  'center'),
            ('共',           6,  'center'),
            ('進度',        18, 'center'),
        ])
        for year, s in stats.items():
            if s.get('_missing'):
                self._table_row([
                    (year, 8, 'center', APPLE_TEXT),
                    ('尚無資料', 56, 'center', APPLE_GREY),
                ])
                continue
            if s.get('_error'):
                self._table_row([
                    (year, 8, 'center', APPLE_TEXT),
                    (f"讀取錯誤：{s['_error']}", 56, 'w', APPLE_RED),
                ])
                continue

            total   = s.get('_total', 1) or 1
            success = s.get('成功', 0)
            pct     = int(success / total * 100)
            bar     = '█' * (pct // 20) + '░' * (5 - pct // 20)

            ct = self._ct_stats.get(year, {})
            self._table_row([
                (year,                              8,  'center', APPLE_TEXT),
                (str(success),                      8,  'center', APPLE_GREEN),
                (str(s.get('未找到中文版報告', 0)),  10, 'center', APPLE_ORANGE),
                (str(s.get('已確認無報告', 0)),      10, 'center', APPLE_GREY),
                (str(s.get('下載失敗', 0)),          8,  'center',
                    APPLE_RED if s.get('下載失敗', 0) else APPLE_GREY),
                (str(total),                        6,  'center', APPLE_TEXT),
                (f"{bar} {pct}%",                   18, 'center', APPLE_BLUE),
            ], on_click=lambda y=year, ds=s, cs=ct: DetailWindow.open(y, ds, cs))

    def _build_cutter_section(self, stats: dict):
        self._section_title("圖表萃取狀態",
                            "來源：掃描 data/{year}/*/images/*.jpg  ·  點選列查看明細")
        self._table_header([
            ('年度',        8,  'center'),
            ('✅ 已萃取',    9,  'center'),
            ('⏳ 待處理',    9,  'center'),
            ('🖼 圖片數',   11, 'center'),
            ('⚠️ 亂碼公司', 10, 'center'),
            ('進度',        18, 'center'),
        ])
        for year, s in stats.items():
            processed = s['processed']
            pending   = s['pending']
            total     = processed + pending or 1
            pct       = int(processed / total * 100)
            bar       = '█' * (pct // 20) + '░' * (5 - pct // 20)

            if processed == 0 and pending == 0:
                status_text  = '尚無資料'
                status_color = APPLE_GREY
            elif pending == 0:
                status_text  = f"{bar} 100%"
                status_color = APPLE_GREEN
            else:
                status_text  = f"{bar} {pct}%"
                status_color = APPLE_BLUE

            dl = self._dl_stats.get(year, {})
            self._table_row([
                (year,                8,  'center', APPLE_TEXT),
                (str(processed),      9,  'center',
                    APPLE_GREEN if processed else APPLE_GREY),
                (str(pending),        9,  'center',
                    APPLE_ORANGE if pending else APPLE_GREY),
                (f"{s['images']:,}",  11, 'center',
                    APPLE_TEXT if s['images'] else APPLE_GREY),
                (str(s['garbled']),   10, 'center',
                    APPLE_ORANGE if s['garbled'] else APPLE_GREY),
                (status_text,         14, 'center', status_color),
            ], on_click=lambda y=year, ds=dl, cs=s: DetailWindow.open(y, ds, cs))

    def run(self):
        self.root.mainloop()


# ============================================================
# 主程式
# ============================================================
if __name__ == '__main__':
    Dashboard().run()
