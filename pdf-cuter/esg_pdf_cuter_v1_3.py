"""
esg_pdf_cuter_v1_3.py
v1_3 核心邏輯的 GUI 版本，供與 esg_pdf_cuter.py 比較萃取品質。

差異：
- Vector：全頁路徑合併成「一個大框」+ 80pt 擴張（不做 Union-Find 聚類）
- 不做文字遮罩（MASK_UNRELATED）
- 不做 QR code / 全頁圖過濾
- 輸出目錄為 images_v1_3/（與 images/ 並存）
"""
import os
import subprocess
import threading
import queue
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, scrolledtext
from pathlib import Path
from datetime import datetime

import fitz           # PyMuPDF
import pandas as pd

# ============================================================
# 路徑設定
# ============================================================
BASE_DIR = Path(__file__).parent.absolute()
DATA_DIR = BASE_DIR.parent / "data"
DATA_DIR.mkdir(exist_ok=True)

IMG_SUBDIR = "images_v1_3"      # 輸出子資料夾名稱

def year_dir(year: str) -> Path:
    return DATA_DIR / str(year)

def year_excel(year: str) -> Path:
    return year_dir(year) / f"ESG_Extract_Results_v1_3_{year}.xlsx"

def _year_range():
    return range(2015, 2025)

def available_years():
    if not DATA_DIR.is_dir():
        return [str(y) for y in _year_range()]
    dirs = [d for d in os.listdir(DATA_DIR)
            if (DATA_DIR / d).is_dir() and d.isdigit()]
    return sorted(dirs) or [str(y) for y in _year_range()]

# ============================================================
# App Icon（Dock / 視窗）
# ============================================================
def set_app_icon(root: tk.Tk, emoji: str = "🌱") -> None:
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
        raw   = bytes(tiff)
        pil   = PILImage.open(BytesIO(raw))
        buf   = BytesIO()
        pil.save(buf, format="PNG")
        photo = tk.PhotoImage(data=base64.b64encode(buf.getvalue()).decode())
        root.iconphoto(True, photo)
        root._icon_ref = photo

    except Exception:
        pass


# ============================================================
# Apple 風格配色
# ============================================================
APPLE_BG     = '#f5f5f7'
APPLE_CARD   = '#ffffff'
APPLE_BLUE   = '#0071e3'
APPLE_TEXT   = '#1d1d1f'
APPLE_GREY   = '#6e6e73'
APPLE_BORDER = '#d2d2d7'

FONT_TITLE = ('Helvetica Neue', 13, 'bold')
FONT_MAIN  = ('Helvetica Neue', 10)
FONT_LABEL = ('Helvetica Neue', 9)
FONT_STAT  = ('Helvetica Neue', 20, 'bold')
FONT_LOG   = ('Menlo', 9)

# ============================================================
# 執行緒狀態
# ============================================================
log_queue    = queue.Queue()
program_done = threading.Event()
pause_event  = threading.Event()
paused_event = threading.Event()

ui_stats = {
    'total': 0, 'done': 0, 'images': 0, 'skipped': 0, 'error': 0,
}

# ============================================================
# 萃取參數（v1_3 邏輯）
# ============================================================
RENDER_SCALE    = 3      # 渲染倍率（3x = 216 DPI）
EXPAND_PT       = 80     # 向量框擴張距離（比 esg_pdf_cuter.py 的 20pt 大）
MIN_AREA_PCT    = 0.5    # 最小面積佔比（%）
MAX_AREA_PCT    = 90     # 最大面積佔比（%）
MIN_PATHS       = 10     # 頁面路徑數 >= 此值才偵測 Vector
MIN_DIM_PT      = 100    # 最小寬/高（pt）
MAX_PAGE_RATIO  = 0.95   # 超過此比例視為整頁背景

# ============================================================
# 核心函式（v1_3：單一 union + 80pt 擴張）
# ============================================================
def process_pdf(pdf_path: str, year: str) -> list[dict]:
    doc       = fitz.open(pdf_path)
    file_stem = Path(pdf_path).stem
    base_dir  = DATA_DIR / str(year) / file_stem
    img_dir   = base_dir / IMG_SUBDIR
    img_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict] = []

    for page_index, page in enumerate(doc):
        page_num  = page_index + 1
        try:
            page_rect = page.rect
            page_area = page_rect.width * page_rect.height
            candidates = []

            # 方法一：Raster 點陣圖
            for img_info in page.get_images(full=True):
                for r in page.get_image_rects(img_info[0]):
                    if r.width > 50 and r.height > 50:
                        candidates.append((r, 'RA'))

            # 方法二：Vector — 全部合併成一個大框（v1_3 核心）
            paths = page.get_drawings()
            if len(paths) >= MIN_PATHS:
                drawing_rects = [
                    p["rect"] for p in paths
                    if p["rect"].width > 5 and p["rect"].height > 5
                ]
                if drawing_rects:
                    combined = drawing_rects[0]
                    for r in drawing_rects[1:]:
                        combined |= r
                    expanded = combined + (-EXPAND_PT, -EXPAND_PT, EXPAND_PT, EXPAND_PT)
                    expanded &= page_rect
                    if (expanded.width > MIN_DIM_PT and expanded.height > MIN_DIM_PT
                            and expanded.width < page_rect.width * MAX_PAGE_RATIO):
                        candidates.append((expanded, 'VC'))

            asset_idx = 0
            for r, tcode in candidates:
                area_pct = round(r.width * r.height / page_area * 100, 4)
                if area_pct < MIN_AREA_PCT or area_pct > MAX_AREA_PCT:
                    continue

                asset_idx += 1
                img_name  = f"{file_stem}_p{page_num}_{asset_idx}_{tcode}.png"
                save_path = img_dir / img_name

                pix = None
                for scale in (RENDER_SCALE, 2, 1):
                    try:
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(scale, scale),
                            clip=r, alpha=False)
                        break
                    except Exception:
                        pix = None
                if pix is None:
                    log_queue.put(('warning',
                        f'  無法渲染 {file_stem} p{page_num} 區塊 {asset_idx}，跳過'))
                    asset_idx -= 1
                    continue

                pix.save(str(save_path))
                pix = None

                results.append({
                    "年份":           year,
                    "PDF檔名":        file_stem,
                    "PDF總頁數":      len(doc),
                    "頁碼":           page_num,
                    "圖片編號":       asset_idx,
                    "圖片面積佔比(%)": area_pct,
                    "類型":           "Raster" if tcode == "RA" else "Vector",
                    "圖片檔名":       img_name,
                    "存檔路徑":       str(save_path),
                })

        except Exception as e:
            log_queue.put(('warning', f'  跳過 {file_stem} 第 {page_num} 頁：{e}'))

    doc.close()
    return results

# ============================================================
# 萃取執行緒
# ============================================================
def _is_already_processed(pdf_path: str, year: str) -> bool:
    file_stem = Path(pdf_path).stem
    img_dir   = DATA_DIR / str(year) / file_stem / IMG_SUBDIR
    if not img_dir.is_dir():
        return False
    return any(img_dir.glob("*.png"))


def run_extraction(years):
    tasks = []
    for year in years:
        pdf_folder = DATA_DIR / year
        if not pdf_folder.is_dir():
            log_queue.put(('warning', f'找不到資料夾：{pdf_folder}'))
            continue
        for pdf_file in sorted(pdf_folder.rglob("*.pdf")):
            tasks.append((str(pdf_file), year))

    total   = len(tasks)
    pending = [(p, y) for p, y in tasks if not _is_already_processed(p, y)]
    skipped = total - len(pending)

    ui_stats.update({'total': total, 'done': skipped, 'images': 0,
                     'skipped': skipped, 'error': 0})
    log_queue.put(('info',
        f'共 {total} 個 PDF，已有輸出跳過 {skipped} 個，待處理 {len(pending)} 個'))

    if not pending:
        log_queue.put(('info', '所有檔案皆已處理完成'))
        program_done.set()
        return

    year_data: dict[str, list] = {}
    for y in set(yr for _, yr in pending):
        xls = year_excel(y)
        if xls.exists():
            try:
                year_data[y] = pd.read_excel(xls).to_dict('records')
            except Exception:
                year_data[y] = []
        else:
            year_data[y] = []

    for i, (pdf_path, year) in enumerate(pending):
        if pause_event.is_set():
            log_queue.put(('warning', '⏸ 已暫停，進度已儲存，可安全關閉視窗'))
            paused_event.set()
            while pause_event.is_set():
                if program_done.is_set():
                    return
                time.sleep(0.2)
            paused_event.clear()
            log_queue.put(('info', '▶ 繼續執行'))

        fname = os.path.basename(pdf_path)
        log_queue.put(('info', f'[{i+1}/{len(pending)}] 處理 {fname}'))

        try:
            results = process_pdf(pdf_path, year)
            year_data[year].extend(results)
            ui_stats['images'] += len(results)
            ui_stats['done']   += 1

            xls = year_excel(year)
            xls.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(year_data[year]).to_excel(xls, index=False)
            log_queue.put(('success', f'  完成：切割 {len(results)} 個區塊'))

        except Exception as e:
            ui_stats['error'] += 1
            log_queue.put(('error', f'  錯誤：{fname} — {e}'))

    log_queue.put(('success',
                   f'全部完成！共切割 {ui_stats["images"]} 個區塊，錯誤 {ui_stats["error"]} 個'))
    program_done.set()

# ============================================================
# 啟動設定視窗
# ============================================================
def create_startup_window():
    selected_years = []

    root = tk.Tk()
    root.title("🌱 ESG 圖表萃取系統 v1_3")
    root.geometry("480x380")
    root.configure(bg=APPLE_BG)
    root.resizable(False, False)
    set_app_icon(root)

    header = tk.Frame(root, bg=APPLE_BLUE, pady=14)
    header.pack(fill=tk.X)
    tk.Label(header, text="ESG 圖表萃取系統 v1_3", font=FONT_TITLE,
             fg='white', bg=APPLE_BLUE).pack()
    tk.Label(header, text="單一 Union + 80pt 擴張（無聚類、無遮罩）", font=FONT_LABEL,
             fg='#a8d4ff', bg=APPLE_BLUE).pack(pady=(2, 0))

    if not DATA_DIR.is_dir():
        messagebox.showerror(
            "找不到資料來源",
            f"找不到以下資料夾：\n{DATA_DIR}\n\n"
            "請確認 data/ 資料夾存在於專案根目錄。"
        )
        root.destroy()
        return selected_years

    content = tk.Frame(root, bg=APPLE_BG, padx=25, pady=15)
    content.pack(fill=tk.BOTH, expand=True)

    tk.Label(content, text="請選擇要處理的年份（可多選）",
             font=('Helvetica Neue', 11, 'bold'),
             fg=APPLE_TEXT, bg=APPLE_BG).pack(anchor='w', pady=(0, 10))

    grid = tk.Frame(content, bg=APPLE_BG)
    grid.pack(fill=tk.X)

    all_years = available_years()
    year_vars = {}
    for i, y in enumerate(all_years):
        var = tk.BooleanVar(value=False)
        cb  = tk.Checkbutton(grid, text=str(y), variable=var,
                             font=FONT_MAIN, bg=APPLE_BG, fg=APPLE_TEXT,
                             activebackground=APPLE_BG, selectcolor=APPLE_CARD,
                             cursor='hand2')
        cb.grid(row=i // 5, column=i % 5, sticky='w', padx=10, pady=4)
        year_vars[y] = var

    def on_start():
        years = sorted(y for y, v in year_vars.items() if v.get())
        if not years:
            messagebox.showwarning("未選擇年份", "請至少選擇一個年份")
            return
        selected_years.extend(years)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", lambda: root.destroy())

    btn_frame = tk.Frame(root, bg=APPLE_BG, pady=15)
    btn_frame.pack()
    tk.Button(btn_frame, text="▶  開始萃取",
              font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              activebackground='#0051a8', activeforeground='white',
              relief='flat', padx=22, pady=9, cursor='hand2',
              command=on_start).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_frame, text="📁  開啟輸出資料夾",
              font=FONT_MAIN, bg=APPLE_CARD, fg=APPLE_TEXT,
              activebackground=APPLE_BORDER, relief='flat', padx=22, pady=9,
              cursor='hand2',
              command=lambda: subprocess.Popen(['open', str(DATA_DIR)])).pack(side=tk.LEFT, padx=8)

    root.mainloop()
    return selected_years

# ============================================================
# 進度視窗
# ============================================================
def create_progress_window(years):
    year_label = '、'.join(str(y) for y in years)
    root = tk.Tk()
    root.title(f"🌱 ESG 圖表萃取系統 v1_3 | {year_label} 年")
    root.geometry("1000x700")
    root.configure(bg=APPLE_BG)
    root.resizable(True, True)
    set_app_icon(root)

    def on_close():
        if program_done.is_set():
            root.destroy()
        elif paused_event.is_set():
            program_done.set()
            pause_event.clear()
            root.destroy()
        else:
            messagebox.showinfo(
                "程式仍在執行",
                "程式正在處理中，請稍候。\n\n"
                "點「⏸ 暫停」等目前這份 PDF 處理完後再關閉。"
            )

    root.protocol("WM_DELETE_WINDOW", on_close)

    # --- Header ---
    header = tk.Frame(root, bg=APPLE_BLUE, pady=12)
    header.pack(fill=tk.X)
    tk.Label(header, text="ESG 圖表萃取系統 v1_3",
             font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=20)
    status_dot = tk.Label(header, text='● 初始化', font=FONT_MAIN,
                          fg='#ffdd57', bg=APPLE_BLUE)
    status_dot.pack(side=tk.RIGHT, padx=20)

    # --- 統計卡片 ---
    cards_frame = tk.Frame(root, bg=APPLE_BG, pady=10)
    cards_frame.pack(fill=tk.X, padx=15)

    def make_stat_card(parent, label):
        card = tk.Frame(parent, bg=APPLE_CARD, bd=0,
                        highlightthickness=1, highlightbackground=APPLE_BORDER)
        card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)
        tk.Label(card, text=label, font=FONT_LABEL,
                 fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(8, 0))
        val_var = tk.StringVar(value='—')
        tk.Label(card, textvariable=val_var, font=FONT_STAT,
                 fg=APPLE_TEXT, bg=APPLE_CARD).pack(pady=(0, 8))
        return val_var

    stat_processed = make_stat_card(cards_frame, '已處理')
    stat_images    = make_stat_card(cards_frame, '圖表張數')
    stat_skipped   = make_stat_card(cards_frame, '已跳過')
    stat_error     = make_stat_card(cards_frame, '錯誤')

    # --- 進度條 ---
    prog_frame = tk.Frame(root, bg=APPLE_BG)
    prog_frame.pack(fill=tk.X, padx=20, pady=(0, 8))
    progress_bar = ttk.Progressbar(prog_frame, mode='determinate', length=960)
    progress_bar.pack(fill=tk.X)

    status_frame = tk.Frame(root, bg=APPLE_BG)
    status_frame.pack(fill=tk.X, padx=20)
    last_status_var = tk.StringVar(value='等待開始...')
    tk.Label(status_frame, textvariable=last_status_var,
             font=FONT_LABEL, fg=APPLE_GREY, bg=APPLE_BG, anchor='w').pack(fill=tk.X)

    tk.Frame(root, bg=APPLE_BORDER, height=1).pack(fill=tk.X, padx=15, pady=6)

    # --- Log 區 ---
    log_frame = tk.Frame(root, bg=APPLE_CARD,
                         highlightthickness=1, highlightbackground=APPLE_BORDER)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 8))
    log_text = scrolledtext.ScrolledText(
        log_frame, state='disabled', wrap=tk.WORD,
        font=FONT_LOG, bg=APPLE_CARD, fg=APPLE_TEXT,
        relief='flat', borderwidth=0, padx=8, pady=6)
    log_text.pack(fill=tk.BOTH, expand=True)
    log_text.tag_configure('success', foreground='#1a7f37')
    log_text.tag_configure('error',   foreground='#cf222e')
    log_text.tag_configure('warning', foreground='#9a6700')
    log_text.tag_configure('info',    foreground=APPLE_BLUE)
    log_text.tag_configure('skip',    foreground=APPLE_GREY)

    # --- 底部列 ---
    bottom = tk.Frame(root, bg=APPLE_BG, pady=8)
    bottom.pack(fill=tk.X, padx=15)

    pause_btn_text = tk.StringVar(value='⏸  暫停（目前 PDF 完成後生效）')

    def toggle_pause():
        if pause_event.is_set():
            pause_event.clear()
            pause_btn_text.set('⏸  暫停（目前 PDF 完成後生效）')
        else:
            if messagebox.askyesno("確認暫停", "確定要暫停嗎？\n目前這份 PDF 處理完後會暫停，進度自動儲存。"):
                pause_event.set()
                pause_btn_text.set('▶  繼續執行')

    tk.Button(bottom, textvariable=pause_btn_text,
              font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              activebackground='#0051a8', activeforeground='white',
              relief='flat', padx=16, pady=7, cursor='hand2', bd=0,
              command=toggle_pause).pack(side=tk.LEFT)
    tk.Button(bottom, text="📁  開啟輸出資料夾",
              font=FONT_MAIN, bg=APPLE_CARD, fg=APPLE_TEXT,
              activebackground=APPLE_BORDER, relief='flat', padx=16, pady=7,
              cursor='hand2', bd=0,
              command=lambda: subprocess.Popen(['open', str(DATA_DIR)])).pack(side=tk.LEFT, padx=8)
    time_label = tk.Label(bottom, text='', font=FONT_LABEL,
                          fg=APPLE_GREY, bg=APPLE_BG)
    time_label.pack(side=tk.RIGHT)

    # --- UI 更新 ---
    def update_ui():
        while not log_queue.empty():
            tag, msg = log_queue.get()
            log_text.configure(state='normal')
            ts = datetime.now().strftime('%H:%M:%S')
            log_text.insert(tk.END, f'[{ts}] ', 'skip')
            log_text.insert(tk.END, msg + '\n', tag)
            log_text.see(tk.END)
            log_text.configure(state='disabled')
            if tag in ('success', 'error', 'info', 'warning'):
                last_status_var.set(msg.strip()[:120])

        tot  = ui_stats['total']
        done = ui_stats['done']
        stat_processed.set(f'{done}/{tot}' if tot else '—')
        stat_images.set(str(ui_stats['images']))
        stat_skipped.set(str(ui_stats['skipped']))
        stat_error.set(str(ui_stats['error']) if ui_stats['error'] else '—')

        if tot > 0:
            progress_bar['value'] = done / tot * 100

        if program_done.is_set():
            status_dot.config(text='■ 已完成', fg='#8e8e93')
        elif paused_event.is_set():
            status_dot.config(text='⏸ 已暫停', fg='#ff9f0a')
        elif done > ui_stats['skipped']:
            status_dot.config(text='● 執行中', fg='#34c759')

        time_label.config(text=f'更新時間 {datetime.now().strftime("%H:%M:%S")}')
        root.after(500, update_ui)

    # --- 啟動執行緒 ---
    threading.Thread(target=run_extraction, args=(years,), daemon=True).start()
    update_ui()
    root.mainloop()

# ============================================================
# 主程式
# ============================================================
if __name__ == '__main__':
    years = create_startup_window()
    if years:
        create_progress_window(years)
