# ESG 報告書自動化分析系統

> 台灣上市公司永續報告書（ESG）大規模批量下載、圖表萃取與 AI 分類平台
> 適用年份：2015 ～ 2024（共 10 個年度）

---

## 1. 系統架構與專案總覽

本系統以「研究用資料蒐集流水線」為核心設計概念，針對台灣 1,078 家上市公司的 ESG 永續報告書，提供從原始 PDF 下載到圖表 AI 分類的全自動化處理流程。

### 整體運作邏輯

```
 台灣上市公司官方網站 https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW
          │
          │  Selenium 爬蟲
          ▼
 ┌───────────────────────┐
 │   report-downloader   │  → data/{year}/{sid}_{name}.pdf
 │ (Step 1：ESG報告書下載) │  → ESG_Download_Progress_{year}.xlsx
 └───────────────────────┘
          │
          ▼
 ┌───────────────────────┐
 │       pdf-cuter       │  → data/{year}/{sid}_{name}/images/*.jpg
 │  (Step 2：萃取 chart)  │  → data/{year}/{sid}_{name}/texts/*.txt
 └───────────────────────┘  → ESG_Extract_Results_{year}.xlsx
          │
          ▼
 ┌───────────────────────┐
 │     chart-counter     │  → data/{year}/{sid}_{name}/charts/*.jpg
 │   (Step 3：cnn 分類)   │  → data/chart_statistics.xlsx
 └───────────────────────┘
          │
          ▼
 ┌───────────────────────┐
 │       dashboard       │  監控上三個步驟的全域進度 (data)
 └───────────────────────┘
```

每個工具皆為獨立可執行的 Python GUI 程式，使用統一的 Apple 設計語言（Helvetica Neue 字體、藍色頂欄、白色卡片配色）。所有工具透過 `data/` 目錄共享狀態，不需要任何中介服務或資料庫。

---

## 2. 完整技術棧解析

### 語言與執行環境

| 項目 | 版本 | 說明 |
|------|------|------|
| Python | 3.12 | 主要語言（macOS 上 3.9 有 Tk 相容性問題，需 3.10+） |
| tkinter / ttk | 內建 | 跨平台 GUI 框架，所有四個工具的使用者介面 |

### 核心相依套件

| 套件 | 版本 | 在本專案的具體功能 |
|------|------|-----------------|
| **PyMuPDF (fitz)** | 1.26.5 | PDF 核心引擎。呼叫 `get_images()`、`get_drawings()`、`get_text()`、`get_pixmap()` 進行圖表偵測與高解析度渲染 |
| **pandas** | 2.3.3 | 讀寫進度 Excel（下載進度、萃取結果）；狀態統計與篩選 |
| **openpyxl** | 3.1.5 | chart-counter 的多 Sheet Excel 輸出；pandas Excel 後端 |
| **Pillow** | 11.3.0 | CLIP 分類前的圖片載入與 RGB 轉換 |
| **selenium** | 4.36.0 | 無頭 Chrome 瀏覽器控制，模擬人工操作下載 PDF |
| **transformers** | ≥4.40.0 | HuggingFace CLIP 模型管理（`CLIPModel`、`CLIPProcessor`） |
| **torch** | ≥2.2.0 | PyTorch 推理引擎；自動偵測 MPS（Apple Silicon）/ CUDA / CPU |
| **pyobjc-core / Cocoa** | 11.1 | macOS 專屬：設定 Dock 圖示（`NSApplication.setApplicationIconImage_`） |

### 基礎設施

| 項目 | 說明 |
|------|------|
| Git + GitHub | 版本控制；PDF 與圖片透過 `.gitignore` 排除，只追蹤程式碼與 Excel 進度檔 |
| Chrome + ChromeDriver | Selenium 需要，由 selenium manager 自動管理版本 |
| HuggingFace Hub | CLIP 模型首次執行時自動下載（~600 MB，快取於 `~/.cache/huggingface/`） |

---

## 3. 目錄結構與模組拆解

```
esg-report/
├── ESG.png                          # 統一 App 圖示（所有工具的 Dock icon）
├── requirements.txt                 # 全域相依套件清單
├── .gitignore                       # 排除 PDF、圖片、.venv
│
├── report-downloader/               # 工具一：爬蟲下載器
│   ├── esg_downloader.py            # 主程式（~900 行）
│   ├── tw_listed.xlsx               # 台灣上市公司清單（1,078 家）
│   └── logs/                        # 每次執行的完整日誌（時間戳命名）
│
├── pdf-cuter/                       # 工具二：PDF 圖表萃取器
│   ├── esg_pdf_cuter.py             # 主程式（~800 行，v2.7）
│   └── logs/                        # 萃取執行日誌
│
├── chart-counter/                   # 工具三：CLIP AI 圖表分類器
│   └── chart_counter.py             # 主程式（~591 行）
│
├── dashboard/                       # 工具四：統一監控面板
│   └── esg-dashboard.py             # 主程式（~671 行）
│
└── data/                            # 所有輸出資料根目錄（大部分被 .gitignore 排除）
    ├── chart_statistics.xlsx        # CLIP 分類結果彙總（11 個 Sheet）
    ├── 2015/
    │   ├── ESG_Download_Progress_2015.xlsx
    │   ├── ESG_Extract_Results_2015.xlsx
    │   ├── 2015_1101_台泥/
    │   │   ├── 2015_1101_台泥.pdf       # .gitignore 排除
    │   │   ├── images/                  # .gitignore 排除（萃取圖片）
    │   │   ├── texts/                   # .gitignore 排除（頁面全文）
    │   │   └── charts/                  # .gitignore 排除（CLIP 確認為圖表）
    │   └── ...
    ├── 2016/
    ├── ...
    ├── 2023/                       
    └── 2024/                       
```

### 模組間資料流

```
tw_listed.xlsx
    │  (1,078 家公司清單)
    ▼
esg_downloader.py
    │  寫出：data/{year}/{sid}_{name}/{sid}_{name}.pdf
    │  寫出：data/{year}/ESG_Download_Progress_{year}.xlsx
    ▼
esg_pdf_cuter.py
    │  讀入：data/{year}/**/*.pdf
    │  寫出：data/{year}/**/images/*.jpg
    │  寫出：data/{year}/**/texts/*.txt
    │  寫出：data/{year}/ESG_Extract_Results_{year}.xlsx
    ▼
chart_counter.py
    │  讀入：data/{year}/**/images/*.jpg
    │  寫出：data/{year}/**/charts/*.jpg（圖表複本）
    │  寫出：data/chart_statistics.xlsx
    ▼
esg-dashboard.py
    │  讀入：所有 ESG_Download_Progress_*.xlsx
    │  讀入：掃描 images/ charts/ garbled_pages.txt
    │  顯示：跨年度進度總覽
```

### 各工具核心設計

#### 工具一：report-downloader

- **爬蟲策略**：Selenium 無頭 Chrome，XPath 動態等待，每 50 家重啟 Chrome 以降低被偵測風險
- **自動防護**：連續 5 家失敗 → 等待解封（最長 2 小時）；網路斷線最多等 30 分鐘；下載卡住 60 秒自動放棄
- **進度持久化**：每家公司處理完立即更新 Excel，支援隨時中斷繼續
- **補抓模式**：啟動時詢問是否重試失敗公司，避免重複處理已成功者

#### 工具二：pdf-cuter

三層圖表偵測演算法，層層過濾：

```
方法一 Raster（嵌入點陣圖）
  ├─ get_images() 取得嵌入圖
  ├─ [A] 過濾 QR code（正方形 + 面積 <9%）
  └─ [B] 過濾全頁背景照片（面積 >80%）

方法二 Vector（向量路徑聚類）
  ├─ get_drawings() 取得路徑
  ├─ Union-Find 聚類（gap=40pt，O(n²)）
  ├─ [C] 過濾頁首/頁尾裝飾線（扁平 + 橫跨）
  └─ [D] 過濾路徑數太少的單一裝飾圖形（<3 條）

方法三 Panel（有色填框）— v2.7 新增
  ├─ 掃描 fill/stroke color 的中型矩形（2.5%～15% 面積）
  ├─ 排除純白無邊框的模板佔位框
  └─ 解決雙頁跨版 PDF 的聚類失效問題
```

#### 工具三：chart-counter

- **模型**：`openai/clip-vit-base-patch32`，零樣本分類（不需要訓練資料）
- **雙 Prompt 策略**：圖表描述 vs 非圖表描述，取 softmax 後的相似度分數
- **分類門檻**：GUI 滑桿可即時調整（0.3～0.9），預設 0.55
- **GPU 加速**：自動偵測 Apple Silicon MPS / NVIDIA CUDA / CPU 降級

#### 工具四：dashboard

- **效能設計**：mtime 指紋快取，只有檔案更新才重算；`os.scandir` 替代 `rglob`（快 10 倍）
- **自動刷新**：後台執行緒每 30 秒掃一次，不阻塞 UI
- **下鑽功能**：點選年度列 → DetailWindow 彈窗，支援搜尋 + 狀態篩選 + 欄位排序

---

## 4. 潛在困境與技術債評估

### 4.1 效能瓶頸

**Union-Find 聚類的 O(n²) 問題**

`_cluster_drawing_rects()` 對頁面每對路徑做兩兩比較。當頁面有 200～700 條向量路徑時（如複雜圖表頁），時間複雜度是 O(n²)。實測每頁最長需 1～2 秒。建議方向：改用空間索引（R-Tree 或 SciPy 的 `cKDTree`）先縮小候選對，可降至 O(n log n)。

**CLIP 模型初始化延遲**

每次啟動 chart-counter，載入 CLIP 模型需要 3～8 秒。若處理大量圖片（如 2024 年 1,040 家），批次推理的 I/O 瓶頸在磁碟讀取（每次開啟 JPEG 都是一次 I/O）。建議：使用 DataLoader 進行非同步預取，減少 GPU 等待 I/O 的空閒時間。

**Dashboard 掃描規模**

2024 年有 1,040 家公司，每次全域掃描要走訪超過 10 萬個檔案節點。雖然已用 `os.scandir` 優化，但 30 秒自動刷新在資料量持續增長下仍可能造成感知延遲。建議改用 `watchdog` 套件進行檔案系統事件監聽，取代輪詢掃描。

---

### 4.2 維護與擴充痛點

**四個程式共享大量重複的 GUI 程式碼**

`set_app_icon()`、Apple 配色常數（`APPLE_BG`、`APPLE_BLUE` 等）、`_make_btn()`、`_open_folder()` 等函式在四個檔案中各自複製一份。一旦需要改動風格（例如改字體或調整按鈕間距），必須在四個地方同步修改，容易漏改。建議抽出為 `shared/ui_toolkit.py` 共享模組。

**執行緒管理使用全域 Event 變數**

各工具使用 `threading.Event()` 全域變數（`pause_event`、`paused_event`、`program_done` 等），狀態散落在模組頂層。若未來需要在同一 session 啟動多個工具，全域狀態會互相汙染。建議改用物件封裝（`class ExtractionSession`）。

**`tw_listed.xlsx` 公司清單的維護責任不明確**

台灣上市公司清單（1,078 家）是整個系統的起點。若公司增刪（上市、下市、更名）沒有更新此檔案，會影響下載完整性。目前沒有自動同步機制，建議定期從 TWSE 公開 API 自動更新。

**macOS 專屬相依**

`pyobjc-core` 和 `pyobjc-framework-Cocoa` 是 macOS 專屬套件，在 Windows 和 Linux 上安裝 `requirements.txt` 會報錯（雖然程式碼有 `try/except` 保護執行期，但安裝期本身就會失敗）。建議拆分 `requirements-mac.txt` 與 `requirements-base.txt`。

---

### 4.3 技術債

**Panel 偵測（v2.7）的精準度存在疑慮**

v2.7 新增的 Panel 偵測方法是針對「卜蜂 PDF 雙頁跨版設計」的緊急修補。它依賴「有色 fill 或 stroke 的中型矩形」這個啟發式規則，在其他廠商的 PDF 上可能產生誤判（例如將有色的文字區塊、表格背景誤認為圖表框）。目前缺乏系統性的測試集來量化偵測精準率，這是最大的技術債。

**跨工具沒有自動化測試**

四個工具都沒有 pytest / unittest。每次修改 PDF 偵測邏輯後，只能手動挑幾份 PDF 驗證，缺乏回歸測試保護。特別是 `_detect_chart_regions()` 的多層過濾邏輯，參數調整後的副作用難以追蹤。建議建立包含 20～30 份代表性 PDF 的黃金測試集。

**絕對路徑殘留問題**

`ESG_Extract_Results_*.xlsx` 的「存檔路徑」欄位儲存了本機的絕對路徑（如 `/Users/kuan/Programming/esg-report/...`）。換機器或換目錄後，這些記錄路徑全部失效。建議改存相對路徑（相對於 `data/` 根目錄）。

**硬編碼的公司總數**

Dashboard 下載進度的百分比計算使用硬編碼常數 `10,780`（1,078 家 × 10 年）。2025 年若增加年份或公司有異動，需要手動修改程式碼才能正確顯示。

---

### 4.4 改善建議

| 優先級 | 改善項目 | 具體做法 |
|--------|---------|---------|
| 🔴 高 | 抽出共用 UI 模組 | 建立 `shared/ui_toolkit.py`，四個工具 `import` 共用 |
| 🔴 高 | 平台相依分離 | 拆分 `requirements-mac.txt` / `requirements-base.txt` |
| 🟡 中 | 建立測試集 | 20～30 份代表性 PDF + pytest，保護 `_detect_chart_regions` |
| 🟡 中 | 改用相對路徑 | Excel 存檔路徑改為 `{year}/{stem}/images/{filename}` |
| 🟡 中 | 動態公司總數 | 從 `tw_listed.xlsx` 動態計算，取代硬編碼常數 |
| 🟢 低 | Union-Find 加速 | 引入空間索引（如 `scipy.spatial.cKDTree`）降低聚類複雜度 |
| 🟢 低 | Dashboard 改事件監聽 | 用 `watchdog` 替代 30 秒輪詢，降低 CPU 閒置用量 |
| 🟢 低 | CLIP 預取優化 | DataLoader 非同步預取圖片，減少 GPU 等待 I/O 的空閒 |

---

## 操作說明

### 環境建置

```bash
git clone git@github.com:leo930206/esg-report.git
cd esg-report
python -m venv .venv

# macOS / Linux
source .venv/bin/activate
pip install -r requirements.txt

# Windows
.venv\Scripts\activate
pip install -r requirements.txt
```

> **注意**：`transformers` + `torch` 首次安裝約 600 MB；CLIP 模型首次執行時另外下載 ~600 MB。

### 下載報告書

```bash
python report-downloader/esg_downloader.py
```

### 萃取圖表

```bash
python pdf-cuter/esg_pdf_cuter.py
```

### cnn 圖表分類

```bash
python chart-counter/chart_counter.py
```

### 查看進度面板

```bash
python dashboard/esg-dashboard.py
```

---

## 資料規模（截至 2026-03）

| 年份 | 公司數 | 備註 |
|------|--------|------|
| 2015 | 268 | |
| 2016 | 329 | |
| 2017 | 344 | |
| 2018 | 364 | |
| 2019 | 403 | |
| 2020 | 431 | |
| 2021 | 493 | |
| 2022 | 639 | |
| 2023 | 725 | |
| 2024 | 1,040 | 規模最大，持續成長 |
| **合計** | **~5,036** | 目標：1,078 家 × 10 年 = 10,780 份 |

## 讀取與下載路徑

### 1. `esg_downloader.py`（下載器）
* **讀取：**
  * `tools/report-downloader/tw_listed.xlsx` — 上市公司清單
  * `data/<year>/ESG_Download_Progress_<year>.xlsx` — 讀取舊進度（斷點續傳）
* **寫入：**
  * `data/<year>/<公司>/2015_1101_台泥.pdf` — 下載的 PDF
  * `data/<year>/ESG_Download_Progress_<year>.xlsx` — 更新下載進度
  * `tools/report-downloader/logs/ESG_Log_*.txt` — 執行日誌

### 2. `esg_pdf_cuter.py`（圖表萃取）
* **讀取：**
  * `data/<year>/<公司>/*.pdf` — 原始 PDF
* **寫入：**
  * `data/<year>/<公司>/images/*.jpg` — 萃取的圖片
  * `data/<year>/<公司>/texts/*.txt` — 每頁文字
  * `data/<year>/<公司>/garbled_pages.txt` — 無法讀取的頁面記錄
  * `data/<year>/ESG_Extract_Results_<year>.xlsx` — 萃取統計

### 3. `chart_counter.py`（圖表計數）
* **讀取：**
  * `data/<year>/<公司>/images/*.jpg` — 萃取的圖片
* **寫入：**
  * `data/<year>/<公司>/charts/*.jpg` — 判定為圖表的圖片（複製）
  * `data/chart_statistics.xlsx` — 各公司圖表數量統計

### 4. `esg-dashboard.py`（主控台）
* **讀取（只讀，不寫入）：**
  * `data/<year>/ESG_Download_Progress_<year>.xlsx` — 下載進度
  * `data/<year>/ESG_Extract_Results_<year>.xlsx` — 萃取統計
  * `data/<year>/<公司>/images/*.jpg` — （僅計算圖片數量）
  * `data/<year>/<公司>/garbled_pages.txt` — 亂碼頁面記錄