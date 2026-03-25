import streamlit as st
import fitz  # PyMuPDF
import os
import pandas as pd
from pathlib import Path
from PIL import Image
import io

# --- 路徑與設定区 ---
BASE_DIR = Path(__file__).parent.absolute()
DATA_DIR = BASE_DIR.parent / "data"          # 統一輸出根目錄
DATA_DIR.mkdir(exist_ok=True)

def year_dir(year: str) -> Path:
    return DATA_DIR / str(year)

def year_excel(year: str) -> Path:
    """每個年份各自的萃取統計 Excel 路徑"""
    return year_dir(year) / f"ESG_Extract_Results_{year}.xlsx"

def _year_range():
    return range(2015, 2025)

def available_years():
    if not DATA_DIR.is_dir():
        return [str(y) for y in _year_range()]
    dirs = [d for d in os.listdir(DATA_DIR)
            if (DATA_DIR / d).is_dir() and d.isdigit()]
    return sorted(dirs) or [str(y) for y in _year_range()]

# 確保輸出目錄存在
if not os.path.exists(GRAPH_DIR):
    os.makedirs(GRAPH_DIR)

# --- 進階圖表擷取函數 ---
def extract_assets_from_pdf(pdf_path, year):
    doc = fitz.open(pdf_path)
    file_name = Path(pdf_path).stem
    specific_output_dir = GRAPH_DIR / year / file_name
    os.makedirs(specific_output_dir, exist_ok=True)
    
    results = []
    asset_counter = 0
    
    for page_index, page in enumerate(doc):
        page_num = page_index + 1
        page_rect = page.rect
        page_area = page_rect.width * page_rect.height
        
        # 準備區域偵測容器
        candidate_rects = []
        
        # 方法一：偵測傳統點陣圖片 (Raster Images)
        image_list = page.get_images(full=True)
        for img_info in image_list:
            xref = img_info[0]
            # 取得圖片在頁面上的實際顯示矩形
            img_rects = page.get_image_rects(xref)
            for r in img_rects:
                # 過濾太小的 icon
                if r.width > 50 and r.height > 50:
                    candidate_rects.append({'rect': r, 'type': 'Raster_Image'})

        # 方法二：偵測向量圖形聚集區域 (Vector Drawings) - 關鍵修正
        paths = page.get_drawings()
        if len(paths) > 10: # 路徑數夠多才視為有圖表
            # 取得所有重要向量路徑的矩形
            drawing_rects = [p["rect"] for p in paths if p["rect"].width > 5 and p["rect"].height > 5]
            
            if drawing_rects:
                # 計算所有向量路徑的聯集範圍
                combined_rect = drawing_rects[0]
                for r in drawing_rects[1:]:
                    combined_rect |= r
                
                # --- 智慧擴張 (Smart Expansion) ---
                # 這是解決「圖二」問題的關鍵：將偵測到的圖形框向外擴張
                # 以確保包覆上方的數字、下方的標籤、旁邊的圖例
                # 我們向四周擴張 80 像素 (可根據實際情況調整)
                expanded_rect = combined_rect + (-80, -80, 80, 80)
                
                # 確保擴張後的框不會超出頁面邊界
                expanded_rect &= page_rect
                
                if expanded_rect.width > 100 and expanded_rect.height > 100:
                    # 避免抓到整頁的背景邊框
                    if expanded_rect.width < page_rect.width * 0.95:
                        candidate_rects.append({'rect': expanded_rect, 'type': 'Vector_Chart_Area'})

        # --- 對所有候選區域進行高解析度裁切 ---
        # 為了避免重疊，進行簡單的去重 (可選)
        
        for cand in candidate_rects:
            r = cand['rect']
            area_pct = round((r.width * r.height / page_area) * 100, 4)
            
            # 再次過濾極小或極大的雜訊區域
            if area_pct < 0.5 or area_pct > 90: continue

            asset_counter += 1
            # 類型代碼：RA (Raster), VC (Vector Chart)
            type_code = "RA" if cand['type'] == 'Raster_Image' else "VC"
            img_name = f"{file_name}_p{page_num}_{asset_counter}_{type_code}.png"
            img_save_path = specific_output_dir / img_name
            
            # 提升渲染解析度 (Matrix 3x3 = 3倍解析度，約 216 DPI)
            # 這能確保 CNN 看得清文字和數據
            mat = fitz.Matrix(3, 3) 
            pix = page.get_pixmap(matrix=mat, clip=r, alpha=False)
            
            # 執行存檔
            pix.save(str(img_save_path))
            
            # 記錄數據
            results.append({
                "年份": year,
                "原始PDF檔名": file_name,
                "PDF總頁數": len(doc),
                "圖片所在頁碼": page_num,
                "圖片累積編號": asset_counter,
                "圖片存檔名稱": img_name,
                "圖片面積占比(%)": area_pct,
                "類型": cand['type'],
                "存檔路徑": str(img_save_path)
            })
            
            pix = None # 釋放記憶體
            
    doc.close()
    return results

# --- Streamlit UI 介面 (保持不變) ---
st.set_page_config(page_title="永續報告書智慧圖表裁切系統", layout="wide")
st.title("📊 ESG 報告書數據圖形自動萃取工具 (CNN 前處理)")
st.markdown("""
本工具專為學術研究設計，解決以下痛點：
1. **完整擷取缺乏外框的向量圖表** (圓餅圖、長條圖)。
2. **自動擴張範圍**，確保數據 (1.50, 209%) 與標籤 (無添加) 不遺漏。
3. **高解析度存檔** (216 DPI)，提升 CNN 辨識率。
""")

# 自動偵測年份資料夾
exclude = ['graph', '__pycache__', 'venv', '.git']
years = sorted([d for d in os.listdir(BASE_DIR) if os.path.isdir(BASE_DIR / d) and d not in exclude and d.isdigit()])

if not years:
    st.error(f"⚠️ 在目錄 {BASE_DIR} 下找不到年份資料夾 (如 2024)。請確認原始 PDF 存放位置。")
    st.stop()

selected_year = st.sidebar.selectbox("請選擇要處理的年份", years)

# 讀取進度紀錄 (斷點續傳)
processed_files = set()
if os.path.exists(LOG_FILE):
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        processed_files = set(line.strip() for line in f.readlines())

# 掃描目標檔案
target_path = BASE_DIR / selected_year
all_pdfs = [str(target_path / f) for f in os.listdir(target_path) if f.lower().endswith(".pdf")]
pending_pdfs = [f for f in all_pdfs if f not in processed_files]

# 狀態顯示儀表板
c1, c2, c3 = st.columns(3)
c1.metric("當前年份", selected_year)
c2.metric("待處理 PDF 數量", len(pending_pdfs))
c3.metric("已完成處理", len(processed_files))

# 執行按鈕
if st.button(f"開始執行 {selected_year} 年數據萃取"):
    if not pending_pdfs:
        st.warning("所有檔案皆已處理完成！")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        all_results_data = []

        # 讀取現有的 Excel 以便追加資料
        if os.path.exists(DATA_EXPORT):
            try:
                all_results_data = pd.read_excel(DATA_EXPORT).to_dict('records')
            except:
                st.error("讀取 Excel 時發生錯誤，將重新建立。")

        for i, pdf_path in enumerate(pending_pdfs):
            fname = os.path.basename(pdf_path)
            status_text.text(f"正在分析 ({i+1}/{len(pending_pdfs)}): {fname}")
            
            try:
                # 執行進階萃取
                file_data = extract_assets_from_pdf(pdf_path, selected_year)
                all_results_data.extend(file_data)
                
                # 寫入斷點紀錄
                with open(LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(f"{pdf_path}\n")
                
                # 每處理完一個檔案就即時更新 Excel
                df = pd.DataFrame(all_results_data)
                df.to_excel(DATA_EXPORT, index=False)
                
                progress_bar.progress((i + 1) / len(pending_pdfs))
            except Exception as e:
                st.error(f"跳過錯誤檔案 {fname}: {str(e)}")
                continue

        st.balloons()
        st.success(f"✅ {selected_year} 年年份處理完畢！")
        st.write(f"所有圖片已根據 PDF 檔名分類存入 `graph/{selected_year}/` 資料夾中。")
        st.write(f"數據統計 Excel 詳見：`{DATA_EXPORT}`")

# 5. 預覽數據
if os.path.exists(DATA_EXPORT):
    with st.expander("查看目前已累積的 Excel 處理紀錄 (前 100 筆)"):
        st.dataframe(pd.read_excel(DATA_EXPORT).head(100))