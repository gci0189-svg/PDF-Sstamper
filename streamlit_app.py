import streamlit as st
import fitz  # PyMuPDF
from io import BytesIO
import zipfile

# 設定網頁標題
st.set_page_config(page_title="公司自動化蓋章系統", layout="wide")

st.title("⚖️ 綠色辦公：自動化 PDF 蓋章工具")
st.write("上傳多個 PDF 工單，系統將自動按比例於右下角蓋章。")

# --- 側邊欄控制 ---
with st.sidebar:
    st.header("印章設定")
    # 讓使用者也可以動態換印章
    user_stamp = st.file_uploader("上傳印章圖片 (PNG)", type=["png"])
    # 設定縮放比例 (例如佔頁面寬度的 15%)
    scale = st.slider("印章縮放比例", 0.1, 0.3, 0.15)

# --- 主畫面：檔案上傳 ---
pdfs = st.file_uploader("請選擇 PDF 工單 (可多選)", type="pdf", accept_multiple_files=True)

if pdfs and user_stamp:
    stamp_img = user_stamp.read()
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for f in pdfs:
            doc = fitz.open(stream=f.read(), filetype="pdf")
            page = doc[-1] # 選取最後一頁
            
            # 計算自動縮放座標
            # W: 頁面寬度, H: 頁面高度
            W, H = page.rect.width, page.rect.height
            stamp_w = W * scale
            margin = W * 0.05
            
            # 定義矩形區域: (x1, y1, x2, y2)
            rect = fitz.Rect(W - stamp_w - margin, H - stamp_w - margin, W - margin, H - margin)
            
            page.insert_image(rect, stream=stamp_img)
            
            # 存入記憶體並壓縮
            pdf_out = BytesIO()
            doc.save(pdf_out)
            zip_file.writestr(f"stamped_{f.name}", pdf_out.getvalue())
            doc.close()
            
    st.success(f"✅ 處理完成！共計 {len(pdfs)} 份。")
    st.download_button("📥 下載所有蓋章工單 (ZIP)", zip_buffer.getvalue(), "stamped_files.zip")
