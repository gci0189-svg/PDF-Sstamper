import streamlit as st
import fitz  # PyMuPDF
from io import BytesIO
import zipfile
from PIL import Image
import random

st.set_page_config(page_title="公司自動化蓋章系統", layout="wide")

st.title("⚖️ 綠色辦公：自動化 PDF 蓋章工具 (含報價單號)")

# --- 側邊欄控制 ---
with st.sidebar:
    st.header("1️⃣ 報價單資訊")
    # 新增文字輸入框
    quote_number = st.text_input("請輸入報價單號", placeholder="例如：QT-20260305")
    
    st.header("2️⃣ 印章設定")
    user_stamp = st.file_uploader("上傳透明底印章 (PNG)", type=["png"])
    scale = st.slider("印章大小比例", 0.05, 0.4, 0.12)
    
    st.header("3️⃣ 逼真度設定")
    base_angle = st.slider("基準傾斜角度", -45, 45, -15)
    use_random = st.checkbox("🎲 開啟隨機微調角度", value=True)

# --- 主畫面 ---
pdfs = st.file_uploader("請選擇 PDF 工單", type="pdf", accept_multiple_files=True)

if pdfs and user_stamp:
    original_pil_img = Image.open(BytesIO(user_stamp.read())).convert("RGBA")
    zip_buffer = BytesIO()
    
    st.success(f"✅ 成功讀取 {len(pdfs)} 份文件！")
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for f in pdfs:
            doc = fitz.open(stream=f.read(), filetype="pdf")
            page = doc[-1] # 最後一頁
            
            # --- 寫入報價單號 (在右上角) ---
            if quote_number:
                # 設定文字位置 (x=450, y=50)，視您的 PDF 格式可能需微調
                text_point = fitz.Point(450, 50) 
                page.insert_text(text_point, f"報價單號: {quote_number}", fontsize=12, color=(0, 0, 0))
            
            # --- 蓋章邏輯 ---
            if use_random:
                current_angle = base_angle + random.randint(-5, 5)
            else:
                current_angle = base_angle
                
            rotated_img = original_pil_img.rotate(current_angle, expand=True, fillcolor=(255,255,255,0))
            img_byte_arr = BytesIO()
            rotated_img.save(img_byte_arr, format='PNG')
            
            W, H = page.rect.width, page.rect.height
            stamp_w = W * scale
            rect = fitz.Rect(W - stamp_w - 50, H - stamp_w - 50, W - 50, H - 50)
            
            page.insert_image(rect, stream=img_byte_arr.getvalue())
            
            # 存檔
            pdf_out = BytesIO()
            doc.save(pdf_out)
            zip_file.writestr(f"stamped_{f.name}", pdf_out.getvalue())
            doc.close()
            
    st.download_button("📥 下載所有蓋章並標註單號的工單", zip_buffer.getvalue(), "stamped_files.zip", type="primary")
