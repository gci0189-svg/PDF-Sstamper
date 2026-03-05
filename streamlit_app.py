import streamlit as st
import fitz
import pandas as pd
from io import BytesIO
import zipfile
from PIL import Image
import random

st.set_page_config(page_title="專業工單處理系統", layout="wide")

st.title("⚖️ 自動化處理：蓋章 + 自訂報價單號")

# --- 側邊欄 ---
with st.sidebar:
    st.header("印章設定")
    user_stamp = st.file_uploader("上傳透明底印章 (PNG)", type=["png"])
    scale = st.slider("印章大小比例", 0.05, 0.4, 0.12)
    # 這裡移除報價單號輸入，改到主畫面做表格化輸入

# --- 主畫面：檔案上傳 ---
pdfs = st.file_uploader("請選擇 PDF 工單", type="pdf", accept_multiple_files=True)

if pdfs:
    # 建立一個表格，讓使用者針對每個檔案輸入單號
    data = {"檔案名稱": [f.name for f in pdfs], "報價單號": ["" for _ in pdfs]}
    df = pd.DataFrame(data)
    
    st.write("### 📝 請在下方表格輸入每一份檔案的報價單號：")
    edited_df = st.data_editor(df, num_rows="fixed", use_container_width=True)

if pdfs and user_stamp and st.button("🚀 開始處理"):
    original_pil_img = Image.open(BytesIO(user_stamp.read())).convert("RGBA")
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for i, f in enumerate(pdfs):
            doc = fitz.open(stream=f.read(), filetype="pdf")
            page = doc[-1]
            
            # --- 處理中文字顯示 ---
            # 使用內建的 "han" 字體支援 (這需要系統環境支援，若報錯可改用預設)
            text_to_write = f"報價單號: {edited_df.iloc[i]['報價單號']}"
            page.insert_text(fitz.Point(400, 30), text_to_write, fontsize=12, fontname="china-ss", color=(0, 0, 0))
            
            # --- 蓋章 ---
            angle = random.randint(-20, -10)
            rotated_img = original_pil_img.rotate(angle, expand=True, fillcolor=(255,255,255,0))
            img_byte_arr = BytesIO()
            rotated_img.save(img_byte_arr, format='PNG')
            
            W, H = page.rect.width, page.rect.height
            stamp_w = W * scale
            rect = fitz.Rect(W - stamp_w - 50, H - stamp_w - 50, W - 50, H - 50)
            page.insert_image(rect, stream=img_byte_arr.getvalue())
            
            pdf_out = BytesIO()
            doc.save(pdf_out)
            zip_file.writestr(f"processed_{f.name}", pdf_out.getvalue())
            doc.close()
            
    st.success("✅ 處理完成！")
    st.download_button("📥 下載所有處理後的檔案", zip_buffer.getvalue(), "processed_files.zip", type="primary")
