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
    st.header("1️⃣ 報價單號位置微調")
    st.info("預設已往左下角移動約 2 公分。若想微調可拉動下方滑桿。")
    # 原始 X 是 400，減 60 約等於往左 2 公分
    text_x = st.slider("左右位置 (越小越靠左)", 100, 500, 340)
    # 原始 Y 是 30，加 60 約等於往下 2 公分
    text_y = st.slider("上下位置 (越大越靠下)", 10, 200, 90)

    st.header("2️⃣ 印章設定")
    user_stamp = st.file_uploader("上傳透明底印章 (PNG)", type=["png"])
    scale = st.slider("印章大小比例", 0.05, 0.4, 0.12)
    
    st.header("3️⃣ 印章防偽與逼真度")
    base_angle = st.slider("基準傾斜角度", -45, 45, -15)
    use_random = st.checkbox("🎲 開啟隨機微調角度", value=True)

# --- 主畫面：檔案上傳 ---
pdfs = st.file_uploader("請選擇 PDF 工單", type="pdf", accept_multiple_files=True)

if pdfs:
    # 建立一個表格，讓使用者針對每個檔案輸入單號
    data = {"檔案名稱":[f.name for f in pdfs], "報價單號": ["" for _ in pdfs]}
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
            
            # --- 完美處理報價單號與字距 ---
            input_number = str(edited_df.iloc[i]['報價單號']).strip()
            
            # 第一段：只印出「報價單號: 」(使用中文字體)
            page.insert_text(fitz.Point(text_x, text_y), "報價單號: ", fontsize=12, fontname="china-ss", color=(0, 0, 0))
            
            # 第二段：印出「輸入的數字」(使用漂亮的標準英文字體 helv，解決字距太寬的問題)
            # X 座標加上 65，剛好接在「報價單號: 」的後面
            if input_number:
                page.insert_text(fitz.Point(text_x + 65, text_y), input_number, fontsize=12, fontname="helv", color=(0, 0, 0))
            
            # --- 蓋章邏輯 ---
            if use_random:
                angle = base_angle + random.randint(-5, 5)
            else:
                angle = base_angle
                
            rotated_img = original_pil_img.rotate(angle, expand=True, fillcolor=(255,255,255,0))
            img_byte_arr = BytesIO()
            rotated_img.save(img_byte_arr, format='PNG')
            
            W, H = page.rect.width, page.rect.height
            stamp_w = W * scale
            rect = fitz.Rect(W - stamp_w - 50, H - stamp_w - 50, W - 50, H - 50)
            page.insert_image(rect, stream=img_byte_arr.getvalue())
            
            # 存檔
            pdf_out = BytesIO()
            doc.save(pdf_out)
            zip_file.writestr(f"processed_{f.name}", pdf_out.getvalue())
            doc.close()
            
    st.success("✅ 處理完成！請點擊下方按鈕下載：")
    st.download_button("📥 下載所有處理後的檔案 (ZIP)", zip_buffer.getvalue(), "processed_files.zip", type="primary")
