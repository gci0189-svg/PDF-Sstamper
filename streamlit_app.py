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
    text_x = st.slider("左右位置 (越小越靠左)", 100, 500, 340)
    text_y = st.slider("上下位置 (越大越靠下)", 10, 200, 20)

    st.header("2️⃣ 印章設定")
    user_stamp = st.file_uploader("上傳透明底印章 (PNG)", type=["png"])
    scale = st.slider("印章大小比例", 0.05, 0.4, 0.12)
    
    st.header("3️⃣ 印章防偽與逼真度")
    base_angle = st.slider("基準傾斜角度", -45, 45, -15)
    use_random = st.checkbox("🎲 開啟隨機微調角度", value=True)

# --- 主畫面：檔案上傳 ---
pdfs = st.file_uploader("請選擇 PDF 工單", type="pdf", accept_multiple_files=True)

if pdfs:
    # 建立表格輸入區
    data = {"檔案名稱":[f.name for f in pdfs], "報價單號": ["" for _ in pdfs]}
    df = pd.DataFrame(data)
    
    st.write("### 📝 請在下方表格輸入每一份檔案的報價單號：")
    edited_df = st.data_editor(df, num_rows="fixed", use_container_width=True)

if pdfs and user_stamp:
    original_pil_img = Image.open(BytesIO(user_stamp.read())).convert("RGBA")
    zip_buffer = BytesIO()
    
    st.write("---")
    st.write("### 👀 第一頁即時預覽 (輸入單號或調位置會自動更新)")
    cols = st.columns(3)
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for i, f in enumerate(pdfs):
            doc = fitz.open(stream=f.getvalue(), filetype="pdf")
            page = doc[0] 
            
            # --- 處理報價單號 ---
            input_number = str(edited_df.iloc[i]['報價單號']).strip()
            
            page.insert_text(fitz.Point(text_x, text_y), "報價單號: ", fontsize=12, fontname="china-ss", color=(0, 0, 0))
            
            if input_number:
                page.insert_text(fitz.Point(text_x + 65, text_y), input_number, fontsize=12, fontname="helv", color=(0, 0, 0))
            
            # --- 處理蓋章 ---
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
            
            # --- 產生預覽圖 ---
            pix = page.get_pixmap(dpi=72)
            preview_bytes = pix.tobytes("png")
            
            with cols[i % 3]:
                with st.expander(f"📄 預覽：{f.name[:10]}..."):
                    st.image(preview_bytes, use_container_width=True)
            
            # --- 存入壓縮檔 (✨ 這裡修改了檔名邏輯) ---
            pdf_out = BytesIO()
            doc.save(pdf_out)
            
            # 1. 拿掉原檔名的 ".pdf" (例如：把 "玉山銀行工單.pdf" 變成 "玉山銀行工單")
            original_name_without_ext = f.name.rsplit('.', 1)[0]
            
            # 2. 組合新檔名：原檔名 + (報價工單) + .pdf
            new_file_name = f"{original_name_without_ext}(報價工單).pdf"
            
            # 3. 寫入壓縮檔
            zip_file.writestr(new_file_name, pdf_out.getvalue())
            doc.close()
            
    st.success("✅ 所有檔案已備妥！請點擊下方按鈕下載：")
    st.download_button("📥 點此下載所有處理後的工單 (ZIP)", zip_buffer.getvalue(), "processed_files.zip", type="primary")
