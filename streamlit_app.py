import streamlit as st
import fitz  # PyMuPDF
from io import BytesIO
import zipfile
from PIL import Image  # 用來處理圖片旋轉
import random          # 用來產生隨機角度

# 設定網頁標題
st.set_page_config(page_title="公司自動化蓋章系統", layout="wide")

st.title("⚖️ 綠色辦公：自動化 PDF 蓋章工具 (逼真版)")
st.write("上傳多個 PDF 工單，系統將自動於右下角蓋章，並支援角度傾斜與隨機微調。")

# --- 側邊欄控制 ---
with st.sidebar:
    st.header("1️⃣ 印章設定")
    user_stamp = st.file_uploader("上傳透明底印章 (PNG)", type=["png"])
    scale = st.slider("印章大小比例", 0.05, 0.4, 0.12)
    
    st.header("2️⃣ 微調位置")
    margin_x = st.slider("左右偏移 (越大越靠左)", 0.01, 0.3, 0.05)
    margin_y = st.slider("上下偏移 (越大越靠上)", 0.01, 0.3, 0.05)
    
    st.header("3️⃣ 逼真度設定 (旋轉)")
    st.info("讓印章歪歪的，看起來更像真人蓋的！")
    # 讓使用者決定基準角度
    base_angle = st.slider("基準傾斜角度 (負數向右，正數向左)", -45, 45, -15)
    # 神級功能：隨機微調
    use_random = st.checkbox("🎲 開啟隨機微調角度", value=True, help="打勾後，每份文件的印章都會在您設定的角度上，再隨機加減 5 度，看起來超級逼真！")

# --- 主畫面：檔案上傳 ---
pdfs = st.file_uploader("請選擇 PDF 工單 (可一次全選多份拖曳上傳)", type="pdf", accept_multiple_files=True)

if pdfs and user_stamp:
    # 讀取使用者上傳的印章並轉成 PIL 圖片格式
    original_pil_img = Image.open(BytesIO(user_stamp.read())).convert("RGBA")
    
    zip_buffer = BytesIO()
    st.success(f"✅ 成功讀取 {len(pdfs)} 份文件！請往下查看預覽：")
    st.write("### 👀 蓋章結果即時預覽 (注意看每張角度)")
    
    cols = st.columns(3)
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for i, f in enumerate(pdfs):
            doc = fitz.open(stream=f.read(), filetype="pdf")
            page = doc[-1] # 選取最後一頁
            
            # --- 處理印章旋轉 ---
            # 如果有開啟隨機，就在基準角度上加減 -5 到 +5 度
            if use_random:
                current_angle = base_angle + random.randint(-5, 5)
            else:
                current_angle = base_angle
                
            # 旋轉圖片 (expand=True 確保旋轉後圖片邊角不會被裁切)
            rotated_img = original_pil_img.rotate(current_angle, expand=True, fillcolor=(255,255,255,0))
            
            # 把旋轉後的圖片轉回位元組流 (Bytes) 讓 PyMuPDF 讀取
            img_byte_arr = BytesIO()
            rotated_img.save(img_byte_arr, format='PNG')
            final_stamp_stream = img_byte_arr.getvalue()
            # ---------------------
            
            # 計算座標
            W, H = page.rect.width, page.rect.height
            stamp_w = W * scale
            mx = W * margin_x
            my = H * margin_y
            
            rect = fitz.Rect(W - stamp_w - mx, H - stamp_w - my, W - mx, H - my)
            
            # 蓋上旋轉後的印章
            page.insert_image(rect, stream=final_stamp_stream)
            
            # 產生預覽圖
            pix = page.get_pixmap(dpi=72)
            preview_bytes = pix.tobytes("png")
            
            with cols[i % 3]:
                # 在標題顯示這張被旋轉了幾度
                with st.expander(f"📄 預覽 (傾斜 {current_angle} 度)"):
                    st.image(preview_bytes, use_container_width=True)
            
            # 存檔
            pdf_out = BytesIO()
            doc.save(pdf_out)
            zip_file.writestr(f"stamped_{f.name}", pdf_out.getvalue())
            doc.close()
            
    st.write("---")
    st.download_button(
        label="📥 確認預覽無誤，點此下載所有蓋章工單 (ZIP)", 
        data=zip_buffer.getvalue(), 
        file_name="stamped_files.zip",
        type="primary"
    )
