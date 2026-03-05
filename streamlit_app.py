import streamlit as st
import fitz  # PyMuPDF
from io import BytesIO
import zipfile

# 設定網頁標題
st.set_page_config(page_title="公司自動化蓋章系統", layout="wide")

st.title("⚖️ 綠色辦公：自動化 PDF 蓋章工具")
st.write("上傳多個 PDF 工單，系統將自動按比例於右下角蓋章，並可即時預覽防遮擋。")

# --- 側邊欄控制 ---
with st.sidebar:
    st.header("1️⃣ 印章設定")
    user_stamp = st.file_uploader("上傳印章圖片 (PNG)", type=["png"])
    scale = st.slider("印章大小比例", 0.05, 0.4, 0.12)
    
    st.header("2️⃣ 微調印章位置")
    st.info("若預覽發現印章擋到字，可在此微調位置")
    # 數值越大，邊距越寬，印章就會往左或往上移
    margin_x = st.slider("左右偏移 (數值越大越靠左)", 0.01, 0.3, 0.05)
    margin_y = st.slider("上下偏移 (數值越大越靠上)", 0.01, 0.3, 0.05)

# --- 主畫面：檔案上傳 ---
pdfs = st.file_uploader("請選擇 PDF 工單 (可一次全選多份拖曳上傳)", type="pdf", accept_multiple_files=True)

if pdfs and user_stamp:
    stamp_img = user_stamp.read()
    zip_buffer = BytesIO()
    
    # 提示實際讀取到的數量，解決介面只顯示 3 個的疑慮
    st.success(f"✅ 成功讀取 {len(pdfs)} 份文件！請往下查看預覽：")
    st.write("### 👀 蓋章結果即時預覽")
    
    # 建立 3 個直列來排版預覽圖，畫面比較好看
    cols = st.columns(3)
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for i, f in enumerate(pdfs):
            doc = fitz.open(stream=f.read(), filetype="pdf")
            page = doc[-1] # 選取最後一頁
            
            # 計算自動縮放與位置座標
            W, H = page.rect.width, page.rect.height
            stamp_w = W * scale
            mx = W * margin_x
            my = H * margin_y
            
            # 定義矩形區域: (x1, y1, x2, y2)
            rect = fitz.Rect(W - stamp_w - mx, H - stamp_w - my, W - mx, H - my)
            
            # 蓋上印章
            page.insert_image(rect, stream=stamp_img)
            
            # --- 預覽功能：將蓋完章的 PDF 頁面轉成圖片顯示 ---
            # dpi=72 解析度用來預覽剛剛好，不會讓網頁變卡
            pix = page.get_pixmap(dpi=72)
            img_bytes = pix.tobytes("png")
            
            # 將預覽圖放到畫面上 (使用下拉展開面板，畫面更清爽)
            with cols[i % 3]:
                with st.expander(f"📄 預覽：{f.name[:10]}..."):
                    st.image(img_bytes, use_container_width=True)
            # -----------------------------------------------
            
            # 存入記憶體並壓縮
            pdf_out = BytesIO()
            doc.save(pdf_out)
            zip_file.writestr(f"stamped_{f.name}", pdf_out.getvalue())
            doc.close()
            
    st.write("---")
    # 將下載按鈕變大、變顯眼 (type="primary")
    st.download_button(
        label="📥 確認預覽無誤，點此下載所有蓋章工單 (ZIP壓縮檔)", 
        data=zip_buffer.getvalue(), 
        file_name="stamped_files.zip",
        type="primary"
    )
