import streamlit as st

# ==========================================
# 1. 網頁基礎配置
# ==========================================
st.set_page_config(page_title="數位稽核工作站", layout="wide", page_icon="🏢")

# ==========================================
# 2. 隱藏預設側邊欄的魔法 CSS
# ==========================================
st.markdown("""
    <style>
        /* 隱藏 Streamlit 預設在最上方的檔案清單導覽列 */
        [data-testid="stSidebarNav"] {display: none;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. 打造客製化側邊欄 (精準控制排版順序)
# ==========================================
st.sidebar.title("🏢 數位稽核工作站")
st.sidebar.caption("Digital Audit Workstation")
st.sidebar.markdown("---")
st.sidebar.subheader("🛠️ 工具列表")

# 手動建立導覽按鈕
st.sidebar.page_link("app.py", label="🏠 主控台首頁")
st.sidebar.page_link("pages/1_Word圖片提取.py", label="📝 Word 圖片提取")
st.sidebar.page_link("pages/2_PDF圖片提取.py", label="📄 PDF 圖片提取")
st.sidebar.page_link("pages/3_PDF與圖檔合併器.py", label="📑 PDF 與圖檔合併器")
st.sidebar.page_link("pages/4_萬能格式轉PDF_直式.py", label="📄 萬能轉 PDF (直式)")
st.sidebar.page_link("pages/5_萬能格式轉PDF_橫式.py", label="📄 萬能轉 PDF (橫式)")
st.sidebar.page_link("pages/6_CSA智能整併系統.py", label="📊 CSA 智能整併系統")
st.sidebar.page_link("pages/7_智能一頁式摘要生成器.py", label="🤖 智能一頁式摘要生成器") # 👈 第7項已完美列入側邊欄

# ==========================================
# 4. 主畫面設計
# ==========================================
st.title("數位稽核工作站 ⚡")
st.write("本平台專為內部控制與稽核作業設計，旨在將重複性的底稿素材處理流程全自動化。")
st.divider()

# --- 第一模塊：底稿素材處理 ---
st.subheader("📁 底稿素材處理")
col1, col2, col3 = st.columns(3)

with col1:
    st.info("📝 **Word 圖片提取**")
    st.write("支援 .doc/.docx，自動按順序擷取原始圖片並轉存 PNG。")
    if st.button("進入工具 👉", key="nav_word"):
        st.switch_page("pages/1_Word圖片提取.py")

with col2:
    st.success("📄 **PDF 圖片精準提取**")
    st.write("自動掃描 PDF 檔案，精準提取內嵌原圖，並依循「頁碼與順序」自動命名歸檔。")
    if st.button("進入工具 👉", key="nav_pdf"):
        st.switch_page("pages/2_PDF圖片提取.py")

with col3:
    st.warning("📑 **PDF 與圖檔合併器**")
    st.write("支援表格互動排版，將多來源圖檔與文件裝訂為單一底稿。")
    if st.button("進入工具 👉", key="nav_merge"):
        st.switch_page("pages/3_PDF與圖檔合併器.py")

# --- 第二模塊：萬能底稿轉檔 ---
st.markdown("---")
st.subheader("🔄 萬能底稿轉檔 (格式自動最佳化)")

col_v, col_h = st.columns(2) 

with col_v:
    st.info("📄 **萬能轉 PDF (直式單頁)**")
    st.write("Word/Excel/圖檔轉 PDF。Excel 自動優化欄寬、重複標題列，並強制縮放為**【直向單頁】**。")
    if st.button("進入直式工具 👉", key="nav_conv_v"):
        st.switch_page("pages/4_萬能格式轉PDF_直式.py")

with col_h:
    st.success("📄 **萬能轉 PDF (橫式單頁)**")
    st.write("Word/Excel/圖檔轉 PDF。Excel 自動優化欄寬、重複標題列，並強制縮放為**【橫向單頁】**。")
    if st.button("進入橫式工具 👉", key="nav_conv_h"):
        st.switch_page("pages/5_萬能格式轉PDF_橫式.py")

# --- 第三模塊：進階稽核分析與監控 ---
st.markdown("---")
st.subheader("🔍 進階稽核分析與監控")

# 💡 核心關鍵修復：明確開出左右兩個獨立欄位空間
col_csa, col_ai = st.columns(2)

with col_csa:
    st.error("📊 **內控自評 (CSA) 智能整併系統**")
    st.write("自動載入總表骨架，透過『唯一 KEY 值』跨表對齊多部門回填檔，並具備紅旗異常智能預警功能。")
    if st.button("進入整併系統 👉", key="nav_csa"):
        st.switch_page("pages/6_CSA智能整併系統.py")

with col_ai:
    # 使用帶有專業科技感的客製化邊框與底色排版
    st.markdown("""
    <div style="background-color: #f3e8ff; border: 1px solid #d8b4fe; padding: 16px; border-radius: 8px; height: 100%;">
        <span style="color: #6b21a8; font-weight: bold; font-size: 16px;">🤖 智能 AI 稽核助理</span><br>
        <span style="color: #4c1d95; font-size: 14px;">串接 Gemini 特化輕量大腦，全自動解鎖資料並套用精緻 HTML/CSS 骨架，一鍵生成高階主管專屬一頁式摘要。</span>
    </div>
    """, unsafe_allow_html=True)
    st.write("") # 留一點呼吸間距
    if st.button("呼叫 AI 助理 👉", key="nav_ai"):
        st.switch_page("pages/7_智能一頁式摘要生成器.py")

st.divider()
st.caption("Digital Audit Workstation © 2026")
