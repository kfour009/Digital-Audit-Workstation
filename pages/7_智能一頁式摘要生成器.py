import streamlit as st
import google.generativeai as genai
import json

# ==========================================
# 網頁基礎配置與全域樣式收斂
# ==========================================
st.set_page_config(page_title="智能一頁式摘要生成器", layout="wide")

st.markdown("""
    <style>
        html, body, [class*="st-"] { font-size: 16px !important; }
        .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
        .stTextArea textarea { line-height: 1.5; }
    </style>
""", unsafe_allow_html=True)

st.title("🤖 智能 AI 稽核助理：一頁式視覺戰情報告")
st.write("全自動生成頂級視覺底稿。嚴格鎖定標題元資料完整性，完美支援「零缺失/純合規」專案動態渲染，一鍵導出滿版去背實體 PDF 檔案。")

# ==========================================
# 暫存狀態 (Session State) 初始化
# ==========================================
if 'ai_generated_data' not in st.session_state:
    st.session_state.ai_generated_data = None
if 'report_generated' not in st.session_state:
    st.session_state.report_generated = False
if 'trigger_api_call' not in st.session_state:
    st.session_state.trigger_api_call = False

def clean_html(html_str):
    return "\n".join([line.strip() for line in html_str.split("\n")])

# ==========================================
# 側邊欄 / 頂部：BYOK 安全金鑰模組與新手指南
# ==========================================
st.subheader("🔑 系統授權連線")

api_key = st.text_input(
    "請輸入您的 Google Gemini API Key 啟動 AI 引擎：", 
    type="password", 
    placeholder="請貼上以 AIza 開頭的專屬金鑰...",
    help="為保障系統公用配額穩定，請使用您個人的金鑰連線。金鑰僅於當前瀏覽器記憶體暫存，關閉網頁即自動銷毀，絕對安全。"
)

with st.expander("💡 第一次使用？點此查看如何免費取得專屬 API Key (3 分鐘快速完成)"):
    st.markdown("""
    本工作站串接 Google 最新推出的特化型輕量大腦 **Gemini 3.1 Flash Lite**，您的免費 Google 帳號每日預設享有高達 **500 次** 的極速呼叫配額，完全足夠日常查核業務使用！
    
    **跟著以下步驟輕鬆取得連線金鑰：**
    1. 點擊前往 👉 **[Google AI Studio 官方控制台](https://aistudio.google.com/)** 並使用您的企業或個人 Google 帳號登入。
    2. 進入畫面後，點擊左上角的 **「Get API key」** 選單。
    3. 點擊畫面中央的藍色按鈕 **「Create API key」**，系統會詢問綁定專案，選擇 **「Create API key in new project」** 即可。
    4. 等待幾秒鐘，系統會彈出一串以 **`AIza`** 開頭的亂碼字串，點擊右側的複製圖示。
    5. 返回本數位工作站，將金鑰貼入上方的密碼欄位中，即可永久解鎖全自動排版引擎！
    """)
st.divider()

# ==========================================
# 主畫面操作區：表單設定
# ==========================================
st.subheader("📌 報告標題設定")
col_t1, col_t2, col_t3 = st.columns(3)
with col_t1:
    audit_period = st.text_input("📅 查核期間", value="", placeholder="範例：2026Q1")
with col_t2:
    audit_target = st.text_input("🏢 受查單位", value="", placeholder="範例：運籌採購總處")
with col_t3:
    audit_item = st.text_input("📑 查核項目", value="", placeholder="範例：採購與付款循環控制作業")

st.markdown("---")

col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("無缺失事項 (執行查核事項)")
    pass_text = st.text_area(
        "請直接貼上所有已執行且無異常之查核程序 (各點請斷行或以數字開頭)：", 
        value="", 
        placeholder="背景範例貼入格式：\n1. 隨機抽核本季度25份供應商建檔審查表，確認均經權責主管核准並檢附完整資格證明文件。\n2. 比對ERP系統中採購單與核決權限表，系統均能依據採購金額自動觸發對應層級之電子簽核流程。\n3. 實地觀察倉儲驗收區域，點收人員均確實核對送貨單與採購單明細，並於系統即時產生驗收單據。", 
        height=220,
        label_visibility="collapsed"
    )

with col_right:
    st.subheader("有缺失事項 (實務觀察與優化建議)")
    issue_text = st.text_area(
        "請直接貼上所有發現之異常初稿或缺失描述 (若本次查核無缺失請直接留白)：", 
        value="", 
        placeholder="背景範例貼入格式 (若無缺失請保持空白)：\n1. 經查有部分特定原物料採購案，因業務端緊急需求未於事前完成請購單申請，即先通知廠商出貨，後續方進行單據補簽，存在作業流程繞道風險。\n2. 盤點長期未異動之閒置資產清冊，發現有兩項生產設備已逾耐用年限且無使用紀錄，然保管單位尚未啟動資產報廢或處分評估程序，影響資產管理效益。", 
        height=220,
        label_visibility="collapsed"
    )
        
st.markdown("<br>", unsafe_allow_html=True)
st.info("⚙️ 報告呈現方向設定 (動態 Prompt 注入)")
col_s1, col_s2 = st.columns(2)
with col_s1:
    tone_setting = st.selectbox("🎯 呈報對象語氣要求", [
        "精煉大局觀 (公司整體管理層級，強調大局衝擊與治理維度)", 
        "建設性導向 (友善受查單位溝通，強調流程優化與共好防禦)"
    ])
with col_s2:
    focus_setting = st.selectbox("🚩 核心側重點", [
        "著重合規與內控防禦深度 (強化控制點嚴謹度/舉證軌跡)",
        "著重營運持續與流程效率 (優化作業瓶頸/降本增效)"
    ])

# ==========================================
# 🚀 啟動 AI 呼叫引擎 (多重合規握手防護鎖)
# ==========================================
if st.button("🚀 呼叫 AI 深度生成視覺戰情報告", type="primary"):
    # 💡 終極業務邊界防護：依序檢驗授權、標題元資料完整性與內容實質性
    if not api_key:
        st.error("❌ 系統攔截：請先於上方貼入您的專屬 Gemini API Key！")
    elif not audit_period.strip() or not audit_target.strip() or not audit_item.strip():
        # 💡 新增攔截：標題三要素絕對不可留白
        st.warning("⚠️ 系統攔截：請完整填寫上方『報告標題設定』的三個欄位（查核期間、受查單位、查核項目不可留白）！")
    elif not pass_text.strip() and not issue_text.strip():
        st.warning("⚠️ 系統攔截：請至少填入『無缺失事項』或『有缺失事項』其中一邊的查核紀錄！")
    else:
        st.session_state.trigger_api_call = True

# 快取防禦引擎實際運作區
if st.session_state.trigger_api_call:
    with st.spinner("✨ 智能助理正調用極速 3.1 Flash Lite 特化大腦進行真實語義萃取與排版，請稍候..."):
        try:
            for k in list(st.session_state.keys()):
                if k.startswith("edit_h_") or k.startswith("edit_i_"):
                    del st.session_state[k]

            genai.configure(api_key=api_key)
            model = genai.GenerativeModel(
                'gemini-3.1-flash-lite',
                generation_config={"response_mime_type": "application/json", "temperature": 0.1}
            )

            eff_pass = pass_text.strip()
            eff_iss = issue_text.strip()

            prompt = f"""
            請身為資深企業內部稽核主管，嚴格依循 JSON 輸出審查結果。
            語氣要求：{tone_setting} / 側重點：{focus_setting}
            
            【最高審查準則：絕對忠實反映留白與實質分流】
            請嚴格遵守以下業務規則：
            1. 若下方提供的「任務輸入文本」為空白或無任何文字，對應的產出陣列請務必直接回傳空陣列 `[]`，絕對不可自行編造假資料！
            2. 若文本中有實質文字，請主動檢視獨立段落的真實語義本質，自動將合規事實分流至 highlights，將管理瑕疵分流至 issues。

            【任務一輸入文本 (執行查核事項)】
            {eff_pass if eff_pass else "無輸入內容。"}

            【任務二輸入文本 (有缺失事項)】
            {eff_iss if eff_iss else "無輸入內容。"}

            請產出標準 JSON 結構：
            - highlights 陣列：針對每個實質合規項目提煉 1. 主旨大標題(不超過12字) 2. 合規描述(約30-45字)。
            - issues 陣列：針對每個實質異常項目獨立剖析 1. 議題標題(不超過14字) 2. 實務觀察 3. 優化建議。

            嚴格輸出格式：
            {{"highlights": [{{"title": "大標題", "desc": "合規描述"}}], "issues": [{{"title": "議題標題", "observation": "觀察內容", "recommendation": "建議內容"}}]}}
            """
            
            res = model.generate_content(prompt)
            st.session_state.ai_generated_data = json.loads(res.text)
            st.session_state.report_generated = True
            
        except Exception as e:
            st.error(f"❌ 呼叫異常：{str(e)}")
            st.info("💡 排解提示：請確認金鑰無誤，或檢查網路連線狀態。")
        finally:
            st.session_state.trigger_api_call = False

# ==========================================
# 📊 報告優先展示與自適應編修區 (前端快取響應)
# ==========================================
if st.session_state.report_generated and st.session_state.ai_generated_data:
    st.divider()
    
    data = st.session_state.ai_generated_data
    
    # 動態抓取實質標題內容 (結合前台攔截把關，保證字串絕對乾淨就位)
    p_val = audit_period.strip()
    t_val = audit_target.strip()
    i_val = audit_item.strip()
    full_title = f"{p_val} {t_val} {i_val}：整體查核總覽"

    icons = ["📋", "🔍", "🛡️", "📊", "⚙️", "💡", "📑", "📌"]
    highlights = data.get("highlights", [])
    issues = data.get("issues", [])

    total_items = len(highlights) + len(issues)
    is_compact = total_items > 6

    card_p_top = "8px" if is_compact else "10px"
    card_p_side = "12px" if is_compact else "14px"
    font_title = "16px" if is_compact else "18px"
    font_body = "14px" if is_compact else "15px"
    margin_b = "12px" if is_compact else "24px"
    issue_p = "14px" if is_compact else "18px"
    issue_mb = "8px" if is_compact else "12px"

    highlights_html = ""
    if highlights:
        for idx, h in enumerate(highlights):
            icon = icons[idx % len(icons)]
            h_title = st.session_state.get(f"edit_h_t_{idx}", h.get('title', ''))
            h_desc = st.session_state.get(f"edit_h_d_{idx}", h.get('desc', ''))
            
            highlights_html += f"""
            <div style="flex: 1; min-width: 200px; border: 1px solid #cbd5e1; border-radius: 6px; overflow: hidden; background: #ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.05); display: flex; flex-direction: column;">
                <div style="background-color: #0284c7; color: #ffffff; padding: {card_p_top} {card_p_side}; font-weight: 700; font-size: {font_title}; display: flex; align-items: center; gap: 6px;">
                    <span>{icon}</span> <span>{h_title}</span>
                </div>
                <div style="padding: {card_p_side}; flex-grow: 1;">
                    <p style="color: #334155; font-size: {font_body}; line-height: 1.5; margin: 0;">{h_desc}</p>
                </div>
            </div>
            """
    else:
        highlights_html = """
        <div style="background-color: #f8fafc; border: 1px dashed #cbd5e1; padding: 12px 20px; border-radius: 6px; width: 100%; color: #64748b; font-size: 15px;">
            ℹ️ 本次報告未記錄常規遵循事項。
        </div>
        """
        
    issues_html = ""
    if issues:
        for idx, item in enumerate(issues):
            i_title = st.session_state.get(f"edit_i_t_{idx}", item.get('title', ''))
            i_obs = st.session_state.get(f"edit_i_o_{idx}", item.get('observation', ''))
            i_rec = st.session_state.get(f"edit_i_r_{idx}", item.get('recommendation', ''))
            
            issues_html += f"""
            <div style="flex: 1; min-width: 300px; border: 1px solid #fde68a; border-radius: 8px; background-color: #fffbeb; padding: {issue_p}; margin-bottom: {issue_mb}; box-shadow: 0 2px 4px rgba(217,119,6,0.05);">
                <div style="border-bottom: 2px solid #f59e0b; padding-bottom: 8px; margin-bottom: 10px; display: flex; align-items: center; gap: 6px;">
                    <span style="font-size: 18px;">💡</span>
                    <span style="font-size: {font_title}; font-weight: 700; color: #92400e;">{i_title}</span>
                </div>
                <div style="background-color: #ffffff; border-left: 4px solid #f59e0b; padding: 10px 12px; margin-bottom: 10px; border-radius: 0 4px 4px 0;">
                    <span style="font-size: {font_body}; font-weight: 700; color: #d97706; display: block; margin-bottom: 4px;">【實務觀察】</span>
                    <p style="color: #451a03; font-size: {font_body}; line-height: 1.5; margin: 0;">{i_obs}</p>
                </div>
                <div style="background-color: #ffffff; border-left: 4px solid #0d9488; padding: 10px 12px; border-radius: 0 4px 4px 0;">
                    <span style="font-size: {font_body}; font-weight: 700; color: #0f766e; display: block; margin-bottom: 4px;">【優化建議】</span>
                    <p style="color: #134e5e; font-size: {font_body}; line-height: 1.5; margin: 0;">{i_rec}</p>
                </div>
            </div>
            """
        
        issues_section = f"""
        <div style="font-size: 17px; font-weight: 700; color: #b45309; margin-bottom: 10px; border-bottom: 2px solid #d97706; padding-bottom: 4px;">
            實務觀察與優化建議
        </div>
        <div style="display: flex; gap: 12px; flex-wrap: wrap;">
            {issues_html}
        </div>
        """
    else:
        issues_section = f"""
        <div style="font-size: 17px; font-weight: 700; color: #0f766e; margin-bottom: 10px; border-bottom: 2px solid #0d9488; padding-bottom: 4px;">
            實務觀察與優化建議
        </div>
        <div style="background-color: #f0fdf4; border-left: 4px solid #16a34a; padding: 14px 20px; border-radius: 0 8px 8px 0; margin-top: 10px;">
            <span style="color: #16a34a; font-weight: 700; font-size: 16px;">🎉 本次專案深度查核結果：合規良好</span>
            <p style="color: #15803d; font-size: {font_body}; margin: 4px 0 0 0;">經實質抽核與程序檢視，受查標的於查核期間內未發現重大內部控制程序缺陷或異常發現，運作機制嚴謹順暢。</p>
        </div>
        """

    raw_html = f"""
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #1e293b; background: #ffffff; padding: 5px 10px;">
        <div class="gradient-header" style="background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); border-radius: 8px; padding: 12px 18px; margin-bottom: 16px;">
            <h1 style="color: #ffffff; font-size: 22px; font-weight: 700; margin: 0; letter-spacing: 0.5px;">{full_title}</h1>
        </div>
        
        <div style="font-size: 17px; font-weight: 700; color: #0f766e; margin-bottom: 10px; border-bottom: 2px solid #0d9488; padding-bottom: 4px;">
            執行查核事項
        </div>
        <div style="display: flex; gap: 10px; margin-bottom: {margin_b}; flex-wrap: wrap;">
            {highlights_html}
        </div>
        
        {issues_section}
    </div>
    """

    st.markdown(clean_html(raw_html), unsafe_allow_html=True)
    st.write("")

    pdf_html_source = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>{full_title}</title>
        <style>
            @page {{ size: A4 landscape; margin: 10mm; }}
            body, html {{ 
                font-family: 'Microsoft JhengHei', 'Segoe UI', sans-serif; 
                background-color: #ffffff; padding: 0; margin: 0; 
                -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;
            }}
            .gradient-header {{
                background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%) !important;
                -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important;
            }}
        </style>
    </head>
    <body>
        {raw_html}
    </body>
    </html>
    """

    col_btn, col_empty = st.columns([1, 2])
    with col_btn:
        try:
            from weasyprint import HTML
            pdf_data = HTML(string=clean_html(pdf_html_source)).write_pdf()
            st.download_button(
                label="📥 滿意成品！直接下載 PDF 實體報告",
                data=pdf_data, file_name=f"{full_title}.pdf", mime="application/pdf", type="primary", use_container_width=True
            )
        except Exception:
            st.download_button(
                label="📥 匯出成高畫質單頁報告檔 (.html)",
                data=clean_html(pdf_html_source).encode('utf-8'), file_name=f"{full_title}.html", mime="text/html", type="primary", use_container_width=True,
                help="點擊開啟後按 Ctrl+P，系統已自動強制保留漸層與背景，可直接完美轉 PDF 存檔。"
            )

    st.markdown("---")

    with st.expander("🛠️ 報告內容文字微調與手動覆核 (若需調整字眼請點此展開)"):
        st.caption("前端快取防護啟用中：在此處敲擊鍵盤修改任何文字皆 **0 消耗** 後台 API 請求配額，上方版面即刻達成零時差映射更新。")
        
        if highlights:
            st.markdown("**🔹 執行查核事項微調**")
            for idx, h in enumerate(highlights):
                c1, c2 = st.columns([1, 3])
                with c1:
                    h["title"] = st.text_input(f"標題 #{idx+1}", value=h.get("title", ""), key=f"edit_h_t_{idx}")
                with c2:
                    h["desc"] = st.text_input(f"內文遵循描述 #{idx+1}", value=h.get("desc", ""), key=f"edit_h_d_{idx}")
                
        if issues:
            st.markdown("<br>**🔸 實務觀察與優化建議微調**", unsafe_allow_html=True)
            for idx, item in enumerate(issues):
                st.markdown(f"**議題 #{idx+1}**")
                item["title"] = st.text_input(f"議題標題 #{idx+1}", value=item.get("title", ""), key=f"edit_i_t_{idx}")
                c_obs, c_rec = st.columns(2)
                with c_obs:
                    item["observation"] = st.text_area(f"實務觀察 #{idx+1}", value=item.get("observation", ""), height=70, key=f"edit_i_o_{idx}")
                with c_rec:
                    item["recommendation"] = st.text_area(f"優化建議 #{idx+1}", value=item.get("recommendation", ""), height=70, key=f"edit_i_r_{idx}")
        else:
            st.success("✨ 本次查核未產生任何待改善議題，無需進行下方文字微調！")

        st.session_state.ai_generated_data = data