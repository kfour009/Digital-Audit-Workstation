import streamlit as st
import os
import shutil
import time
from PIL import Image
import pythoncom
import win32com.client as win32

# ==========================================
# 轉換核心函數
# ==========================================
def convert_word(input_path, output_path, log):
    pythoncom.CoInitialize()
    word = win32.DispatchEx('Word.Application')
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, 17) # 17 = PDF
        doc.Close(False)
        log.text(f"✅ Word 轉換成功: {os.path.basename(input_path)}")
    except Exception as e:
        log.error(f"❌ Word 轉換失敗: {str(e)}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

def convert_excel(input_path, output_path, log):
    pythoncom.CoInitialize()
    excel = win32.DispatchEx('Excel.Application')
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(input_path)
        for ws in wb.Worksheets:
            # 稽核底稿排版引擎 (直式)
            ws.UsedRange.WrapText = True           # 開啟自動換行
            ws.Columns.AutoFit()                   # 自動調整欄寬
            ws.PageSetup.PrintTitleRows = "$1:$1"  # 鎖定第一列為標題
            ws.PageSetup.Orientation = 1           # 1 = 直向 (Portrait)
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1        # 強制縮放為一頁寬
            ws.PageSetup.FitToPagesTall = False  # 💡 關鍵修正：將 0 改成大寫開頭的 False
            
            # 邊界設定 (18點)
            margin = 18
            ws.PageSetup.LeftMargin = margin
            ws.PageSetup.RightMargin = margin
            ws.PageSetup.TopMargin = margin
            ws.PageSetup.BottomMargin = margin
            
        wb.ExportAsFixedFormat(0, output_path)
        wb.Close(False)
        log.text(f"✅ Excel 轉換成功 (直式自動分頁): {os.path.basename(input_path)}")
    except Exception as e:
        log.error(f"❌ Excel 轉換失敗: {str(e)}")
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

def convert_image(input_path, output_path, log):
    try:
        img = Image.open(input_path)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img.save(output_path, "PDF", resolution=100.0)
        log.text(f"✅ 圖片轉換成功: {os.path.basename(input_path)}")
    except Exception as e:
        log.error(f"❌ 圖片轉換失敗: {str(e)}")

# ==========================================
# 網頁介面
# ==========================================
st.title("📄 萬能轉 PDF (直式單頁寬)")
st.write("支援 Word、Excel 與圖片轉為 PDF。**內建直式底稿引擎**：Excel 自動換行、鎖定標題列，寬度縮放一頁，高度自動分頁。")

supported_exts = ['doc', 'docx', 'xls', 'xlsx', 'csv', 'jpg', 'jpeg', 'png', 'bmp', 'webp']
files = st.file_uploader("📂 拖曳或點擊此處上傳檔案", type=supported_exts, accept_multiple_files=True)

if files and st.button("🚀 開始直式轉換", type="primary"):
    task_id = time.strftime("%Y%m%d_%H%M%S")
    WORK_DIR = os.path.abspath(f"temp_conv_v_{task_id}")
    OUT_DIR = os.path.abspath(f"out_conv_v_{task_id}")
    os.makedirs(WORK_DIR, exist_ok=True)
    os.makedirs(OUT_DIR, exist_ok=True)
    
    log = st.container()
    progress_bar = st.progress(0)
    total_files = len(files)
    
    for idx, f in enumerate(files):
        ext = os.path.splitext(f.name)[1].lower()
        input_path = os.path.join(WORK_DIR, f.name)
        output_path = os.path.join(OUT_DIR, f"{os.path.splitext(f.name)[0]}.pdf")
        
        with open(input_path, "wb") as tmp: tmp.write(f.getbuffer())
            
        if ext in ['.doc', '.docx']: convert_word(input_path, output_path, log)
        elif ext in ['.xls', '.xlsx', '.csv']: convert_excel(input_path, output_path, log)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.webp']: convert_image(input_path, output_path, log)
            
        progress_bar.progress((idx + 1) / total_files)

    st.success("🎉 直式轉換完畢！")
    zip_name = f"PDF_Portrait_{task_id}"
    shutil.make_archive(zip_name, 'zip', OUT_DIR)
    
    with open(f"{zip_name}.zip", "rb") as z:
        st.download_button("📥 下載直式 PDF 結果 (ZIP)", z, file_name="直式轉換結果.zip", type="primary")
        
    shutil.rmtree(WORK_DIR, ignore_errors=True)
    shutil.rmtree(OUT_DIR, ignore_errors=True)