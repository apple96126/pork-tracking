import streamlit as st
import io
from datetime import date
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

st.set_page_config(page_title="肉品追溯系統", layout="centered")
st.title("🐖 豬肉追溯系統 - 高解析照片自動拉開版")

# 輔助函數：快速格式化儲存格區域
def style_range(ws, cell_range, font=None, alignment=None, fill=None, border=None):
    for row in ws[cell_range]:
        for cell in row:
            if font: cell.font = font
            if alignment: cell.alignment = alignment
            if fill: cell.fill = fill
            if border: cell.border = border

# 輔助函數：將日期格式化為 YYYY/M/D 斜線格式
def format_date_slash(dt):
    if isinstance(dt, date):
        return f"{dt.year}/{dt.month}/{dt.day}"
    return str(dt)

# =================【🥩 依廠商分類的核心主資料庫 🥩】=================
PORK_HIERARCHY = {
    "香里食品企業股份有限公司": {
        "ME-320039": {"品名": "1.2後腿絞肉206g", "需要QR": True},
        "ME-320040": {"品名": "預設肉品項目206g", "需要QR": True},
        "ME-320018": {"品名": "豬皮", "需要QR": False},
        "ME-330036": {"品名": "中油角(1.5)", "幕後QR": False},
        "ME-320043": {"品名": "梅花肉丁", "需要QR": False},
        "ME-330048": {"品名": "中油角", "需要QR": False}
    },
    "弘飛": {
        "ME-320035": {"品名": "板油絞(4公斤)", "需要QR": False},
        "ME-320024": {"品名": "買賣類-豬上排(單隻)", "需要QR": False},
        "ME-320023": {"品名": "腩排丁", "需要QR": False},
        "ME-300064": {"品名": "買賣類-豬炒片", "需要QR": False},
        "ME-300046": {"品名": "去皮五花肉片", "需要QR": False},
        "ME-300062": {"品名": "帶骨里肌肉片", "需要QR": False}
    },
    "泰安": {
        "ME-300045": {"品名": "附皮肥肉丁", "需要QR": False},
        "ME-320004": {"品名": "帶皮五花肉片", "需要QR": False}
    },
    "超秦企業股份有限公司": {
        "ME-230002": {"品名": "雞絞肉", "需要QR": True}
    }
}

# =================【1. 前端網頁輸入介面】=================
st.subheader("📝 請填寫追溯表資料")

supplier_list = list(PORK_HIERARCHY.keys())
supplier = st.selectbox("1. 請先選取供應商", supplier_list)

available_skus = PORK_HIERARCHY[supplier]
sku_options = [f"{sku} - {info['品名']}" for sku, info in available_skus.items()]
selected_sku_string = st.selectbox("2. 請選取產品品項", sku_options)

selected_sku_code = selected_sku_string.split(" - ")[0]
product_name = available_skus[selected_sku_code]["品名"]
need_qr_flow = available_skus[selected_sku_code].get("需要QR", False)

col1, col2 = st.columns(2)
with col1:
    st.info(f"📦 自動對應料號：{selected_sku_code}")
with col2:
    in_date = st.date_input("4. 進貨日", value=date.today())

slaughter_date, slaughter_unit, meat_type, cut_date, drugs_check, bio_check = "", "", "", "", "", ""

st.write("---")
col3, col4 = st.columns(2)
with col3:
    make_date = st.date_input("11. 製造日期", value=date.today())
with col4:
    valid_date = st.date_input("12. 有效日期", value=date.today())

# =================【2. 📸 照片上傳區】=================
st.write("---")
st.subheader("📸 基礎作業照片")

col_img1, col_img2 = st.columns(2)
with col_img1:
    st.markdown("**1. 產品包裝圖示照片 (可多選)**")
    uploaded_image_files = st.file_uploader("請上傳肉品包裝照片（可多選）：", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="photo1_multi")
with col_img2:
    st.markdown("**2. 進貨單據 / 銷貨單照片**")
    uploaded_receipt_file = st.file_uploader("請選擇或拍攝單據照片：", type=["jpg", "jpeg", "png"], key="photo2")

uploaded_qr_screenshot_1 = None
uploaded_qr_screenshot_2 = None

if need_qr_flow:
    st.write("---")
    st.warning("🔍 偵測到此特定產品需要進行 QR Code 查驗，請於下方補登相關截圖（匯出後將以高解析大圖垂直排列於 Excel 下方）：")
    col_qr1, col_qr2 = st.columns(2)
    with col_qr1:
        st.markdown("**3. QR Code 頁面特定部分截圖**")
        uploaded_qr_screenshot_1 = st.file_uploader("請上傳第一頁特定範圍截圖：", type=["jpg", "jpeg", "png"], key="qr_snap1")
    with col_qr2:
        st.markdown("**4. 進入下個頁面之進階截圖**")
        uploaded_qr_screenshot_2 = st.file_uploader("請上傳次頁進階資訊截圖：", type=["jpg", "jpeg", "png"], key="qr_snap2")

# =================【3. 後端 Excel 高解析縱向自動拉開邏輯】=================
st.write("---")

wb = Workbook()
ws = wb.active
ws.title = "加工肉品追蹤追溯表"

font_title = Font(name="微軟正黑體", size=16, bold=True)
font_section = Font(name="微軟正黑體", size=11, bold=True)
font_grid_header = Font(name="微軟正黑體", size=10, bold=True)
font_body = Font(name="微軟正黑體", size=10)
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
thin_border = Side(style='thin', color='000000')
border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

# 寫入大標題 (第 1 行)
ws.merge_cells('A1:L1')
ws['A1'] = "瓦城泰統集團\n加工肉品追蹤追溯表"
ws['A1'].font = font_title
ws['A1'].alignment = align_center
ws.row_dimensions.height = 45

# 區塊一：基本料號資訊 (第 2, 3 行)
formatted_in_date = format_date_slash(in_date)
labels_v1 = [
    ('A2:B2', '料號', 'C2:D2', selected_sku_code), ('E2:F2', '供應商', 'G2:L2', supplier),
    ('A3:B3', '品名', 'C3:D3', product_name), ('E3:F3', '進貨日', 'G3:L3', formatted_in_date)
]
for lbl_rng, lbl_txt, val_rng, val_txt in labels_v1:
    ws.merge_cells(lbl_rng)
    ws.merge_cells(val_rng)
    ws[lbl_rng.split(':')[0]] = lbl_txt  
    ws[val_rng.split(':')[0]] = val_txt  
    style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 24
ws.row_dimensions.height = 24

# 區塊二：原料肉資料 (第 4, 5 行)
ws.merge_cells('A4:L4')
ws['A4'] = "原料肉資料"
style_range(ws, 'A4:L4', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 22

headers_meat = ['屠宰日期', '屠宰單位', '原料肉名稱', '原料肉分切日期', '動物用藥檢驗', '微生物檢驗']
values_meat = [slaughter_date, slaughter_unit, meat_type, cut_date, drugs_check, bio_check]
col_pairs = [('A5:B5','A6:B6'), ('C5:D5','C6:D6'), ('E5:F5','E6:F6'), ('G5:H5','G6:H6'), ('I5:J5','I6:J6'), ('K5:L5','K6:L6')]
for i, (h_rng, v_rng) in enumerate(col_pairs):
    ws.merge_cells(h_rng)
    ws.merge_cells(v_rng)
    ws[h_rng.split(':')] = headers_meat[i]  
    ws[v_rng.split(':')] = values_meat[i]  
    style_range(ws, h_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, v_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 24
ws.row_dimensions.height = 24

# 區塊三：產品加工資料 (第 7, 8 行)
ws.merge_cells('A7:L7')
ws['A7'] = "產品加工資料"
style_range(ws, 'A7:L7', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 22

formatted_make_date = format_date_slash(make_date)
formatted_valid_date = format_date_slash(valid_date)
labels_v4 = [
    ('A8:B8', '產品規格', 'C8:D8', ""), ('E8:F8', '製造日期', 'G8:H8', formatted_make_date), ('I8:J8', '有效日期', 'K8:L8', formatted_valid_date),
    ('A9:B9', '產品批號', 'C9:D9', ""), ('E9:F9', '生產數量', 'G9:H9', ""), ('I9:J9', '備註說明', 'K9:L9', "")
]
for lbl_rng, lbl_txt, val_rng, val_txt in labels_v4:
    ws.merge_cells(lbl_rng)
    ws.merge_cells(val_rng)
    ws[lbl_rng.split(':')] = lbl_txt  
    ws[val_rng.split(':')] = val_txt  
    style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 24
ws.row_dimensions.height = 24

# 🌟 核心改版：建立一個「高解析動態縱向塞圖」的通用函數
current_row = 10  # 紀錄目前 Excel 已經寫到第幾行

def insert_large_photo_block(ws, title_text, file_objects, row_start):
    """橫跨 A~L 欄建立大型圖示區塊，並自動依圖片比例拉開行高"""
    # 1. 寫入該照片的區塊小標題
    ws.merge_cells(f'A{row_start}:L{row_start}')
    ws[f'A{row_start}'] = title_text
    style_range(ws, f'A{row_start}:L{row_start}', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions[row_start].height = 22
    
    img_row = row_start + 1
    
    if file_objects:
        # 如果是多張（清單），就跑巡迴往下蓋樓
        files = file_objects if isinstance(file_objects, list) else [file_objects]
        for f in files:
            if f is not None:
                # 合併這一列的 A~L 欄作為放置大圖的底盤
                ws.merge_cells(f'A{img_row}:L{img_row}')
                style_range(ws, f'A{img_row}:L{img_row}', border=border_all)
                
                # 讀取並放大照片：寬度加寬至 500 像素，維持極高清晰度
                pil_img = PILImage.open(f)
                w, h = pil_img.size
                scale_ratio = 500 / float(w)
                new_h = int(float(h) * scale_ratio)
                
                # 🌟 核心防呆：限制最高 450 像素避免單張圖過長，並將 PIL 圖片轉換為 Excel 可以讀的格式
                pil_img.thumbnail((500, 450))
                final_w, final_h = pil_img.size
                
                img_stream = io.BytesIO()
                pil_img.save(img_stream, format='PNG')
                img_stream.seek(0)
                
                # 🌟 神奇的自動拉開行高：將圖片的像素高度，精準換算成 Excel 的 Row Height 點數
                excel_row_height = int(final_h * 0.75) + 15
                ws.row_dimensions[img_row].height = excel_row_height
                
                # 將大圖塞入 A 欄起點
                ws.add_image(OpenpyxlImage(img_stream), f'A{img_row}')
                img_row += 1
    else:
        # 未上傳時建立一個小小的留白格子
        ws.merge_cells(f'A{img_row}:L{img_row}')
        ws[f'A{img_row}'] = "（現場同仁未上傳/拍攝此項目照片證明）"
        ws[f'A{img_row}'].font = font_body
        ws[f'A{img_row}'].alignment = align_center
        style_range(ws, f'A{img_row}:L{img_row}', border=border_all)
        ws.row_dimensions[img_row].height = 30
        img_row += 1
        
    return img_row

# 2. 依序動態向下建立 4 個高解析大圖格子
current_row = insert_large_photo_block(ws, "📷 項目一：產品包裝圖示證明", uploaded_image_files, current_row)
current_row = insert_large_photo_block(ws, "📄 項目二：相關進貨單據 / 銷貨單收據證明", uploaded_receipt_file, current_row)

if need_qr_flow:
    current_row = insert_large_photo_block(ws, "🔍 特定品項 - 項目三：QR Code 產銷履歷查驗特定範圍截圖", uploaded_qr_screenshot_1, current_row)
    current_row = insert_large_photo_block(ws, "📱 特定品項 - 項目四：進入下個頁面之進階截圖證明", uploaded_qr_screenshot_2, current_row)

# 調整固定欄寬
for col in ['A','B','C','D','E','F','G','H','I','J','K','L']:
    ws.column_dimensions[col].width = 11

excel_data = io.BytesIO()
wb.save(excel_data)
excel_data.seek(0)

st.download_button(
    label="📥 下載 Excel 報表",
    data=excel_data,
    file_name=f"加工肉品追蹤追溯表_{selected_sku_code}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
