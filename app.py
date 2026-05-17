import streamlit as st
import io
from datetime import date
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

st.set_page_config(page_title="肉品追溯系統", layout="centered")
st.title("🐖 豬肉追溯系統 - 雙照片升級版")

# =================【1. 前端網頁輸入介面】=================
st.subheader("📝 請填寫追溯表資料")
col1, col2 = st.columns(2)

with col1:
    part_no = st.selectbox("1. 料號", ["ME-320039", "ME-320040"])
    product_name = st.text_input("2. 品名", value="1.2後腿絞肉206g")

with col2:
    supplier = st.text_input("3. 供應商", value="香里食品企業股份有限公司")
    in_date = st.date_input("4. 進貨日", value=date(2025, 12, 26))

# 隱藏與留空變數設定（網頁不顯示，Excel 留空）
slaughter_date = ""  
slaughter_unit = ""  
meat_type = ""       
cut_date = ""        

drugs_check = "乙型受體素、鹽酸克倫特羅、抗生物質"
bio_check = "總生菌數"

# 加工日期
st.write("---")
col3, col4 = st.columns(2)
with col3:
    make_date = st.date_input("11. 製造日期", value=date(2025, 12, 20))
with col4:
    valid_date = st.date_input("12. 有效日期", value=date(2026, 12, 19))

# =================【2. 📸 產品包裝圖示上傳】=================
st.write("---")
st.subheader("📸 1. 產品包裝圖示照片")

upload_method_1 = st.radio("選擇第一張相片來源：", ["📷 用手機直接拍照", "📁 從相簿選取照片"], horizontal=True, key="photo1_method")
uploaded_image_file = None

if upload_method_1 == "📷 用手機直接拍照":
    uploaded_image_file = st.camera_input("請對準肉品標籤或外包裝拍照：", key="photo1_cam")
else:
    uploaded_image_file = st.file_uploader("請選擇肉品包裝照片：", type=["jpg", "jpeg", "png"], key="photo1_file")

# =================【3. 📄 銷貨單/進貨單據上傳】=================
st.write("---")
st.subheader("📄 2. 進貨單據 / 銷貨單收據照片")

upload_method_2 = st.radio("選擇第二張相片來源：", ["📷 用手機直接拍照", "📁 從相簿選取照片"], horizontal=True, key="photo2_method")
uploaded_receipt_file = None

if upload_method_2 == "📷 用手機直接拍照":
    uploaded_receipt_file = st.camera_input("請對準銷貨單/進貨收據拍照：", key="photo2_cam")
else:
    uploaded_receipt_file = st.file_uploader("請選擇單據照片：", type=["jpg", "jpeg", "png"], key="photo2_file")

# ===================================================
# 輔助函數：快速格式化儲存格區域
def style_range(ws, cell_range, font=None, alignment=None, fill=None, border=None):
    for row in ws[cell_range]:
        for cell in row:
            if font: cell.font = font
            if alignment: cell.alignment = alignment
            if fill: cell.fill = fill
            if border: cell.border = border

# =================【4. 後端 Excel 雙圖嵌入邏輯】=================
st.write("---")
if st.button("🚀 匯出完美還原 Excel 報表", use_container_width=True):
    
    # 建立活頁簿
    wb = Workbook()
    ws = wb.active
    ws.title = "加工肉品追蹤追溯表"

    # 定義樣式與字型
    font_title = Font(name="微軟正黑體", size=16, bold=True)
    font_section = Font(name="微軟正黑體", size=11, bold=True)
    font_grid_header = Font(name="微軟正黑體", size=10, bold=True)
    font_body = Font(name="微軟正黑體", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_border = Side(style='thin', color='000000')
    border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

    # 寫入大標題 (第一行)
    ws.merge_cells('A1:L1')
    ws['A1'] = "瓦城泰統集團\n加工肉品追蹤追溯表"
    ws['A1'].font = font_title
    ws['A1'].alignment = align_center
    ws.row_dimensions.height = 50

    # 區塊一：基本料號資訊表
    labels_v1 = [
        ('A2:B2', '料號', 'C2:D2', part_no),
        ('E2:F2', '供應商', 'G2:L2', supplier),
        ('A3:B3', '品名', 'C3:D3', product_name),
        ('E3:F3', '進貨日', 'G3:L3', str(in_date))
    ]
    for lbl_rng, lbl_txt, val_rng, val_txt in labels_v1:
        ws.merge_cells(lbl_rng)
        ws.merge_cells(val_rng)
        ws[lbl_rng.split(':')] = lbl_txt
        ws[val_rng.split(':')] = val_txt
        style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
        style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
    ws.row_dimensions.height = 25

    # 區塊二：產品包裝圖示欄位（抬頭）
    ws.merge_cells('A4:L4')
    ws['A4'] = "產品包裝圖示"
    style_range(ws, 'A4:L4', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions.height = 20

    # 【照片一】包裝圖示圖片置放儲存格（A5:L5 大格子）
    ws.merge_cells('A5:L5')
    style_range(ws, 'A5:L5', border=border_all)
    ws.row_dimensions.height = 180  

    if uploaded_image_file is not None:
        pil_img = PILImage.open(uploaded_image_file)
        pil_img.thumbnail((400, 230))
        img_byte_arr = io.BytesIO()
        pil_img.save(img_byte_arr, format='PNG')
        img_byte_arr.seek(0)
        xl_img = OpenpyxlImage(img_byte_arr)
        ws.add_image(xl_img, 'A5')
    else:
        ws['A5'] = "（現場未上傳/拍攝包裝照片）"
        ws['A5'].font = font_body
        ws['A5'].alignment = align_center

    # 區塊三：原料肉資料 (5~8項空字串)
    ws.merge_cells('A6:L6')
    ws['A6'] = "原料肉資料"
    style_range(ws, 'A6:L6', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)

    headers_meat = ['屠宰日期', '屠宰單位', '原料肉名稱', '原料肉分切日期', '動物用藥檢驗', '微生物檢驗']
    values_meat = [slaughter_date, slaughter_unit, meat_type, cut_date, drugs_check, bio_check]
    col_pairs = [('A7:B7','A8:B8'), ('C7:D7','C8:D8'), ('E7:F7','E8:F8'), ('G7:H7','G8:H8'), ('I7:J7','I8:J8'), ('K7:L7','K8:L8')]
    
    for i, (h_rng, v_rng) in enumerate(col_pairs):
        ws.merge_cells(h_rng)
        ws.merge_cells(v_rng)
        ws[h_rng.split(':')] = headers_meat[i]
        ws[v_rng.split(':')] = values_meat[i]
        style_range(ws, h_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
        style_range(ws, v_rng, font=font_body, alignment=align_center, border=border_all)

    # 區塊四：產品加工資料
    ws.merge_cells('A9:L9')
    ws['A9'] = "產品加工資料"
    style_range(ws, 'A9:L9', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)

    labels_v4 = [
        ('A10:B10', '產品規格', 'C10:D10', ""),
        ('E10:F10', '製造日期', 'G10:H10', str(make_date)),
        ('I10:J10', '有效日期', 'K10:L10', str(valid_date)),
        ('A11:B11', '產品批號', 'C11:D11', ""),
        ('E11:F11', '生產數量', 'G11:H11', ""),
        ('I11:J11', '備註說明', 'K11:L11', "")
    ]
    for lbl_rng, lbl_txt, val_rng, val_txt in labels_v4:
        ws.merge_cells(lbl_rng)
        ws.merge_cells(val_rng)
        ws[lbl_rng.split(':')] = lbl_txt
        ws[val_rng.split(':')] = val_txt
        style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
        style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
    ws.row_dimensions.height = 25
    ws.row_dimensions.height = 25

    # 🌟 區塊五：新增「銷貨單據/收據」欄位（最下方）
    ws.merge_cells('A12:L12')
    ws['A12'] = "相關進貨單據 / 銷貨單收據證明"
    style_range(ws, 'A12:L12', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions.height = 20

    # 【照片二】單據圖片置放儲存格（A13:L13 大格子）
    ws.merge_cells('A13:L13')
    style_range(ws, 'A13:L13', border=border_all)
    ws.row_dimensions.height = 220  # 收據通常比較長，格子拉稍微高一點

    if uploaded_receipt_file is not None:
        pil_receipt = PILImage.open(uploaded_receipt_file)
        # 銷貨單寬度可以拉寬一點，這裡設定最大寬度 450 像素，最大高度 280 像素
        pil_receipt.thumbnail((450, 280))
        img_byte_arr2 = io.BytesIO()
        pil_receipt.save(img_byte_arr2, format='PNG')
        img_byte_arr2.seek(0)
        xl_receipt = OpenpyxlImage(img_byte_arr2)
        ws.add_image(xl_receipt, 'A13')
    else:
        ws['A13'] = "（現場未上傳/拍攝單據照片）"
        ws['A13'].font = font_body
        ws['A13'].alignment = align_center

    for col in ['A','B','C','D','E','F','G','H','I','J','K','L']:
        ws.column_dimensions[col].width = 11

    # 轉換成 Streamlit 下載流
    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)

    st.download_button(
        label="📥 下載內嵌照片之完美排版 Excel 檔",
        data=excel_data,
        file_name=f"加工肉品追蹤追溯表_{part_no}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
