import streamlit as st
import io
from datetime import date
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

st.set_page_config(page_title="肉品追溯系統", layout="centered")
st.title("🐖 豬肉追溯系統 - 4張照片完美整合版")

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

# =================【🥩 瓦城泰統肉品核心主資料庫 🥩】=================
PORK_DATABASE = {
    "ME-320039": {"品名": "1.2後腿絞肉206g", "供應商": "香里食品企業股份有限公司", "需要QR": False},
    "ME-320040": {"品名": "預設肉品項目206g", "供應商": "香里食品企業股份有限公司", "需要QR": False},
    "ME-320035": {"品名": "板油絞(4公斤)", "供應商": "弘飛", "需要QR": False},
    "ME-320024": {"品名": "買賣類-豬上排(單隻)", "供應商": "弘飛", "需要QR": False},
    "ME-320023": {"品名": "腩排丁", "供應商": "弘飛", "LowerQR": False},
    "ME-300064": {"品名": "買賣類-豬炒片", "供應商": "弘飛", "需要QR": False},
    "ME-300046": {"品名": "去皮五花肉片", "供應商": "弘飛", "需要QR": False},
    "ME-300062": {"品名": "帶骨里肌肉片", "供應商": "弘飛", "需要QR": False},
    "ME-300045": {"品名": "附皮肥肉丁", "供應商": "泰安", "需要QR": False},
    "ME-320004": {"品名": "帶皮五花肉片", "供應商": "泰安", "需要QR": False},
    "ME-320018": {"品名": "豬皮", "供應商": "香里", "需要QR": False},
    "ME-330036": {"品名": "中油角(1.5)", "供應商": "香里", "需要QR": False},
    "ME-320043": {"品名": "梅花肉丁", "供應商": "香里", "需要QR": False},
    "ME-330048": {"品名": "中油角", "供應商": "香里", "需要QR": False},
    "ME-230002": {"品名": "雞絞肉", "供應商": "超秦企業股份有限公司", "需要QR": True}
}

# =================【1. 前端網頁輸入介面】=================
st.subheader("📝 請填寫追溯表資料")

sku_list = list(PORK_DATABASE.keys())
part_no = st.selectbox("1. 請選取產品料號", sku_list)

auto_product_name = PORK_DATABASE[part_no]["品名"]
auto_supplier = PORK_DATABASE[part_no]["供應商"]
need_qr_flow = PORK_DATABASE[part_no].get("需要QR", False) # 💡 檢查這個產品要不要走 QR 流程

col1, col2 = st.columns(2)
with col1:
    product_name = st.text_input("2. 品名 (已由料號自動連動)", value=auto_product_name)
with col2:
    supplier = st.text_input("3. 供應商 (已由料號自動連動)", value=auto_supplier)
    in_date = st.date_input("4. 進貨日", value=date.today())

# 隱藏與留空變數設定
slaughter_date, slaughter_unit, meat_type, cut_date, drugs_check, bio_check = "", "", "", "", "", ""

st.write("---")
col3, col4 = st.columns(2)
with col3:
    make_date = st.date_input("11. 製造日期", value=date.today())
with col4:
    valid_date = st.date_input("12. 有效日期", value=date.today())

# =================【2. 📸 基礎照片上傳區】=================
st.write("---")
st.subheader("📸 基礎作業照片")

col_img1, col_img2 = st.columns(2)
with col_img1:
    st.markdown("**1. 產品包裝圖示照片**")
    uploaded_image_file = st.file_uploader("請選擇或拍攝肉品包裝照片：", type=["jpg", "jpeg", "png"], key="photo1")
with col_img2:
    st.markdown("**2. 進貨單據 / 銷貨單照片**")
    uploaded_receipt_file = st.file_uploader("請選擇或拍攝單據照片：", type=["jpg", "jpeg", "png"], key="photo2")

# =================【3. 📱 條件觸發：特定產品的 QR Code 截圖上傳區】=================
uploaded_qr_screenshot_1 = None
uploaded_qr_screenshot_2 = None

if need_qr_flow:
    st.write("---")
    st.warning("🔍 偵測到此特定產品需要進行 QR Code 查驗，請於下方補登相關截圖：")
    
    col_qr1, col_qr2 = st.columns(2)
    with col_qr1:
        st.markdown("**3. QR Code 頁面特定部分截圖**")
        uploaded_qr_screenshot_1 = st.file_uploader("請上傳第一頁特定範圍截圖：", type=["jpg", "jpeg", "png"], key="qr_snap1")
    with col_qr2:
        st.markdown("**4. 進入下個頁面之進階截圖**")
        uploaded_qr_screenshot_2 = st.file_uploader("請上傳次頁進階資訊截圖：", type=["jpg", "jpeg", "png"], key="qr_snap2")

# =================【4. 後端 Excel 4圖嵌入與一鍵下載邏輯】=================
st.write("---")

wb = Workbook()
ws = wb.active
ws.title = "加工肉品追蹤追溯表"

# 定義樣式
font_title = Font(name="微軟正黑體", size=16, bold=True)
font_section = Font(name="微軟正黑體", size=11, bold=True)
font_grid_header = Font(name="微軟正黑體", size=10, bold=True)
font_body = Font(name="微軟正黑體", size=10)
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
thin_border = Side(style='thin', color='000000')
border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

# 寫入大標題
ws.merge_cells('A1:L1')
ws['A1'] = "瓦城泰統集團\n加工肉品追蹤追溯表"
ws['A1'].font = font_title
ws['A1'].alignment = align_center
ws.row_dimensions.height = 50

# 區塊一：基本料號資訊
formatted_in_date = format_date_slash(in_date)
labels_v1 = [
    ('A2:B2', '料號', 'C2:D2', part_no), ('E2:F2', '供應商', 'G2:L2', supplier),
    ('A3:B3', '品名', 'C3:D3', product_name), ('E3:F3', '進貨日', 'G3:L3', formatted_in_date)
]
for lbl_rng, lbl_txt, val_rng, val_txt in labels_v1:
    ws.merge_cells(lbl_rng)
    ws.merge_cells(val_rng)
    ws[lbl_rng.split(':')] = lbl_txt  
    ws[val_rng.split(':')] = val_txt  
    style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 25

# 區塊二：產品包裝圖示
ws.merge_cells('A4:L4')
ws['A4'] = "產品包裝圖示"
style_range(ws, 'A4:L4', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 20

ws.merge_cells('A5:L5')
style_range(ws, 'A5:L5', border=border_all)
ws.row_dimensions.height = 180  

if uploaded_image_file is not None:
    pil_img = PILImage.open(uploaded_image_file)
    pil_img.thumbnail((400, 230))
    img_byte_arr = io.BytesIO()
    pil_img.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    ws.add_image(OpenpyxlImage(img_byte_arr), 'A5')
else:
    ws['A5'] = "（現場未上傳包裝照片）"
    ws['A5'].font = font_body
    ws['A5'].alignment = align_center

# 區塊三：原料肉資料 (留空)
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

formatted_make_date = format_date_slash(make_date)
formatted_valid_date = format_date_slash(valid_date)
labels_v4 = [
    ('A10:B10', '產品規格', 'C10:D10', ""), ('E10:F10', '製造日期', 'G10:H10', formatted_make_date), ('I10:J10', '有效日期', 'K10:L10', formatted_valid_date),
    ('A11:B11', '產品批號', 'C11:D11', ""), ('E11:F11', '生產數量', 'G11:H11', ""), ('I11:J11', '備註說明', 'K11:L11', "")
]
for lbl_rng, lbl_txt, val_rng, val_txt in labels_v4:
    ws.merge_cells(lbl_rng)
    ws.merge_cells(val_rng)
    ws[lbl_rng.split(':')] = lbl_txt  
    ws[val_rng.split(':')] = val_txt  
    style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 25

# 區塊五：相關進貨單據
ws.merge_cells('A12:L12')
ws['A12'] = "相關進貨單據 / 銷貨單收據證明"
style_range(ws, 'A12:L12', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 20

ws.merge_cells('A13:L13')
style_range(ws, 'A13:L13', border=border_all)
ws.row_dimensions.height = 220  

if uploaded_receipt_file is not None:
    pil_receipt = PILImage.open(uploaded_receipt_file)
    pil_receipt.thumbnail((450, 280))
    img_byte_arr2 = io.BytesIO()
    pil_receipt.save(img_byte_arr2, format='PNG')
    img_byte_arr2.seek(0)
    ws.add_image(OpenpyxlImage(img_byte_arr2), 'A13')
else:
    ws['A13'] = "（現場未上傳單據照片）"
    ws['A13'].font = font_body
    ws['A13'].alignment = align_center

# 🌟 區塊六：全新加開「QR Code 履歷查驗截圖區」（第 14、15 行大格子）
ws.merge_cells('A14:L14')
ws['A14'] = "特定品項 - QR Code 產銷履歷查驗與進階截圖證明"
style_range(ws, 'A14:L14', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 20

# 為了在同一行大格子橫向塞入「兩張截圖」，我們將 A15:F15 合併放第一張，G15:L15 合併放第二張
ws.merge_cells('A15:F15')
ws.merge_cells('G15:L15')
style_range(ws, 'A15:F15', border=border_all)
style_range(ws, 'G15:L15', border=border_all)
ws.row_dimensions.height = 240 # 給予足夠高度放置兩張截圖

# 塞入 QR 截圖 1
if uploaded_qr_screenshot_1 is not None:
    pil_qr1 = PILImage.open(uploaded_qr_screenshot_1)
    pil_qr1.thumbnail((260, 300))
    img_arr_qr1 = io.BytesIO()
    pil_qr1.save(img_arr_qr1, format='PNG')
    img_arr_qr1.seek(0)
    ws.add_image(OpenpyxlImage(img_arr_qr1), 'A15')
else:
    ws['A15'] = "（非特定品項，或未上傳 QR 畫面截圖）"
    ws['A15'].font = font_body
    ws['A15'].alignment = align_center

# 塞入 QR 截圖 2
if uploaded_qr_screenshot_2 is not None:
    pil_qr2 = PILImage.open(uploaded_qr_screenshot_2)
    pil_qr2.thumbnail((260, 300))
    img_arr_qr2 = io.BytesIO()
    pil_qr2.save(img_arr_qr2, format='PNG')
    img_arr_qr2.seek(0)
    ws.add_image(OpenpyxlImage(img_arr_qr2), 'G15')
else:
    ws['G15'] = "（非特定品項，或未上傳 下頁進階截圖）"
    ws['G15'].font = font_body
    ws['G15'].alignment = align_center

for col in ['A','B','C','D','E','F','G','H','I','J','K','L']:
    ws.column_dimensions[col].width = 11

excel_data = io.BytesIO()
wb.save(excel_data)
excel_data.seek(0)

st.download_button(
    label="📥 下載 Excel 報表",
    data=excel_data,
    file_name=f"加工肉品追蹤追溯表_{part_no}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
