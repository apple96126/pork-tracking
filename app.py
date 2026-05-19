import streamlit as st
import io
from datetime import date
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from email.utils import encode_rfc2231

def send_email(to_email, subject, body, attachment=None):
    from_email = "你的gmail帳號@gmail.com"
    password = "你的應用程式密碼"

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email

    # 信件內容用 UTF-8
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 如果要附上 Excel 報表
    if attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.getvalue())
        encoders.encode_base64(part)

        # 🔹用 RFC2231 編碼檔名，避免 ASCII 錯誤
        filename = "肉品追溯表.xlsx"
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=encode_rfc2231(filename, "utf-8")
        )
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(from_email, password)
        server.sendmail(from_email, [to_email], msg.as_string())

st.set_page_config(page_title="肉品追溯系統", layout="centered")
st.title("🐖 豬肉追溯系統")

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
        "ME-330036": {"品名": "中油角(1.5)", "需要QR": False},
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
# 【1. 前端網頁輸入介面】
st.subheader("請填寫追溯表資料")

supplier_list = list(PORK_HIERARCHY.keys())
supplier = st.selectbox("1. 請先選取供應商", supplier_list)

available_skus = PORK_HIERARCHY[supplier]
sku_options = [f"{sku} - {info['品名']}" for sku, info in available_skus.items()]
selected_sku_string = st.selectbox("2. 請選取產品品項", sku_options)

selected_sku_code = selected_sku_string.split(" -")[0]
product_name = available_skus[selected_sku_code]["品名"]
need_qr_flow = available_skus[selected_sku_code].get("專屬與特定的QR流程", False) or available_skus[selected_sku_code].get("需要QR", False)

col1, col2 = st.columns(2)
with col1:
    st.info(f"自動對應料號:{selected_sku_code}")
with col2:
    in_date = st.date_input("4. 進貨日", value=date.today())

slaughter_date, slaughter_unit, meat_type, cut_date, drugs_check, bio_check = "", "", "", "", "", ""

st.write("---")
col3, col4 = st.columns(2)
with col3:
    make_date = st.date_input("11. 製造日期", value=date.today())
with col4:
    valid_date = st.date_input("12. 有效日期", value=date.today())

# 新增信箱輸入 + 驗證
email_address = st.text_input("5. 請輸入聯絡信箱")

email_valid = False
if email_address:
    if "@" not in email_address or "." not in email_address.split("@")[-1]:
        st.error("⚠️ 信箱格式不正確，請重新輸入 (必須包含 @ 與網域)")
    else:
        st.success("✅ 信箱格式正確")
        email_valid = True
else:
    st.warning("⚠️ 請輸入聯絡信箱後才能下載 Excel 報表")

# =================【2. 📸 照片上傳區】=================
st.write("---")
st.subheader("📸 進貨照片")

col_img1, col_img2 = st.columns(2)
with col_img1:
    st.markdown("**1. 產品包裝照片 (可上傳多張)**")
    uploaded_image_files = st.file_uploader("請上傳肉品包裝照片（可多選）：", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="photo1_multi")
with col_img2:
    st.markdown("**2. 進貨單據 / 銷貨單照片**")
    uploaded_receipt_file = st.file_uploader("請選擇或拍攝單據照片：", type=["jpg", "jpeg", "png"], key="photo2")

uploaded_qr_screenshot_1 = None
uploaded_qr_screenshot_2 = None

if need_qr_flow:
    st.write("---")
    st.warning("🔍 偵測到此特定產品需要進行 QR Code 查驗，請於下方補登相關截圖：")
    col_qr1, col_qr2 = st.columns(2)
    with col_qr1:
        st.markdown("**3. QR Code 屠宰日期**")
        uploaded_qr_screenshot_1 = st.file_uploader("請上傳第一頁特定範圍截圖：", type=["jpg", "jpeg", "png"], key="qr_snap1")
    with col_qr2:
        st.markdown("**4. 國產生鮮禽肉溯源平台**")
        uploaded_qr_screenshot_2 = st.file_uploader("請上傳次頁進階資訊截圖：", type=["jpg", "jpeg", "png"], key="qr_snap2")

# =================【3. 後端 Excel 多圖嵌入與一鍵下載邏輯】=================
st.write("---")

wb = Workbook()
ws = wb.active

# 🔹工作表名稱改成供應商+料號+品名
ws.title = f"{supplier}_{selected_sku_code}_{product_name}"


# 修改為 A4 直向列印設定
ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT   
ws.page_setup.paperSize = ws.PAPERSIZE_A4             
ws.sheet_properties.pageSetUpPr.fitToPage = True      
ws.page_setup.fitToWidth = 1                          
ws.page_setup.fitToHeight = 0                         

font_title = Font(name="標楷體", size=20, bold=True)
font_section = Font(name="標楷體", size=11, bold=True)
font_grid_header = Font(name="標楷體", size=10, bold=True)
font_body = Font(name="標楷體", size=10)
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
thin_border = Side(style='thin', color='000000')
border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

# 統一設定欄寬：直向 A4 為了塞入 12 欄，將欄寬調整為 9.5 較為緊湊剛好
for col in ['A','B','C','D','E','F','G','H','I','J','K','L']:
    ws.column_dimensions[col].width = 9.5

# 寫入大標題
ws.merge_cells('A1:L1')
ws['A1'] = "瓦城泰統集團\n加工肉品追蹤追溯表"
ws['A1'].font = font_title
ws['A1'].alignment = align_center
ws.row_dimensions[1].height = 45

# 區塊一：基本料號資訊
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

# ==========================================
# 🌟【雙軌分流】：區塊二與圖片並排邏輯
# ==========================================
ws.merge_cells('A4:L4')
ws.merge_cells('A5:L5')
style_range(ws, 'A5:L5', border=border_all)

if need_qr_flow:
    # ----------------------------------------------------
    # 【分流一：需要 QR Code 的排版程式碼】
    # ----------------------------------------------------
    ws['A4'] = "產品包裝與 QR 查驗圖示"
    style_range(ws, 'A4:L4', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 247.50  # 保持原設定高列高
    
    # 整合所有要並排的照片 (最多 3 張包裝照片 + 2 張 QR 截圖，總共最多 5 張)
    all_images_to_pack = []
    if uploaded_image_files:
        for f in uploaded_image_files[:3]: 
            all_images_to_pack.append(f)
    if uploaded_qr_screenshot_1 is not None:
        all_images_to_pack.append(uploaded_qr_screenshot_1)
    if uploaded_qr_screenshot_2 is not None:
        all_images_to_pack.append(uploaded_qr_screenshot_2)
        
    # 設定 6 個等比例橫並排的虛擬欄位起點
    target_columns = ['A5', 'C5', 'E5', 'G5', 'I5', 'K5']
    
    if all_images_to_pack:
        for idx, img_file in enumerate(all_images_to_pack[:6]):
            if idx < len(target_columns):
                pil_img = PILImage.open(img_file)
                pil_img.thumbnail((150, 310))  # 保持寬 150，高度適應 247.50 的安全尺寸
                
                img_stream = io.BytesIO()
                pil_img.save(img_stream, format='PNG')
                img_stream.seek(0)
                
                img_obj = OpenpyxlImage(img_stream)
                img_obj.anchor = target_columns[idx]
                ws.add_image(img_obj)
    else:
        ws['A5'] = "（現場未上傳包裝或查驗照片）"
        ws['A5'].font = font_body
        ws['A5'].alignment = align_center

else:
    # ----------------------------------------------------
    # 【分流二：沒有 QR Code 的排版程式碼】
    # ----------------------------------------------------
    ws['A4'] = "產品包裝圖示"
    style_range(ws, 'A4:L4', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 150.00  # 縮小沒有 QR Code 流程時的格子高度，維持排版緊湊美觀
    
    # 只抓取產品包裝照片 (最多 3 張)
    all_images_to_pack = []
    if uploaded_image_files:
        for f in uploaded_image_files[:3]: 
            all_images_to_pack.append(f)
            
    # 調配 3 張圖在 12 欄中的均勻虛擬起點 (每隔 4 欄均分放一張)
    target_columns_no_qr = ['A5', 'E5', 'I5']
    
    if all_images_to_pack:
        for idx, img_file in enumerate(all_images_to_pack[:3]):
            if idx < len(target_columns_no_qr):
                pil_img = PILImage.open(img_file)
                pil_img.thumbnail((150, 200))  # 調降高度，適配 150.00 的列高
                
                img_stream = io.BytesIO()
                pil_img.save(img_stream, format='PNG')
                img_stream.seek(0)
                
                img_obj = OpenpyxlImage(img_stream)
                img_obj.anchor = target_columns_no_qr[idx]
                ws.add_image(img_obj)
    else:
        ws['A5'] = "（現場未上傳產品包裝照片）"
        ws['A5'].font = font_body
        ws['A5'].alignment = align_center


# 區塊三：原料肉資料
ws.merge_cells('A6:L6')
ws['A6'] = "原料肉資料"
style_range(ws, 'A6:L6', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 22

headers_meat = ['屠宰日期', '屠宰單位', '原料肉名稱', '原料肉分切日期', '動物用藥檢驗', '微生物檢驗']
values_meat = [slaughter_date, slaughter_unit, meat_type, cut_date, drugs_check, bio_check]
col_pairs = [('A7:B7','A8:B8'), ('C7:D7','C8:D8'), ('E7:F7','E8:F8'), ('G7:H7','G8:H8'), ('I7:J7','I8:J8'), ('K7:L7','K8:L8')]
for i, (h_rng, v_rng) in enumerate(col_pairs):
    ws.merge_cells(h_rng)
    ws.merge_cells(v_rng)
    ws[h_rng.split(':')[0]] = headers_meat[i]  
    ws[v_rng.split(':')[0]] = values_meat[i]  
    style_range(ws, h_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, v_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 24
ws.row_dimensions.height = 24

# 區塊四：產品加工資料
ws.merge_cells('A9:L9')
ws['A9'] = "產品加工資料"
style_range(ws, 'A9:L9', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 22

formatted_make_date = format_date_slash(make_date)
formatted_valid_date = format_date_slash(valid_date)
labels_v4 = [
    ('A10:B10', '產品規格', 'C10:D10', ""), ('E10:F10', '製造日期', 'G10:H10', formatted_make_date), ('I10:J10', '有效日期', 'K10:L10', formatted_valid_date),
    ('A11:B11', '產品批號', 'C11:D11', ""), ('E11:F11', '生產數量', 'G11:H11', ""), ('I11:J11', '備註說明', 'K11:L11', "")
]
for lbl_rng, lbl_txt, val_rng, val_txt in labels_v4:
    ws.merge_cells(lbl_rng)
    ws.merge_cells(val_rng)
    ws[lbl_rng.split(':')[0]] = lbl_txt  
    ws[val_rng.split(':')[0]] = val_txt  
    style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
ws.row_dimensions.height = 24
ws.row_dimensions.height = 24

# 區塊五：相關進貨單據 (🌟 指定列高 355.90)
ws.merge_cells('A12:L12')
ws['A12'] = "進貨單據"
style_range(ws, 'A12:L12', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
ws.row_dimensions.height = 20

ws.merge_cells('A13:L13')
style_range(ws, 'A13:L13', border=border_all)
ws.row_dimensions[13].height = 355.90  # 修正：精確指定第 13 行的列高

if uploaded_receipt_file is not None:
    pil_receipt = PILImage.open(uploaded_receipt_file)
    pil_receipt.thumbnail((750, 450))
    
    img_byte_arr2 = io.BytesIO()
    pil_receipt.save(img_byte_arr2, format='PNG')
    img_byte_arr2.seek(0)
    
    receipt_obj = OpenpyxlImage(img_byte_arr2)
    receipt_obj.anchor = 'A13'
    ws.add_image(receipt_obj)
else:
    ws['A13'] = "（現場未上傳單據照片）"
    ws['A13'].font = font_body
    ws['A13'].alignment = align_center

# 匯出與下載處理
excel_data = io.BytesIO()
wb.save(excel_data)
excel_data.seek(0)

# 🔹下載檔案名稱改成供應商+料號+品名
download_filename = f"{supplier}_{selected_sku_code}_{product_name}.xlsx"

# 🔹只有在信箱驗證通過時才顯示下載按鈕
if email_valid:
    st.download_button(
        label="📥下載 Excel 報表",
        data=excel_data,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
)

 # 🔹寄送 Email 按鈕
    if st.button("寄送 Email"):
        body_text = f"""您好，這是您填寫的追溯表通知。
供應商: {supplier}
品名: {product_name}
進貨日: {in_date}
"""
        try:
            send_email(
                to_email=email_address,
                subject="肉品追溯表通知",
                body=body_text,
                attachment=excel_data  # 附上 Excel 報表
            )
            st.success("✅ 已寄送到您的信箱")
        except Exception as e:
            st.error(f"❌ 寄信失敗，錯誤訊息: {e}")
else:
    st.error("⚠️ 請確認信箱格式正確，才能下載或寄送 Excel 報表")
