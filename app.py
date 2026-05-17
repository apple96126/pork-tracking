import streamlit as st
import io
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

st.set_page_config(page_title="肉品追溯系統", layout="centered")
st.title("🐖 豬肉追溯系統")

# =================【前端網頁輸入介面】=================
st.subheader("📝 請填寫追溯表資料")
col1, col2 = st.columns(2)

with col1:
    part_no = st.selectbox("1. 料號", ["ME-320039"])
    product_name = st.text_input("2. 品名", value="1.2後腿絞肉206g")
    supplier = st.text_input("3. 供應商", value="香里食品企業股份有限公司")
    in_date = st.date_input("4. 進貨日", value=date(2025, 12, 26))

with col2:
    slaughter_date = st.text_input("5. 屠宰日期", value="12月18日")
    slaughter_unit = st.text_input("6. 屠宰單位", value="南投縣農產運銷股份有限公司")
    meat_type = st.text_input("7. 原料肉名稱", value="後腿肉")
    cut_date = st.text_input("8. 原料肉分切日期", value="12月19日")

# 檢驗與加工額外欄位
st.write("---")
drugs_check = st.text_input("9. 動物用藥檢驗項目", value="乙型受體素、鹽酸克倫特羅、抗生物質")
bio_check = st.text_input("10. 微生物檢驗", value="總生菌數")
make_date = st.date_input("11. 製造日期", value=date(2025, 12, 20))
valid_date = st.date_input("12. 有效日期", value=date(2026, 12, 19))

# =================【後端 Excel 完美刻版邏輯】=================
if st.button("🚀 匯出 Excel 報表", use_container_width=True):
    
    # 建立活頁簿
    wb = Workbook()
    ws = wb.active
    ws.title = "加工肉品追蹤追溯表"
    ws.views.sheetView[0].showGridLines = True # 強制顯示 Excel 預設網格線

    # ---- 1. 定義各式樣式與字型 ----
    font_title = Font(name="微軟正黑體", size=16, bold=True)
    font_section = Font(name="微軟正黑體", size=11, bold=True)
    font_grid_header = Font(name="微軟正黑體", size=10, bold=True)
    font_body = Font(name="微軟正黑體", size=10)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # 邊框設定：仿製圖片的黑色一格一格邊框
    thin_border = Side(style='thin', color='000000')
    border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

    # ---- 2. 寫入大標題 (第一行) ----
    ws.merge_cells('A1:L1')
    ws['A1'] = "瓦城泰統集團\n加工肉品追蹤追溯表"
    ws['A1'].font = font_title
    ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 50

    # ---- Helper 函數：快速格式化儲存格區域 ----
    def style_range(ws, cell_range, font=None, alignment=None, fill=None, border=None):
        for row in ws[cell_range]:
            for cell in row:
                if font: cell.font = font
                if alignment: cell.alignment = alignment
                if fill: cell.fill = fill
                if border: cell.border = border

    # ---- 3. 區塊一：基本料號資訊表 (橫跨 A~L 欄，共12欄寬) ----
    # 欄位合併配對規劃
    labels_v1 = [
        ('A2:B2', '料號', 'C2:D2', part_no),
        ('E2:F2', '供應商', 'G2:L2', supplier),
        ('A3:B3', '品名', 'C3:D3', product_name),
        ('E3:F3', '進貨日', 'G3:L3', str(in_date))
    ]
    for lbl_rng, lbl_txt, val_rng, val_txt in labels_v1:
        ws.merge_cells(lbl_rng)
        ws.merge_cells(val_rng)
        ws[lbl_rng.split(':')[0]] = lbl_txt
        ws[val_rng.split(':')[0]] = val_txt
        style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
        style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 25

    # ---- 4. 區塊二：產品包裝圖示欄位 ----
    ws.merge_cells('A4:L4')
    ws['A4'] = "產品包裝圖示"
    style_range(ws, 'A4:L4', font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions[4].height = 20

    ws.merge_cells('A5:L5')
    ws['A5'] = "（現場照片示意欄位）"
    style_range(ws, 'A5:L5', font=font_body, alignment=align_center, border=border_all)
    ws.row_dimensions[5].height = 100  # 留大格子未來可以放圖

    # ---- 5. 區塊三：原料肉資料標題 ----
    ws.merge_cells('A6:L6')
    ws['A6'] = "原料肉資料"
    style_range(ws, 'A6:L6', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions[6].height = 22

    # 原料肉明細表格 (橫向細分 6 區塊，每區塊佔2欄寬)
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
    ws.row_dimensions[7].height = 25
    ws.row_dimensions[8].height = 35

    # ---- 6. 區塊四：產品加工資料標題 ----
    ws.merge_cells('A9:L9')
    ws['A9'] = "產品加工資料"
    style_range(ws, 'A9:L9', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions[9].height = 22

    # 加工資料明細表格
    labels_v4 = [
            # ---- 6. 區塊四：產品加工資料標題 ----
    ws.merge_cells('A9:L9')
    ws['A9'] = "產品加工資料"
    style_range(ws, 'A9:L9', font=font_section, alignment=align_center, fill=fill_gray, border=border_all)
    ws.row_dimensions.height = 22

    # 📢 修改這裡：把產品規格、批號、生產數量的內容改成 "" (空字串)
    # 這樣 Excel 匯出時，格子就會維持一格一格的空白狀態，方便其他單位手寫或補登！
    labels_v4 = [
        ('A10:B10', '產品規格', 'C10:D10', ""),           # 留白給其他單位
        ('E10:F10', '製造日期', 'G10:H10', str(make_date)), # 您這邊自動帶出
        ('I10:J10', '有效日期', 'K10:L10', str(valid_date)),# 您這邊自動帶出
        ('A11:B11', '產品批號', 'C11:D11', ""),           # 留白給其他單位
        ('E11:F11', '生產數量', 'G11:H11', ""),           # 留白給其他單位
        ('I11:J11', '備註說明', 'K11:L11', "")            # 留白給其他單位
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

    ]
    for lbl_rng, lbl_txt, val_rng, val_txt in labels_v4:
        ws.merge_cells(lbl_rng)
        ws.merge_cells(val_rng)
        ws[lbl_rng.split(':')[0]] = lbl_txt
        ws[val_rng.split(':')[0]] = val_txt
        style_range(ws, lbl_rng, font=font_grid_header, alignment=align_center, fill=fill_gray, border=border_all)
        style_range(ws, val_rng, font=font_body, alignment=align_center, border=border_all)
    ws.row_dimensions[10].height = 25
    ws.row_dimensions[11].height = 25

    # 強制手動微調各欄寬，確保合併後文字不被遮擋
    for col in ['A','B','C','D','E','F','G','H','I','J','K','L']:
        ws.column_dimensions[col].width = 11

    # ---- 7. 轉換成 Streamlit 下載流 ----
    excel_data = io.BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)

    st.download_button(
        label="📥 下載「一模一樣」排版 Excel 檔",
        data=excel_data,
        file_name=f"加工肉品追蹤追溯表_{part_no}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.success("🎉 報表樣式渲染完成！請點擊上方按鈕下載查看。")
