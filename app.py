import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

st.set_page_config(page_title="Delivery Formatter Pro", layout="wide")

def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in [str(v) for v in row.values])
    
    header_row_raw = raw_df.iloc[header_row_index].fillna("").astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    
    branch_to_code = {}
    for name, code in zip(header_row_raw, code_row_raw):
        s_name = str(name)
        if s_name and s_name != 'nan' and 'Unnamed' not in s_name and s_name != "":
            branch_to_code[s_name] = str(code) if pd.notna(code) else ""

    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    
    valid_columns = []
    for col in df.columns:
        s_col = str(col)
        if 'Unnamed' not in s_col and s_col != 'nan':
            valid_columns.append(col)
    df = df[valid_columns]
    df.columns = [str(c).strip() for c in df.columns]

    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty']).query('Qty > 0')
    final_list = final_list[~final_list['Branch'].astype(str).str.contains('Unnamed|#|nan', na=False)]

    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            s_branch = str(branch_name)
            store_code = branch_to_code.get(s_branch.strip(), "")
            clean_sheet_name = "".join([c for c in s_branch if c.isalnum() or c in ' -_'])[:30]
            
            items_df = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            items_df.to_excel(writer, sheet_name=clean_sheet_name, index=False, header=False, startrow=10)
            apply_styles_to_sheet(writer.sheets[clean_sheet_name], s_branch, store_code, current_date, is_summary=False)

        summary_all = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary_all.insert(0, 'No', range(1, len(summary_all) + 1))
        summary_all['MBL'] = ""; summary_all['BNN'] = ""
        
        summary_all.to_excel(writer, sheet_name="Summary_All", index=False, header=False, startrow=10)
        ws_sum = writer.sheets["Summary_All"]
        
        last_row = 10 + len(summary_all) + 1
        ws_sum.cell(row=last_row, column=3, value="Grand Total (ยอดรวมทั้งหมด)").font = Font(name='Cordia New', bold=True, size=14)
        qty_total = ws_sum.cell(row=last_row, column=5, value=summary_all['Qty'].sum())
        qty_total.font = Font(name='Cordia New', bold=True, size=14)
        qty_total.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        qty_total.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))

        apply_styles_to_sheet(ws_sum, "สรุปยอดรวมทุกรายการ", "ALL", current_date, is_summary=True)

    return output.getvalue()

def apply_styles_to_sheet(ws, branch_name, store_code, current_date, is_summary=False):
    font_name = 'Cordia New'
    f_title = Font(name=font_name, bold=True, size=20)
    f_header = Font(name=font_name, bold=True, size=14)
    f_data = Font(name=font_name, size=14)
    
    # --- ปรับสีให้เหมาะกับการปริ้นท์ขาวดำ ---
    f_black_bold = Font(name=font_name, bold=True, color="000000", size=14) # หัวตารางเป็นตัวหนาสีดำ
    fill_light_green = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid") # เขียวอ่อน (พาสเทล)
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # ขาวล้วน
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # หัวที่อยู่บริษัท
    ws.merge_cells('A1:G1')
    ws['A1'] = "ใบสรุปรายการเบิกสินค้า" if is_summary else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
    ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_header
    ws['A3'] = "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540"; ws['A3'].font = f_data
    ws['A4'] = "โทร. 02-337-1200 แฟกซ์. 02-337-1201"; ws['A4'].font = f_data
    
    ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_header
    ws['G4'] = f"Delivery Date: {current_date}"; ws['G4'].alignment = Alignment(horizontal='right'); ws['G4'].font = f_header
    
    ws['A7'] = f"Code: {store_code}"; ws['C7'] = f"Name: {branch_name}"
    ws['A7'].font = ws['C7'].font = f_header

    # หัวตาราง
    ws.merge_cells('E9:G9')
    ws['E9'] = "Total Qty" if is_summary else "Qty"
    ws['E9'].font, ws['E9'].fill, ws['E9'].border = f_black_bold, fill_light_green, border
    ws['E9'].alignment = Alignment(horizontal='center')
    
    headers = ['No', 'Product Code', 'Product Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN']
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=10, column=i, value=h)
        cell.font, cell.fill, cell.border = f_black_bold, fill_light_green, border
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # เนื้อหาข้างใน (เป็นขาวล้วนทั้งหมด)
    for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
        if str(ws.cell(row=row[0].row, column=3).value) == "Grand Total (ยอดรวมทั้งหมด)": continue
        for cell in row:
            cell.border = border
            cell.font = f_data
            cell.fill = fill_white # ขาวล้วน
            if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

    # ท้ายกระดาษ
    if not is_summary:
        curr_row = ws.max_row + 2
        labels = ["ผู้รับสินค้า:", "ผู้ส่งสินค้า:", "ทะเบียนรถ:", "คลังสินค้า:"]
        for i, label in enumerate(labels):
            ws.cell(row=curr_row + i, column=1, value=f"{label} .......................................................").font = f_header
        
        # ตารางตะกร้า
        s_row = curr_row
        for i, h in enumerate(["", "MBL", "BNN"]):
            c = ws.cell(row=s_row, column=5+i, value=h)
            c.font, c.fill, c.border = f_black_bold, fill_light_green, border
            c.alignment = Alignment(horizontal='center')
        for i, label in enumerate(["ตะกร้าใหญ่", "ตะกร้าเล็ก"], 1):
            ws.cell(row=s_row + i, column=5, value=label).font = f_header
            for col_idx in range(5, 8):
                ws.cell(row=s_row + i, column=col_idx).border = border

    ws.page_setup.paperSize = 9
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    widths = {'A': 6, 'B': 16, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

# --- Streamlit UI ---
st.title("🚚 | POD BNN - Print Friendly Edition | 🚚")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    with st.spinner('กำลังปรับสีให้เหมาะกับการปริ้นท์ขาวดำนะคะแม่...'):
        try:
            excel_bytes = process_excel_to_buffer(file)
            st.success("💅🏻 กริบ! สีอ่อนโยน ปริ้นท์ออกมาดูแพงแน่นอนค่ะ 💅🏻")
            st.download_button(label="📥 โหลดไฟล์เวอร์ชันปริ้นท์ตรงนี้ค่ะ", data=excel_bytes, file_name=f"POD_PrintFriendly_{datetime.now().strftime('%H%M')}.xlsx")
        except Exception as e:
            st.error(f"อุ๊ย! ผิดพลาดนิดหน่อยค่ะแม่: {e}")