import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

st.set_page_config(page_title="Delivery Formatter Pro", layout="wide")

def process_excel_to_buffer(uploaded_file):
    # 1. อ่านไฟล์ดิบ
    raw_df = pd.read_excel(uploaded_file, header=None)
    
    # 2. หาแถวที่มีคำว่า 'Item No.' (ป้องกัน Error กรณีเจอช่องว่าง)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in [str(v) for v in row.values])
    
    # --- แก้จุดที่ 1: บังคับเป็น String ก่อนเช็ก 'Unnamed' ---
    header_row_raw = raw_df.iloc[header_row_index].fillna("").astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    
    branch_to_code = {}
    for name, code in zip(header_row_raw, code_row_raw):
        # บังคับ name เป็น string ก่อนเช็ก 'Unnamed'
        s_name = str(name)
        if s_name and s_name != 'nan' and 'Unnamed' not in s_name and s_name != "":
            branch_to_code[s_name] = str(code) if pd.notna(code) else ""

    # 3. อ่านข้อมูลหลัก
    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    
    # --- แก้จุดที่ 2: กรองคอลัมน์ที่เป็น Unnamed หรือค่าว่างออก (ป้องกัน Float error) ---
    valid_columns = []
    for col in df.columns:
        s_col = str(col) # บังคับเป็น String ตรงนี้เลยค่ะแม่!
        if 'Unnamed' not in s_col and s_col != 'nan':
            valid_columns.append(col)
            
    df = df[valid_columns]
    df.columns = [str(c).strip() for c in df.columns]

    id_cols = ['Item No.', 'Description', 'UNIT']
    # Melt ข้อมูล
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    
    # กรองแถวที่ไม่มีจำนวนสั่งซื้อ
    final_list = df_melted.dropna(subset=['Item No.', 'Qty']).query('Qty > 0')
    
    # กรองชื่อสาขาที่เป็นขยะออกอีกรอบ
    final_list = final_list[~final_list['Branch'].astype(str).str.contains('Unnamed|#|nan', na=False)]

    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            s_branch = str(branch_name)
            store_code = branch_to_code.get(s_branch.strip(), "")
            
            # ล้างชื่อชีทให้ปลอดภัย
            clean_sheet_name = "".join([c for c in s_branch if c.isalnum() or c in ' -_'])[:30]
            if not clean_sheet_name: clean_sheet_name = f"Sheet_{hash(s_branch)}"
            
            items_df = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            items_df.to_excel(writer, sheet_name=clean_sheet_name, index=False, header=False, startrow=10)
            apply_styles_to_sheet(writer.sheets[clean_sheet_name], s_branch, store_code, current_date)

        # หน้า Summary
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
    # (โค้ดส่วนตกแต่งเหมือนเดิมเป๊ะค่ะแม่ ลูกไม่ตัดทอนจริตแน่นอน)
    font_name = 'Cordia New'
    f_title = Font(name=font_name, bold=True, size=20)
    f_header = Font(name=font_name, bold=True, size=14)
    f_data = Font(name=font_name, size=14)
    f_white = Font(name=font_name, bold=True, color="FFFFFF", size=14)
    fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    fill_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.merge_cells('A1:G1')
    ws['A1'] = "ใบสรุปรายการเบิกสินค้า" if is_summary else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
    ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_header
    ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_header
    ws['A7'] = f"Code: {store_code}"; ws['C7'] = f"Name: {branch_name}"
    ws['A7'].font = ws['C7'].font = f_header

    headers = ['No', 'Product Code', 'Product Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN']
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=10, column=i, value=h)
        cell.font, cell.fill, cell.border = f_white, fill_green, border
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7), 1):
        if str(ws.cell(row=row[0].row, column=3).value) == "Grand Total (ยอดรวมทั้งหมด)": continue
        for cell in row:
            cell.border = border; cell.font = f_data
            if r_idx % 2 == 0: cell.fill = fill_light
            if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

    ws.page_setup.paperSize = 9
    widths = {'A': 6, 'B': 16, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

# --- หน้าบ้าน ---
st.title("🚚 | POD BNN - Final Fix | 🚚")
uploaded_file = st.file_uploader("Upload Excel (จัดมาค่ะแม่ รอบนี้ไม่ Error แน่นอน)", type="xlsx")

if uploaded_file:
    with st.spinner('กำลังใช้จริตจัดการคอลัมน์เจ้าปัญหา...'):
        try:
            excel_bytes = process_excel_to_buffer(uploaded_file)
            st.success("💅🏻 กริบ! เล็บเจลยังต้องยอม รอบนี้ผ่านฉลุยค่ะแม่ 💅🏻")
            st.download_button(label="📥 โหลดไฟล์ที่นี่เลยงับ", data=excel_bytes, file_name=f"POD_Final_Fixed_{datetime.now().strftime('%H%M')}.xlsx")
        except Exception as e:
            st.error(f"อุ๊ย! แม่คะ พลังงานลึกลับบอกว่า: {e}")