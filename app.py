import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

def process_excel_to_buffer(uploaded_file):
    # 1. อ่านข้อมูลดิบและหาตำแหน่ง Header
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    
    # ค้นหาแถวที่มี 'Item No.' เพื่อเป็นจุดเริ่มของตาราง
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # ดึง Store Code (บรรทัดที่อยู่ใต้ชื่อสาขา)
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    # 2. เตรียม DataFrame สินค้า
    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT']
    # Melt ข้อมูลเพื่อให้ 1 แถว คือ 1 รายการสินค้าต่อ 1 สาขา
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    
    # กรองเอาเฉพาะที่มีจำนวนสั่งซื้อ > 0
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    # 3. สร้าง Excel
    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            
            # เตรียมข้อมูลที่จะเขียนลงตาราง (เริ่มแถว 11)
            items_df = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "",
                'BNN': ""
            })
            items_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            
            # --- ตั้งค่า Styles ---
            f_title = Font(name='Sarabun', bold=True, size=16)
            f_bold = Font(name='Sarabun', bold=True, size=10)
            f_norm = Font(name='Sarabun', size=10)
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)
            fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- ส่วนหัวกระดาษ (Header 1-8) ---
            ws.merge_cells('A1:G1')
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"
            ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')

            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_bold
            ws['A3'] = "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ"; ws['A3'].font = f_norm
            
            ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_bold
            ws['G4'] = f"Delivery: {current_date}"; ws['G4'].alignment = Alignment(horizontal='right'); ws['G4'].font = f_bold

            ws['A6'] = f"Customer Name: {customer_name}"; ws['A6'].font = f_bold
            ws['A7'] = f"Store Code: {store_code}"; ws['A7'].font = f_bold
            ws['C7'] = f"Store Name: {branch_name}"; ws['C7'].font = f_bold

            # --- จัดการหัวตาราง (Rows 9-10) ---
            # Merge แนวนอนสำหรับ Qty
            ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=7)
            qty_cell = ws.cell(row=9, column=5, value="Qty")
            qty_cell.font = f_white; qty_cell.fill = fill_green; qty_cell.border = border_thin
            qty_cell.alignment = Alignment(horizontal='center', vertical='center')

            # หัวตารางหลักอื่นๆ
            main_headers = ['No', 'Product Code', 'Product Name', 'Unit']
            for i, h in enumerate(main_headers, 1):
                ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                c = ws.cell(row=9, column=i, value=h)
                c.font = f_white; c.fill = fill_green; c.border = border_thin
                c.alignment = Alignment(horizontal='center', vertical='center')

            # หัวตารางย่อย (ORDER, MBL, BNN)
            sub_headers = ['ORDER', 'MBL', 'BNN']
            for i, h in enumerate(sub_headers, 5):
                c = ws.cell(row=10, column=i, value=h)
                c.font = f_white; c.fill = fill_green; c.border = border_thin
                c.alignment = Alignment(horizontal='center')

            # --- ตีเส้นขอบและจัดฟอนต์ข้อมูลสินค้า ---
            for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
                for cell in row:
                    cell.border = border_thin
                    cell.font = f_norm
                    if cell.column in [1, 5, 6, 7]: # จัดกึ่งกลางเลขที่และจำนวน
                        cell.alignment = Alignment(horizontal='center')

            # --- ส่วนท้าย (Signatures) ---
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า").font = f_norm
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ").font = f_norm

            # --- ตั้งค่าหน้ากระดาษ A4 ---
            ws.page_setup.paperSize = 9 # A4
            ws.page_setup.orientation = 'portrait'
            ws.sheet_view.view = "pageLayout" # มุมมองแบบเห็นขอบกระดาษ
            ws.page_setup.fitToWidth = 1 # บีบให้กว้างพอดี 1 หน้า
            
            # ปรับความกว้างคอลัมน์ให้สมดุล
            widths = {'A': 6, 'B': 15, 'C': 35, 'D': 10, 'E': 10, 'F': 10, 'G': 10}
            for col, w in widths.items():
                ws.column_dimensions[col].width = w

    return output.getvalue()

# Streamlit UI
st.title("🚚 Delivery Formatter A4 (Final Version)")
st.write("อัปโหลดไฟล์ Excel เพื่อจัดรูปแบบใบส่งสินค้าให้ถูกต้องตามมาตรฐาน A4")

file = st.file_uploader("Upload Excel File", type="xlsx")

if file:
    try:
        excel_data = process_excel_to_buffer(file)
        st.success("จัดฟอร์แมตเสร็จสมบูรณ์! ข้อมูลครบถ้วนและพร้อมพิมพ์")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ใบส่งสินค้า (A4)",
            data=excel_data,
            file_name=f"Delivery_A4_{datetime.now().strftime('%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")