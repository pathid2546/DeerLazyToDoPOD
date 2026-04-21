import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            
            data_to_write = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            data_to_write.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            
            # --- Styles ---
            f_title = Font(name='Sarabun', bold=True, size=16)
            f_bold = Font(name='Sarabun', bold=True, size=10)
            f_norm = Font(name='Sarabun', size=10)
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)
            fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- Header Drawing (เหมือนเดิม) ---
            ws.merge_cells('A1:G1')
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"; ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_bold
            ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_bold
            ws['A6'] = f"Customer Name: {customer_name}"; ws['A6'].font = f_bold
            ws['A7'] = f"Store Code: {store_code}"; ws['A7'].font = f_bold
            ws['C7'] = f"Store Name: {branch_name}"; ws['C7'].font = f_bold

            # --- Merge หัวตาราง Qty (ตรงกลาง) ---
            ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=7)
            ws.cell(row=9, column=5, value="Qty").alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=9, column=5).font = f_white; ws.cell(row=9, column=5).fill = fill_green; ws.cell(row=9, column=5).border = border_thin
            
            # หัวย่อย ORDER MBL BNN และหัวอื่นๆ
            headers = ['No', 'Product Code', 'Product Name', 'Unit MOM', 'ORDER', 'MBL', 'BNN']
            for i, h in enumerate(headers, 1):
                cell = ws.cell(row=10, column=i, value=h)
                cell.font, cell.fill, cell.border = f_white, fill_green, border_thin
                cell.alignment = Alignment(horizontal='center')
                if i <= 4:
                    ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                    ws.cell(row=9, column=i, value=h).alignment = Alignment(horizontal='center', vertical='center')

            # --- ตั้งค่าหน้ากระดาษให้เห็นเป็น A4 ทันที ---
            ws.sheet_view.view = "pageLayout" # << จุดสำคัญ! เปิดมาเห็นเป็นหน้ากระดาษเลย
            ws.page_setup.paperSize = 9      # A4
            ws.page_setup.orientation = 'portrait'
            ws.page_setup.fitToWidth = 1     # บีบให้พอดี 1 หน้ากระดาษในแนวขวาง
            ws.page_setup.fitToHeight = 0    # ปล่อยความสูงตามจำนวนสินค้า
            
            # ปรับ Margins (นิ้ว)
            ws.page_margins.left = 0.3
            ws.page_margins.right = 0.3
            ws.page_margins.top = 0.5
            ws.page_margins.bottom = 0.5

            # ลายเซ็น
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า").font = f_norm
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ").font = f_norm

            # ปรับความกว้างคอลัมน์ให้เหมาะสมกับ A4
            widths = {'A': 5, 'B': 14, 'C': 30, 'D': 8, 'E': 9, 'F': 9, 'G': 9}
            for col, w in widths.items(): ws.column_dimensions[col].width = w

    return output.getvalue()

# Streamlit UI
st.title("🚚 Delivery Generator (Page Layout View)")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    try:
        excel_bytes = process_excel_to_buffer(file)
        st.success("สร้างไฟล์สำเร็จ! เมื่อเปิดใน Excel จะเห็นเป็นหน้า A4 ทันที")
        st.download_button("📥 Download Final Excel", excel_bytes, "delivery_A4.xlsx")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")