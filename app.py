import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

def process_excel_to_buffer(uploaded_file):
    # 1. อ่านข้อมูล
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

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            
            # เขียนข้อมูลสินค้าเริ่มแถว 11
            items_df = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            items_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            
            # --- Styles ---
            f_title = Font(name='Sarabun', bold=True, size=16)
            f_bold = Font(name='Sarabun', bold=True, size=10)
            f_norm = Font(name='Sarabun', size=10)
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=9)
            fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- หัวกระดาษ (Rows 1-8) ---
            ws.merge_cells('A1:G1')
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"; ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_bold
            ws['A6'] = f"Customer Name: {customer_name}"; ws['A6'].font = f_bold
            ws['A7'] = f"Store Code: {store_code}"; ws['A7'].font = f_bold
            ws['C7'] = f"Store Name: {branch_name}"; ws['C7'].font = f_bold

            # --- จัดการหัวตาราง (Rows 9-10) ---
            # Merge Qty
            ws.merge_cells('E9:G9')
            ws['E9'] = "Qty"; ws['E9'].alignment = Alignment(horizontal='center', vertical='center')
            ws['E9'].font = f_white; ws['E9'].fill = fill_green; ws['E9'].border = border
            
            # หัวตารางอื่นๆ
            h_list = ['No', 'Product Code', 'Product Name', 'Unit', 'ORDER', 'MBL', 'BNN']
            for i, h in enumerate(h_list, 1):
                cell = ws.cell(row=10, column=i, value=h)
                cell.font, cell.fill, cell.border = f_white, fill_green, border
                cell.alignment = Alignment(horizontal='center')
                if i <= 4:
                    ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                    ws.cell(row=9, column=i, value=h).alignment = Alignment(horizontal='center', vertical='center')
                    ws.cell(row=9, column=i).font, ws.cell(row=9, column=i).fill, ws.cell(row=9, column=i).border = f_white, fill_green, border

            # --- เส้นขอบข้อมูล ---
            for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
                for cell in row:
                    cell.border = border
                    cell.font = f_norm
                    if cell.column in [1, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

            # ลายเซ็น
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า").font = f_norm
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ").font = f_norm

            # --- **ส่วนสำคัญ: ปรับขนาดให้พอดี A4** ---
            ws.page_setup.paperSize = 9 # A4
            ws.page_setup.orientation = 'portrait'
            ws.print_options.horizontalCentered = True
            
            # บังคับให้อยู่ในหน้าเดียว (Fit to Page)
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            
            # ปรับความกว้างคอลัมน์ (รวมกันให้ได้ประมาณ 75-80 หน่วยสำหรับ A4)
            widths = {'A': 4.5, 'B': 14, 'C': 32, 'D': 8, 'E': 8, 'F': 8, 'G': 8}
            for col, w in widths.items():
                ws.column_dimensions[col].width = w
            
            # ปรับระยะขอบให้แคบเพื่อให้ตารางดูใหญ่ขึ้นในหน้ากระดาษ
            ws.page_margins.left = 0.25
            ws.page_margins.right = 0.25
            ws.page_margins.top = 0.5
            ws.page_margins.bottom = 0.5

    return output.getvalue()

# Streamlit UI
st.title("🚚 Delivery Generator (Perfect A4 Fit)")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    try:
        excel_bytes = process_excel_to_buffer(file)
        st.success("ประมวลผลสำเร็จ! ตารางถูกบีบให้พอดี A4 และจัดความกว้างใหม่แล้ว")
        st.download_button("📥 Download Final Excel", excel_bytes, "delivery_note_A4.xlsx")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")