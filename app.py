import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

def process_excel_to_buffer(uploaded_file):
    # 1. อ่านข้อมูลดิบ
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    
    # หาตำแหน่งบรรทัดที่มี Item No. (ปกติคือแถวที่ 5 หรือ 6)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # Mapping Store Code จากไฟล์ต้นฉบับ
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    # 2. จัดการข้อมูลสินค้า
    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    # 3. เริ่มเขียนไฟล์ Excel
    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            
            # ข้อมูลสินค้าที่จะลงตาราง (เริ่มบรรทัดที่ 11)
            data_to_write = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "",
                'BNN': ""
            })
            data_to_write.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            
            # --- กำหนด Styles ---
            f_title = Font(name='Sarabun', bold=True, size=16)
            f_bold = Font(name='Sarabun', bold=True, size=10)
            f_norm = Font(name='Sarabun', size=10)
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)
            fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- เขียนหัวกระดาษ (Header) ใหม่ทั้งหมด ---
            # บรรทัดที่ 1: หัวเรื่อง
            ws.merge_cells('A1:G1')
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"
            ws['A1'].font = f_title
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

            # บรรทัดที่ 2-4: ข้อมูลบริษัทและวันที่ (ชิดขวา)
            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_bold
            ws['A3'] = "278 หมู่ที่ 9 ตำบลบางโฉลง"; ws['A3'].font = f_norm
            ws['A4'] = "อำเภอบางพลี จังหวัดสมุทรปราการ 10540"; ws['A4'].font = f_norm
            
            ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_bold
            ws['G3'] = "Zone: "; ws['G3'].alignment = Alignment(horizontal='right'); ws['G3'].font = f_bold
            ws['G4'] = f"Delivery: {current_date}"; ws['G4'].alignment = Alignment(horizontal='right'); ws['G4'].font = f_bold

            # บรรทัดที่ 6-7: ข้อมูลลูกค้า (Store Code / Name อยู่บรรทัดเดียวกัน)
            ws['A6'] = f"Customer Name: {customer_name}"; ws['A6'].font = f_bold
            ws['A7'] = f"Store Code: {store_code}"; ws['A7'].font = f_bold
            ws['C7'] = f"Store Name: {branch_name}"; ws['C7'].font = f_bold

            # --- จัดการหัวตารางแบบ Merge (Row 9-10) ---
            # หัวหลัก (Merge แนวตั้ง 1-4, Merge แนวนอน Qty 5-7)
            h_main = ['No', 'Product Code', 'Product Name', 'Unit MOM', 'Qty', '', '']
            h_sub = ['', '', '', '', 'ORDER', 'MBL', 'BNN']
            
            # วาดหัวตาราง
            for col, val in enumerate(h_main, 1):
                cell = ws.cell(row=9, column=col, value=val)
                cell.font, cell.fill, cell.border = f_white, fill_green, border_thin
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
            for col, val in enumerate(h_sub, 1):
                cell = ws.cell(row=10, column=col, value=val)
                cell.font, cell.fill, cell.border = f_white, fill_green, border_thin
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # สั่ง Merge จริง
            for c in range(1, 5): # No ถึง Unit
                ws.merge_cells(start_row=9, start_column=c, end_row=10, end_column=c)
            ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=7) # Qty

            # --- ตีเส้นขอบข้อมูลสินค้า ---
            for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
                for cell in row:
                    cell.border = border_thin
                    cell.font = f_norm
                    if cell.column in [1, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

            # --- ส่วนท้าย (ลายเซ็น) ---
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า").font = f_norm
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ").font = f_norm

            # --- ตั้งค่าหน้ากระดาษ A4 ---
            ws.page_setup.paperSize = 9
            ws.page_setup.orientation = 'portrait'
            ws.column_dimensions['C'].width = 35

    return output.getvalue()

# --- Streamlit UI ---
st.title("🚚 Delivery Note Generator (Fixed Header)")
file = st.file_uploader("อัปโหลดไฟล์ Excel", type="xlsx")

if file:
    try:
        excel_bytes = process_excel_to_buffer(file)
        st.success("สร้างไฟล์สำเร็จ! ตรวจสอบหัวกระดาษและ Qty ได้เลยครับ")
        st.download_button("📥 Download Excel", excel_bytes, "delivery_final.xlsx")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")