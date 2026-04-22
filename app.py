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
    
    # 2. หาแถวที่มีคำว่า 'Item No.' เพื่อเป็นจุดเริ่มต้นของตาราง
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # --- จุดที่แก้: กรองข้อมูลรหัสสาขาให้เป๊ะขึ้น ---
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    
    # สร้าง Dict เก็บชื่อสาขาและรหัส โดยกรองพวกค่าว่างหรือ Unnamed ออก
    branch_to_code = {}
    for name, code in zip(header_row_raw, code_row_raw):
        if name and name != 'nan' and 'Unnamed' not in name:
            branch_to_code[name] = str(code) if pd.notna(code) else ""

    # 3. อ่านข้อมูลหลักโดยใช้ Header Row ที่หาได้
    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = df.columns.str.strip()

    # --- จุดที่แก้: กรองคอลัมน์ที่เป็น Unnamed ออกก่อน Melt ---
    valid_columns = [col for col in df.columns if 'Unnamed' not in str(col)]
    df = df[valid_columns]

    id_cols = ['Item No.', 'Description', 'UNIT']
    # Melt ข้อมูล (เปลี่ยนจากตารางแนวนอนเป็นแนวตั้ง)
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    
    # --- จุดที่แก้: กรองแถวที่เป็นค่าว่าง และกรองสาขาที่เป็นสัญลักษณ์แปลกๆ ออก ---
    final_list = df_melted.dropna(subset=['Item No.', 'Qty']).query('Qty > 0')
    final_list = final_list[~final_list['Branch'].astype(str).str.contains('Unnamed|#', na=False)]

    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # --- หน้าแยกตามสาขา ---
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            # ล้างชื่อชีทให้สะอาด ไม่ให้มีอักขระพิเศษที่ Excel ไม่ชอบ
            clean_sheet_name = "".join([c for c in str(branch_name) if c.isalnum() or c in ' -_'])[:30]
            
            items_df = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            items_df.to_excel(writer, sheet_name=clean_sheet_name, index=False, header=False, startrow=10)
            ws = writer.sheets[clean_sheet_name]
            apply_styles_to_sheet(ws, branch_name, store_code, current_date, is_summary=False)

        # --- หน้าสรุปยอดรวม (Summary All) ---
        summary_all = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary_all.insert(0, 'No', range(1, len(summary_all) + 1))
        summary_all['MBL'] = ""
        summary_all['BNN'] = ""
        
        sheet_name_sum = "Summary_All"
        summary_all.to_excel(writer, sheet_name=sheet_name_sum, index=False, header=False, startrow=10)
        ws_sum = writer.sheets[sheet_name_sum]
        
        # เขียนแถวรวมยอดท้ายตาราง
        last_row = 10 + len(summary_all) + 1
        ws_sum.cell(row=last_row, column=3, value="Grand Total (ยอดรวมทั้งหมด)").font = Font(name='Cordia New', bold=True, size=14)
        ws_sum.cell(row=last_row, column=3).alignment = Alignment(horizontal='right')
        
        qty_total_cell = ws_sum.cell(row=last_row, column=5, value=summary_all['Qty'].sum())
        qty_total_cell.font = Font(name='Cordia New', bold=True, size=14)
        qty_total_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        qty_total_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
        qty_total_cell.alignment = Alignment(horizontal='center')

        apply_styles_to_sheet(ws_sum, "สรุปยอดรวมทุกรายการ", "ALL", current_date, is_summary=True)

    return output.getvalue()

def apply_styles_to_sheet(ws, branch_name, store_code, current_date, is_summary=False):
    font_name = 'Cordia New'
    f_title = Font(name=font_name, bold=True, size=20)
    f_header = Font(name=font_name, bold=True, size=14)
    f_data = Font(name=font_name, size=14)
    f_white = Font(name=font_name, bold=True, color="FFFFFF", size=14)
    fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    fill_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # จัดการ Header ข้อมูลบริษัทและชื่อสาขา
    ws.merge_cells('A1:G1')
    ws['A1'] = "ใบสรุปรายการเบิกสินค้า" if is_summary else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
    ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_header
    ws['A3'] = "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540"; ws['A3'].font = f_data
    ws['A4'] = "โทร. 02-337-1200 แฟกซ์. 02-337-1201"; ws['A4'].font = f_data
    
    ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_header
    ws['G4'] = f"Delivery Date: {current_date}"; ws['G4'].alignment = Alignment(horizontal='right'); ws['G4'].font = f_header
    
    ws['A6'] = "Customer Name"; ws['A6'].font = f_header
    ws['A7'] = f"Code: {store_code}"; ws['A7'].font = f_header
    ws['C7'] = f"Name: {branch_name}"; ws['C7'].font = f_header

    # Table Header (แถวที่ 9-10)
    ws.merge_cells('E9:G9')
    ws['E9'] = "Total Qty" if is_summary else "Qty"
    ws['E9'].font = f_white; ws['E9'].fill = fill_green; ws['E9'].border = border; ws['E9'].alignment = Alignment(horizontal='center')
    
    headers = ['No', 'Product Code', 'Product Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN']
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=10, column=i, value=h)
        cell.font, cell.fill, cell.border = f_white, fill_green, border
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # ใส่เส้นขอบและสลับสีแถวให้ดูง่าย
    for r_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7), 1):
        if ws.cell(row=row[0].row, column=3).value == "Grand Total (ยอดรวมทั้งหมด)": continue
        for cell in row:
            cell.border = border; cell.font = f_data
            if r_idx % 2 == 0: cell.fill = fill_light
            if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

    # Footer (ผู้รับ/ผู้ส่ง)
    if not is_summary:
        curr_row = ws.max_row + 2
        labels = ["ผู้รับสินค้า:", "ผู้ส่งสินค้า:", "ทะเบียนรถ:", "คลังสินค้า:"]
        for i, label in enumerate(labels):
            ws.cell(row=curr_row + i, column=1, value=f"{label} .......................................................").font = f_header
        
        # ตารางตะกร้า
        s_row = curr_row
        for i, h in enumerate(["", "MBL", "BNN"]):
            c = ws.cell(row=s_row, column=5+i, value=h)
            c.font, c.fill, c.border = f_white, fill_green, border; c.alignment = Alignment(horizontal='center')
        for i, label in enumerate(["ตะกร้าใหญ่", "ตะกร้าเล็ก"], 1):
            ws.cell(row=s_row + i, column=5, value=label).font = f_header
            for col_idx in range(5, 8):
                ws.cell(row=s_row + i, column=col_idx).border = border

    # ตั้งค่าหน้ากระดาษ
    ws.page_setup.paperSize = 9
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    widths = {'A': 6, 'B': 16, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

# --- หน้าบ้าน Streamlit ---
st.title("🚚 | POD BNN - แม่กรองขยะให้แล้วงับ | 🚚")
uploaded_file = st.file_uploader("Upload Excel", type="xlsx")

if uploaded_file:
    with st.spinner('กำลังล้างเครื่องหมาย # และคอลัมน์ขยะให้คุณนาย...'):
        try:
            excel_bytes = process_excel_to_buffer(uploaded_file)
            st.success("💅🏻 เริ่ดมาก! ข้อมูลสะอาดกริบ สมฐานะตัวแม่ 💅🏻")
            st.download_button(
                label="📥 โหลดไฟล์ที่ล้างแล้วตรงนี้ค่ะแม่", 
                data=excel_bytes, 
                file_name=f"POD_Cleaned_{datetime.now().strftime('%H%M')}.xlsx",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"อุ๊ย! ไฟล์ Excel ของแม่ท้าทายมากค่ะ มีที่ผิดตรงนี้: {e}")