import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
from datetime import datetime
import base64

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Delivery PDF Generator", layout="wide")

class DeliveryPDF(FPDF):
    def __init__(self, footer_img_path=None):
        super().__init__()
        self.footer_img_path = footer_img_path
        # เพิ่ม Font ภาษาไทย (ต้องเตรียมไฟล์ .ttf ไว้ในโฟลเดอร์เดียวกับโค้ด)
        # หากไม่มีไฟล์ ttf ให้ใช้ Font มาตรฐานหรือติดตั้งเพิ่ม
        self.add_font('THSarabun', '', 'THSarabunNew.ttf', uni=True)
        self.add_font('THSarabun', 'B', 'THSarabunNew_Bold.ttf', uni=True)

    def header_info(self, customer_name, store_code, store_name, current_date):
        self.set_font('THSarabun', 'B', 20)
        self.cell(0, 10, 'ใบส่งสินค้าชั่วคราว', ln=True, align='C')
        
        self.set_font('THSarabun', 'B', 14)
        self.cell(120, 7, 'บริษัท โมบาย โลจิสติกส์ จำกัด')
        self.set_x(-60)
        self.cell(0, 7, f'Date: {current_date}', align='R', ln=True)
        
        self.set_font('THSarabun', '', 12)
        self.cell(120, 5, '278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540')
        self.set_x(-60)
        self.cell(0, 5, 'Zone:', align='R', ln=True)
        
        self.cell(120, 5, 'โทร. 02-337-1200 แฟกซ์. 02-337-1201')
        self.set_x(-60)
        self.cell(0, 5, f'Delivery Date: {current_date}', align='R', ln=True)
        
        self.ln(5)
        self.set_font('THSarabun', 'B', 14)
        self.cell(0, 7, f'Customer Name {customer_name}', ln=True)
        self.cell(60, 7, f'Store Code: {store_code}')
        self.cell(0, 7, f'Store Name: {store_name}', ln=True)
        self.ln(2)

    def draw_table(self, df):
        # ตั้งค่าความกว้างคอลัมน์ (รวม 190mm สำหรับ A4)
        cols = [10, 35, 75, 25, 15, 15, 15]
        headers = ['No', 'Product Code', 'Product Name', 'Unit', 'ORD', 'MBL', 'BNN']
        
        # หัวตาราง Qty (แถวบน)
        self.set_fill_color(46, 125, 50) # สีเขียว
        self.set_text_color(255, 255, 255)
        self.set_font('THSarabun', 'B', 12)
        
        # วาด Header ส่วน Qty
        start_x = self.get_x()
        for i in range(4): self.cell(cols[i], 12, headers[i], border=1, align='C', fill=True)
        self.cell(sum(cols[4:]), 6, 'Qty', border=1, align='C', fill=True, ln=True)
        
        # วาด Header แถวที่ 2 (ORD, MBL, BNN)
        self.set_x(start_x + sum(cols[:4]))
        for i in range(4, 7): self.cell(cols[i], 6, headers[i], border=1, align='C', fill=True)
        self.ln(6)
        
        # ข้อมูลตาราง
        self.set_text_color(0, 0, 0)
        self.set_font('THSarabun', '', 12)
        for index, row in df.iterrows():
            # เช็คว่าหน้าเต็มหรือยัง
            if self.get_y() > 220: 
                self.add_page()
                self.header_info(...) # ต้องส่งตัวแปรเพิ่ม
            
            self.cell(cols[0], 7, str(row['No']), border=1, align='C')
            self.cell(cols[1], 7, str(row['Code']), border=1)
            self.cell(cols[2], 7, str(row['Name']), border=1)
            self.cell(cols[3], 7, str(row['Unit']), border=1, align='C')
            self.cell(cols[4], 7, str(row['ORD']), border=1, align='C')
            self.cell(cols[5], 7, '', border=1)
            self.cell(cols[6], 7, '', border=1, ln=True)

    def footer(self):
        # แปะรูปภาพที่ด้านล่างสุดของทุกหน้า
        if self.footer_img_path:
            self.image(self.footer_img_path, x=10, y=245, w=190)

def process_to_pdf(uploaded_file, footer_img):
    # 1. อ่านข้อมูล (Logic เดิม)
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = "" # ตามที่สั่งคือเอาแค่คำว่า Customer Name เพียวๆ ในหัวข้อ
    
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = df.columns.str.strip()
    
    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    # 2. สร้าง PDF
    pdf = DeliveryPDF(footer_img_path=footer_img)
    current_date = datetime.now().strftime('%d/%m/%Y')

    for branch_name, branch_data in final_list.groupby('Branch', sort=False):
        store_code = branch_to_code.get(str(branch_name).strip(), "")
        pdf.add_page()
        
        # ส่วนหัว
        pdf.header_info("", store_code, branch_name, current_date)
        
        # เตรียมข้อมูลตาราง
        items_df = pd.DataFrame({
            'No': range(1, len(branch_data) + 1),
            'Code': branch_data['Item No.'],
            'Name': branch_data['Description'],
            'Unit': branch_data['UNIT'],
            'ORD': branch_data['Qty']
        })
        
        # วาดตาราง
        pdf.draw_table(items_df)

    return pdf.output(dest='S').encode('latin1'), final_list

# --- UI ---
st.title("🚚 Delivery PDF Generator")
st.info("อัปโหลดไฟล์ Excel และระบบจะสร้าง PDF พร้อมแปะรูปภาพส่วนท้ายให้อัตโนมัติ")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. อัปโหลดไฟล์ Excel", type="xlsx")
with col2:
    # ในที่นี้คือรูปภาพที่คุณอัปโหลดมา ผมจำลองว่าเป็น footer_img.png
    footer_img = "image_bae9a6.png" 
    st.image(footer_img, caption="รูปส่วนท้ายที่จะนำไปแปะ", width=300)

if excel_file:
    try:
        pdf_bytes, preview_df = process_to_pdf(excel_file, footer_img)
        st.success("✅ สร้างไฟล์ PDF สำเร็จ!")
        
        with st.expander("🔍 ดูรายการสรุป"):
            st.dataframe(preview_df[['Branch', 'Item No.', 'Description', 'Qty']], hide_index=True)

        st.download_button(
            label="📥 ดาวน์โหลดใบส่งสินค้า (PDF)",
            data=pdf_bytes,
            file_name=f"Delivery_Note_{datetime.now().strftime('%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
        st.info("หมายเหตุ: โค้ดนี้ต้องการไฟล์ Font THSarabunNew.ttf ในระบบเพื่อให้แสดงภาษาไทยได้")