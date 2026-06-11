<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🚚 | POD BNN - Uniform Edition | 🚚</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f0f2f6;
            color: #31333F;
            margin: 0;
            padding: 2rem;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .container {
            background-color: white;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            width: 100%;
            max-width: 600px;
            text-align: center;
        }
        h1 { font-size: 1.5rem; margin-bottom: 1.5rem; color: #1E293B; }
        .upload-area {
            border: 2px dashed #10B981;
            border-radius: 10px;
            padding: 2rem;
            cursor: pointer;
            margin-bottom: 1.5rem;
            background-color: #f0fDF4;
            transition: 0.3s;
            display: block;
        }
        .upload-area:hover { background-color: #DCFCE7; }
        input[type="file"] { display: none; }
        .btn-download {
            background-color: #10B981;
            color: white;
            border: none;
            padding: 12px 20px;
            border-radius: 5px;
            font-size: 1rem;
            font-weight: bold;
            cursor: pointer;
            display: none;
            width: 100%;
            margin-top: 1rem;
            transition: 0.2s;
        }
        .btn-download:hover { background-color: #059669; }
        .status { margin-top: 1rem; font-weight: bold; color: #4B5563; }
        .spinner {
            display: none;
            border: 4px solid #E5E7EB;
            border-top: 4px solid #10B981;
            border-radius: 50%;
            width: 35px;
            height: 35px;
            animation: spin 1s linear infinite;
            margin: 15px auto;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>

<div class="container">
    <h1>🚚 | POD BNN - Uniform Edition | 🚚</h1>
    
    <label class="upload-area" id="drop-area">
        <input type="file" id="file-upload" accept=".xlsx" />
        <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">👕</div>
        <p>ลากไฟล์มาวางตรงนี้ หรือ <b>คลิกเพื่ออัปโหลด Excel</b></p>
    </label>

    <div class="spinner" id="spinner"></div>
    <div class="status" id="status-text"></div>
    <button class="btn-download" id="btn-download">📥 โหลดไฟล์เวอร์ชันปริ้นท์ตรงนี้ค่ะแม่</button>
</div>

<script>
    document.getElementById('file-upload').addEventListener('change', async function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const spinner = document.getElementById('spinner');
        const statusText = document.getElementById('status-text');
        const btnDownload = document.getElementById('btn-download');

        spinner.style.display = 'block';
        statusText.innerHTML = 'กำลังจัดฟอร์มและสร้าง Tab เบิก Uniform ให้อยู่นะคะแม่...';
        btnDownload.style.display = 'none';

        try {
            const arrayBuffer = await file.arrayBuffer();
            const readWb = new ExcelJS.Workbook();
            await readWb.xlsx.load(arrayBuffer);
            const rawSheet = readWb.worksheets[0];

            let headerRowIndex = -1;
            let colIndices = { itemNo: 1, description: 2, unit: 3 };

            // 1. ค้นหาแถวที่เป็นหัวตาราง (รองรับทั้ง ไทย/อังกฤษ)
            for (let r = 1; r <= rawSheet.rowCount; r++) {
                let row = rawSheet.getRow(r);
                let isHeader = false;
                
                for (let c = 1; c <= row.cellCount; c++) {
                    let cellVal = String(row.getCell(c).value || '');
                    if (cellVal.includes('Item No.') || cellVal.includes('รหัสสินค้า')) {
                        isHeader = true;
                        break;
                    }
                }
                
                if (isHeader) {
                    headerRowIndex = r;
                    for (let c = 1; c <= row.cellCount; c++) {
                        let cellVal = String(row.getCell(c).value || '').trim();
                        if (cellVal.includes('Item No.') || cellVal.includes('รหัสสินค้า')) colIndices.itemNo = c;
                        if (cellVal.includes('Description') || cellVal.includes('ชื่อสินค้า')) colIndices.description = c;
                        if (cellVal.includes('UNIT') || cellVal.includes('หน่วย')) colIndices.unit = c;
                    }
                    break;
                }
            }

            if (headerRowIndex === -1) throw new Error("ไม่พบคอลัมน์รหัสสินค้าหรือ Item No. ในไฟล์ค่ะ");

            const headerRow = rawSheet.getRow(headerRowIndex);
            const nextRow = rawSheet.getRow(headerRowIndex + 1);
            const nextRowFirstCell = String(nextRow.getCell(colIndices.itemNo).value || '').trim();

            // ตรวจสอบว่าแถวถัดไปเป็นข้อมูลเลย (แบบไฟล์แบน) หรือเป็นรหัสสาขา
            let isNextRowData = false;
            if (nextRowFirstCell !== '' && nextRowFirstCell !== 'null' && !nextRowFirstCell.includes('Unnamed')) {
                if (nextRowFirstCell.startsWith('EX') || nextRowFirstCell.match(/[A-Za-z0-9]+/)) {
                    isNextRowData = true;
                }
            }

            // 2. ดึงข้อมูลรายชื่อสาขาและรหัสสาขา (ถ้ามี)
            const branchToCode = {};
            const branchCols = [];
            let maxIdCol = Math.max(colIndices.itemNo, colIndices.description, colIndices.unit);
            
            headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                if (colNumber > maxIdCol) {
                    const branchName = String(cell.value || '').trim();
                    if (branchName && branchName !== 'null' && !branchName.includes('Unnamed') && !branchName.includes('Grand Total') && !branchName.includes('รวม')) {
                        let code = '';
                        if (nextRow && !isNextRowData) {
                            const codeVal = nextRow.getCell(colNumber).value;
                            code = (codeVal !== null && codeVal !== undefined) ? String(codeVal).trim() : '';
                        }
                        branchToCode[branchName] = code;
                        branchCols.push({ index: colNumber, name: branchName });
                    }
                }
            });

            // 3. ทรานส์ฟอร์มข้อมูล (Melt Data)
            const startDataRow = headerRowIndex + (isNextRowData ? 1 : 2);
            const finalData = [];
            const summaryAll = {};

            for (let r = startDataRow; r <= rawSheet.rowCount; r++) {
                const row = rawSheet.getRow(r);
                const itemNo = String(row.getCell(colIndices.itemNo).value || '').trim();
                const description = String(row.getCell(colIndices.description).value || '').trim();
                const unit = String(row.getCell(colIndices.unit).value || '').trim();

                if (!itemNo || itemNo === 'null' || itemNo === 'Grand Total' || itemNo.includes('ยอดรวม')) continue;

                branchCols.forEach(branch => {
                    const qtyVal = row.getCell(branch.index).value;
                    const qty = Number(qtyVal);
                    if (!isNaN(qty) && qty > 0) {
                        finalData.push({
                            branch: branch.name, itemNo, description, unit, qty
                        });

                        const sumKey = `${itemNo}|${description}|${unit}`;
                        if (!summaryAll[sumKey]) {
                            summaryAll[sumKey] = { itemNo, description, unit, qty: 0 };
                        }
                        summaryAll[sumKey].qty += qty;
                    }
                });
            }

            // กลุ่มข้อมูลตามสาขาเพื่อสร้างชีทระบบเดิม
            const branchGroups = finalData.reduce((acc, curr) => {
                if (!acc[curr.branch]) acc[curr.branch] = [];
                acc[curr.branch].push(curr);
                return acc;
            }, {});

            // 4. สร้างไฟล์ Excel ผลลัพธ์อันใหม่
            const writeWb = new ExcelJS.Workbook();
            const dateStr = new Date().toLocaleDateString('th-TH');

            // ฟังก์ชันจัดสไตล์หน้าตาตารางให้แพง ปริ้นท์สวย
            const applyStyles = (ws, branchName, storeCode, isSummary) => {
                const fontName = 'Cordia New';
                const fTitle = { name: fontName, bold: true, size: 20 };
                const fHeader = { name: fontName, bold: true, size: 14 };
                const fData = { name: fontName, size: 14 };
                const fBlackBold = { name: fontName, bold: true, color: { argb: 'FF000000' }, size: 14 };
                const fillLightGreen = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8E6C9' } };
                const border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

                ws.mergeCells('A1:G1');
                ws.getCell('A1').value = isSummary ? "ใบสรุปรายการเบิกสินค้า (Uniform)" : "ใบส่งสินค้าชั่วคราว";
                ws.getCell('A1').font = fTitle;
                ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };

                ws.getCell('A2').value = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws.getCell('A2').font = fHeader;
                ws.getCell('A3').value = "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540"; ws.getCell('A3').font = fData;
                ws.getCell('A4').value = "โทร. 02-337-1200 แฟกซ์. 02-337-1201"; ws.getCell('A4').font = fData;

                ws.getCell('G2').value = `Date: ${dateStr}`; ws.getCell('G2').font = fHeader; ws.getCell('G2').alignment = { horizontal: 'right' };
                ws.getCell('G4').value = `Delivery Date: ${dateStr}`; ws.getCell('G4').font = fHeader; ws.getCell('G4').alignment = { horizontal: 'right' };

                ws.getCell('A7').value = `Code: ${storeCode}`; ws.getCell('A7').font = fHeader;
                ws.getCell('C7').value = `Name: ${branchName}`; ws.getCell('C7').font = fHeader;

                ws.mergeCells('E9:G9');
                const e9 = ws.getCell('E9');
                e9.value = isSummary ? "Total Qty" : "Qty";
                e9.font = fBlackBold; e9.fill = fillLightGreen; e9.border = border; e9.alignment = { horizontal: 'center' };
                ws.getCell('F9').border = border; ws.getCell('F9').fill = fillLightGreen;
                ws.getCell('G9').border = border; ws.getCell('G9').fill = fillLightGreen;

                const headers = ['No', 'Product Code', 'Product Name', 'Unit', isSummary ? 'TOTAL' : 'ORDER', 'MBL', 'BNN'];
                headers.forEach((h, i) => {
                    const cell = ws.getCell(10, i + 1);
                    cell.value = h;
                    cell.font = fBlackBold; cell.fill = fillLightGreen; cell.border = border;
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                });

                ws.columns = [
                    { width: 6 }, { width: 16 }, { width: 35 }, { width: 10 }, { width: 12 }, { width: 10 }, { width: 10 }
                ];
                
                ws.pageSetup = { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 1 };
                return border;
            };

            // สร้างชีทแยกตามสาขา (ระบบเดิมยังอยู่ครบ)
            for (const [branchName, items] of Object.entries(branchGroups)) {
                let cleanName = branchName.replace(/[^a-zA-Z0-9ก-๙ \-_]/g, '').substring(0, 30);
                const ws = writeWb.addWorksheet(cleanName);
                const code = branchToCode[branchName] || "";
                
                const borderStyle = applyStyles(ws, branchName, code, false);

                let currentRow = 11;
                items.forEach((item, index) => {
                    ws.getRow(currentRow).values = [index + 1, item.itemNo, item.description, item.unit, item.qty, '', ''];
                    
                    for(let c = 1; c <= 7; c++) {
                        let cell = ws.getCell(currentRow, c);
                        cell.font = { name: 'Cordia New', size: 14 };
                        cell.border = borderStyle;
                        if ([1, 4, 5, 6, 7].includes(c)) cell.alignment = { horizontal: 'center' };
                    }
                    currentRow++;
                });

                // ท้ายกระดาษ + ลายเซ็น
                currentRow += 2;
                ["ผู้รับสินค้า:", "ผู้ส่งสินค้า:", "ทะเบียนรถ:", "คลังสินค้า:"].forEach((label, i) => {
                    ws.getCell(currentRow + i, 1).value = `${label} .......................................................`;
                    ws.getCell(currentRow + i, 1).font = { name: 'Cordia New', bold: true, size: 14 };
                });

                // ตารางใส่จำนวนตะกร้าขนส่ง
                let basketRow = currentRow;
                ["", "MBL", "BNN"].forEach((h, i) => {
                    let c = ws.getCell(basketRow, 5 + i);
                    c.value = h; c.font = { name: 'Cordia New', bold: true, size: 14 };
                    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8E6C9' } };
                    c.border = borderStyle; c.alignment = { horizontal: 'center' };
                });
                ["ตะกร้าใหญ่", "ตะกร้าเล็ก"].forEach((label, i) => {
                    ws.getCell(basketRow + i, 5).value = label;
                    ws.getCell(basketRow + i, 5).font = { name: 'Cordia New', bold: true, size: 14 };
                    for(let c = 5; c <= 7; c++) ws.getCell(basketRow + i, c).border = borderStyle;
                });
            }

            // 5. สร้างหน้าสรุปยอดรวม (เปลี่ยนชื่อเป็น "เบิก Uniform" ตามสั่ง)
            const wsSum = writeWb.addWorksheet("เบิก Uniform");
            const sumBorder = applyStyles(wsSum, "สรุปยอดรวมทุกรายการ", "ALL", true);
            
            let sumRow = 11;
            let totalQtySum = 0;
            Object.values(summaryAll).forEach((item, index) => {
                wsSum.getRow(sumRow).values = [index + 1, item.itemNo, item.description, item.unit, item.qty, '', ''];
                totalQtySum += item.qty;
                for(let c = 1; c <= 7; c++) {
                    let cell = wsSum.getCell(sumRow, c);
                    cell.font = { name: 'Cordia New', size: 14 };
                    cell.border = sumBorder;
                    if ([1, 4, 5, 6, 7].includes(c)) cell.alignment = { horizontal: 'center' };
                }
                sumRow++;
            });

            // แถว Grand Total (สรุปยอดท้ายสุด)
            wsSum.getCell(sumRow, 3).value = "Grand Total (ยอดรวมทั้งหมด)";
            wsSum.getCell(sumRow, 3).font = { name: 'Cordia New', bold: true, size: 14 };
            
            let qtyTotalCell = wsSum.getCell(sumRow, 5);
            qtyTotalCell.value = totalQtySum;
            qtyTotalCell.font = { name: 'Cordia New', bold: true, size: 14 };
            qtyTotalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // เน้นไฮไลท์สีเหลือง
            qtyTotalCell.border = { top: {style:'thin'}, bottom: {style:'double'}, left: {style:'thin'}, right: {style:'thin'} };
            qtyTotalCell.alignment = { horizontal: 'center' };

            // คอนเวิร์ตไฟล์ออกเป็น Blob เพื่อดาวน์โหลด
            const outBuffer = await writeWb.xlsx.writeBuffer();
            const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            
            const timeStr = new Date().toTimeString().slice(0,5).replace(':','');
            
            spinner.style.display = 'none';
            statusText.innerHTML = '<span style="color: #059669;">💅🏻 เรียบร้อยค่ะแม่! แยกชีทรายสาขาครบ พร้อมมีหน้า "เบิก Uniform" ให้แล้วค่ะ 💅🏻</span>';
            
            btnDownload.style.display = 'block';
            btnDownload.onclick = function() {
                saveAs(blob, `POD_UniformRequisition_${timeStr}.xlsx`);
            };

        } catch (error) {
            spinner.style.display = 'none';
            statusText.innerHTML = `<span style="color: #DC2626;">อุ๊ย! เกิดข้อผิดพลาด: ${error.message}</span>`;
            console.error(error);
        }
    });
</script>

</body>
</html>
