/**
 * readExcelRaw.js (Robust Version)
 * ดึงข้อมูลละเอียด: Data, Formulas, Merged Cells, Data Validation, Styles
 */

const fs = require('fs');
const path = require('path');
const Excel = require('exceljs');

const input = process.argv[2];
const output = process.argv[3] || (input ? path.basename(input, path.extname(input)) + '.raw.json' : 'output.raw.json');

if (!input) {
    console.error('Usage: node lib/readExcelRaw.js <input.xlsx> [output.json]');
    process.exit(1);
}

// Helper: แปลงเลข Column เป็นตัวอักษร (1 -> A, 27 -> AA)
function colNumberToLetter(n) {
    let s = '';
    while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

async function run() {
    try {
        const workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(input);

        const result = {
            file: path.basename(input),
            readAt: new Date().toISOString(),
            sheetCount: workbook.worksheets.length,
            sheets: {}
        };

        workbook.eachSheet((ws) => {
            // หาขอบเขตสูงสุดของ Sheet
            let maxRow = ws.actualRowCount || 0;
            let maxCol = 0;
            ws.eachRow((row) => {
                maxCol = Math.max(maxCol, row.cellCount || 0);
            });

            const sheetData = {
                name: ws.name,
                merges: ws.model.merges || [],
                rowsDetailed: [] // เก็บข้อมูลแบบ Array 2 มิติ
            };

            // วนลูปอ่านข้อมูลทีละแถวตามลำดับ
            for (let r = 1; r <= maxRow; r++) {
                const row = ws.getRow(r);
                const rowArray = [];

                for (let c = 1; c <= maxCol; c++) {
                    const cell = row.getCell(c);
                    const addr = cell.address;

                    // ป้องกัน Error: เข้าถึง Master Cell อย่างปลอดภัย
                    let targetCell = cell;
                    if (cell.isMerged && cell.master) {
                        targetCell = cell.master;
                    }

                    const info = {
                        address: addr,
                        row: r,
                        col: c,
                        t: cell.type, // ExcelJS Cell Type
                        v: null,      // Value
                        f: null,      // Formula
                        w: null,      // Text ที่แสดงผล
                        style: (cell.style && Object.keys(cell.style).length) ? cell.style : null,
                        validation: cell.dataValidation || null,
                        isMerged: cell.isMerged,
                        master: cell.master ? cell.master.address : addr
                    };

                    // ดึงค่า Value และ Formula (ลำดับความสำคัญ: สูตร > ค่า)
                    try {
                        const val = targetCell.value;
                        if (val !== null && typeof val === 'object' && val.formula) {
                            info.f = val.formula;
                            info.v = val.result !== undefined ? val.result : null;
                        } else {
                            info.v = val;
                        }

                        // ดึง Text ที่แสดงผล (ปลอดภัยกว่า .text)
                        if (val !== null && val !== undefined) {
                            info.w = info.f ? (info.v !== null ? info.v.toString() : "") : val.toString();
                        }
                    } catch (e) {
                        info.w = ""; // กรณีเกิด Error ในการแปลงค่า
                    }

                    rowArray.push(info);
                }
                sheetData.rowsDetailed.push(rowArray);
            }

            result.sheets[ws.name] = sheetData;
        });

        // บันทึกไฟล์ JSON
        fs.writeFileSync(output, JSON.stringify(result, null, 2), 'utf8');
        console.log(`\x1b[32m✅ สำเร็จ: เขียนข้อมูลลงใน ${output}\x1b[0m`);

    } catch (err) {
        console.error('\x1b[31m❌ เกิดข้อผิดพลาดในการอ่านไฟล์:\x1b[0m', err.message);
        process.exit(1);
    }
}

run();