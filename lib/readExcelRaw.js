/**
 * readExcelRaw.js
 * Usage: node lib/readExcelRaw.js <input.xlsx> [output.json]
 *
 * Reads all sheets from an Excel file and writes raw rows (header:1) to a JSON file
 * for offline/future processing.
 */

const fs = require('fs')
const path = require('path')
const XLSX = require('xlsx')

const input = process.argv[2]
const output = process.argv[3] || (input ? path.basename(input, path.extname(input)) + '.raw.json' : 'output.raw.json')

if (!input) {
  console.error('Usage: node lib/readExcelRaw.js <input.xlsx> [output.json]')
  process.exit(1)
}

try {
  const workbook = XLSX.readFile(input, { cellDates: true })
  const result = {
    file: path.basename(input),
    readAt: new Date().toISOString(),
    sheetCount: workbook.SheetNames.length,
    sheets: {}
  }

  workbook.SheetNames.forEach((name) => {
    const ws = workbook.Sheets[name]
    // raw rows as arrays (header:1). defval:null to preserve empty cells
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })

    // collect cell-level details (value, formula, type, formatted text, format)
    const cells = {}
    Object.keys(ws).forEach((addr) => {
      if (addr[0] === '!') return
      const c = ws[addr]
      cells[addr] = {
        address: addr,
        v: c.v === undefined ? null : c.v,
        t: c.t || null,
        f: c.f || null,
        w: c.w || null,
        z: c.z || null
      }
    })

    // build detailed rows preserving formulas and cell-level info per cell
    const rowsDetailed = []
    try {
      const ref = ws['!ref']
      const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: (rows[0] ? rows[0].length - 1 : 0) } }
      const maxRow = range.e.r + 1
      const maxCol = range.e.c + 1
      for (let r = 1; r <= maxRow; r++) {
        const rArr = []
        for (let c = 1; c <= maxCol; c++) {
          const addr = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 })
          const cell = cells[addr] || { address: addr, v: null, t: null, f: null, w: null, z: null }
          rArr.push(cell)
        }
        rowsDetailed.push(rArr)
      }
    } catch (e) {
      // fallback: include only rows as-is
    }

    // capture merges, cols, rows, and data validations if present
    const merges = ws['!merges'] || []
    const cols = ws['!cols'] || []
    const rowsMeta = ws['!rows'] || []
    const dataValidation = ws['!dataValidation'] || ws['!dv'] || null

    result.sheets[name] = {
      rows,
      rowsDetailed,
      cells,
      merges,
      cols,
      rowsMeta,
      dataValidation
    }
  })

  // capture named ranges and workbook-level metadata if present
  result.definedNames = (workbook.Workbook && workbook.Workbook.Names) ? workbook.Workbook.Names : null

  fs.writeFileSync(output, JSON.stringify(result, null, 2), 'utf8')
  console.log('✅ Raw Excel data written to', output)
} catch (err) {
  console.error('❌ Error reading/writing file:', err)
  process.exit(2)
}
