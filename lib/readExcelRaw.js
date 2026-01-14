/**
 * readExcelRaw.js (exceljs)
 * Usage: node lib/readExcelRaw.js <input.xlsx> [output.json]
 *
 * Reads all sheets from an Excel file and writes raw rows and per-cell details to a JSON file
 * preserving formulas (preferred over values) and dataValidation info so the raw JSON can be
 * used to reconstruct an identical workbook (formulas, dropdowns, merges, formats)
 */

const fs = require('fs')
const path = require('path')
const Excel = require('exceljs')

const input = process.argv[2]
const output = process.argv[3] || (input ? path.basename(input, path.extname(input)) + '.raw.json' : 'output.raw.json')

if (!input) {
  console.error('Usage: node lib/readExcelRaw.js <input.xlsx> [output.json]')
  process.exit(1)
}

function colNumberToLetter(n) {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

async function run() {
  try {
    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(input)

    const result = {
      file: path.basename(input),
      readAt: new Date().toISOString(),
      sheetCount: workbook.worksheets.length,
      sheets: {}
    }

    // defined names (best-effort)
    try {
      result.definedNames = workbook.model && workbook.model.definedNames ? workbook.model.definedNames : null
    } catch (e) {
      result.definedNames = null
    }

    workbook.eachSheet((ws) => {
      // determine max used row/col
      let maxRow = ws.actualRowCount || ws.rowCount || 0
      let maxCol = 0
      ws.eachRow((row) => {
        maxCol = Math.max(maxCol, row.cellCount || 0)
      })

      // rows as simple arrays of displayed values (null for empty)
      const rows = []
      for (let r = 1; r <= maxRow; r++) {
        const rr = []
        const row = ws.getRow(r)
        for (let c = 1; c <= maxCol; c++) {
          const cell = row.getCell(c)
          rr.push(cell && cell.text !== undefined && cell.text !== '' ? cell.text : null)
        }
        rows.push(rr)
      }

      const cells = {}
      // gather cell-level info using exceljs objects
      ws.eachRow((row) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const addr = cell.address
          const info = {
            address: addr,
            row: cell.row,
            col: cell.col,
            t: cell.type !== undefined ? cell.type : null,
            w: cell.text !== undefined ? cell.text : null,
            z: (cell.numFmt !== undefined ? cell.numFmt : null),
            style: (cell.style && Object.keys(cell.style).length ? cell.style : null),
            f: null,
            v: null,
            dataValidation: null
          }

          // formula prioritized over value
          try {
            if (cell.value && typeof cell.value === 'object' && Object.prototype.hasOwnProperty.call(cell.value, 'formula')) {
              info.f = cell.value.formula || null
              info.result = cell.value.result !== undefined ? cell.value.result : null
              info.v = null
            } else {
              info.v = cell.value !== undefined ? cell.value : null
            }
          } catch (e) {
            info.v = cell.value !== undefined ? cell.value : null
          }

          // per-cell dataValidation
          try {
            if (cell.dataValidation) {
              info.dataValidation = cell.dataValidation
            }
          } catch (e) {
            info.dataValidation = null
          }

          cells[addr] = info
        })
      })

      // rowsDetailed built from cells ensuring formula prioritized
      const rowsDetailed = []
      for (let r = 1; r <= maxRow; r++) {
        const rArr = []
        for (let c = 1; c <= maxCol; c++) {
          const addr = colNumberToLetter(c) + r
          const cellInfo = cells[addr] || { address: addr, v: null, f: null, t: null, w: null, z: null, style: null, dataValidation: null }
          rArr.push(cellInfo)
        }
        rowsDetailed.push(rArr)
      }

      // merges
      let merges = []
      try {
        if (ws.model && ws.model.merges) merges = ws.model.merges
        else if (ws._merges) merges = Array.from(ws._merges.keys())
      } catch (e) {
        merges = []
      }

      // cols
      let cols = null
      try {
        cols = ws.columns && ws.columns.length ? ws.columns.map((c, idx) => ({ key: c.key || idx + 1, width: c.width || null, hidden: c.hidden || false })) : null
      } catch (e) {
        cols = null
      }

      // rows meta (height/hidden)
      const rowsMeta = []
      try {
        for (let r = 1; r <= maxRow; r++) {
          const row = ws.getRow(r)
          if (row && (row.height || row.hidden)) {
            rowsMeta.push({ row: r, height: row.height || null, hidden: row.hidden || false })
          }
        }
      } catch (e) {
        // ignore
      }

      // worksheet-level dataValidation
      let dataValidation = null
      try {
        if (ws.dataValidations && ws.dataValidations.model) dataValidation = ws.dataValidations.model
        else if (ws.model && ws.model.dataValidations) dataValidation = ws.model.dataValidations
      } catch (e) {
        dataValidation = null
      }

      result.sheets[ws.name] = {
        rows,
        rowsDetailed,
        cells,
        merges,
        cols,
        rowsMeta,
        dataValidation
      }
    })

    fs.writeFileSync(output, JSON.stringify(result, null, 2), 'utf8')
    console.log('✅ Raw Excel data written to', output)
  } catch (err) {
    console.error('❌ Error reading/writing file:', err)
    process.exit(2)
  }
}

run()
