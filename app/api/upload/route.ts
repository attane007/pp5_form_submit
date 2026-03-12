import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'
import { prisma } from '@/lib/prisma'
import { ReportData } from '@/lib/pdfTypes'
import { generateGeminiContentWithRetry, getGeminiKeyPoolSize } from '@/lib/geminiKeyPool'

// ฟังก์ชันสำหรับประมวลผล PDF ด้วย Gemini
async function processPdfWithGemini(pdfFile: File) {
    // แปลง PDF เป็น base64
    const arrayBuffer = await pdfFile.arrayBuffer()
    const base64String = Buffer.from(arrayBuffer).toString('base64')

    const prompt = `
    วิเคราะห์รายงาน SGS (ปพ.5) จากไฟล์ PDF นี้และส่งคืนข้อมูลในรูปแบบ JSON ที่มีโครงสร้างดังนี้:

    {
      "course_id": "รหัสวิชา (string)",
      "course_name": "ชื่อวิชา (string)",
      "academic_year": "ปีการศึกษา พ.ศ. (string)",
      "semester": "เทอม (string)",
      "grade_level": "ระดับชั้น เช่น ม.1 (string)",
      "section": "กลุ่มเรียน (number)",
      "teacher": "ชื่อครูผู้สอน (string)",
      "grade_valid": "มีผลการเรียนมากกว่า 0 อย่างน้อย 70% ของนักเรียน (boolean)",
      "attitude_valid": "สรุปผลการ ลักษณะอันพึงประสงค์มากกว่า 0 อย่างน้อย 80% ของนักเรียนทั้งหมดในห้อง (boolean)",
      "read_analyze_write_valid": "สรุปผลการ อ่าน คิด วิเคราะห์ เขียน มากกว่า 0 อย่างน้อย 80% ของนักเรียนทั้งห้อง (boolean)"
    }

    หากข้อมูลบางอย่างไม่สามารถอ่านได้ ให้ใส่ค่าว่าง ("") สำหรับ string, 0 สำหรับ number หรือ false สำหรับ boolean
    
    ตอบกลับเป็น JSON object เท่านั้น ไม่ต้องมีข้อความอธิบายเพิ่มเติม
    `

    try {
        const primaryModelName = process.env.GEMINI_MODEL || 'gemini-2.5-flash-lite'
        const { text, attempt, totalAttempts, keyIndex, modelName } = await generateGeminiContentWithRetry({
            modelName: primaryModelName,
            responseMimeType: 'application/json',
            contents: [
                {
                    inlineData: {
                        data: base64String,
                        mimeType: 'application/pdf'
                    }
                },
                prompt
            ]
        })

        const usedFallback = modelName !== primaryModelName

        console.log(
            `[GeminiKeyPool] Success on attempt ${attempt}/${totalAttempts} using key #${keyIndex} (pool size: ${getGeminiKeyPoolSize()}) and model "${modelName}"${usedFallback ? ' (fallback)' : ' (primary)'}`
        )

        console.log('Raw Gemini response:', text)

        // ลองแปลง JSON โดยตรงก่อน
        try {
            const jsonResult = JSON.parse(text)

            // ตรวจสอบและแปลงค่าเป็น type ที่ถูกต้อง
            const processedResult = {
                course_id: String(jsonResult.course_id || ""),
                course_name: String(jsonResult.course_name || ""),
                academic_year: String(jsonResult.academic_year || ""),
                semester: String(jsonResult.semester || ""),
                grade_level: String(jsonResult.grade_level || ""),
                section: Number(jsonResult.section) || 0,
                teacher: String(jsonResult.teacher || ""),
                grade_valid: Boolean(jsonResult.grade_valid),
                attitude_valid: Boolean(jsonResult.attitude_valid),
                read_analyze_write_valid: Boolean(jsonResult.read_analyze_write_valid)
            }

            return processedResult

        } catch (parseError) {
            console.error('Direct JSON parse failed, trying to extract JSON from response')

            // พยายามแยก JSON จากการตอบกลับ
            const jsonMatch = text.match(/\{[\s\S]*\}/)
            if (jsonMatch) {
                const jsonString = jsonMatch[0]
                const extractedJson = JSON.parse(jsonString)

                // ตรวจสอบและแปลงค่าเป็น type ที่ถูกต้อง
                const processedResult = {
                    course_id: String(extractedJson.course_id || ""),
                    course_name: String(extractedJson.course_name || ""),
                    academic_year: String(extractedJson.academic_year || ""),
                    semester: String(extractedJson.semester || ""),
                    grade_level: String(extractedJson.grade_level || ""),
                    section: Number(extractedJson.section) || 0,
                    teacher: String(extractedJson.teacher || ""),
                    grade_valid: Boolean(extractedJson.grade_valid),
                    attitude_valid: Boolean(extractedJson.attitude_valid),
                    read_analyze_write_valid: Boolean(extractedJson.read_analyze_write_valid)
                }

                console.log('Extracted and processed OCR result:', processedResult)
                return processedResult
            } else {
                throw new Error('ไม่สามารถแยก JSON จากการตอบกลับของ Gemini ได้')
            }
        }
    } catch (error) {
        console.error('Error processing PDF with Gemini:', error)
        throw error
    }
}

export async function POST(request: NextRequest) {
    try {
        const data = await request.formData()
        const academicYear = data.get('academicYear') as string
        const semester = data.get('semester') as string
        const xlsxFile = data.get('file_xlsx') as File
        const pdfFile = data.get('file_pdf') as File | null

        // Log all formData keys to debug
        console.log('\n=== FormData Keys ===');
        for (const key of data.keys()) {
            console.log(`Key: ${key}, Type: ${typeof data.get(key)}, Value:`, data.get(key));
        }

        // Validate required fields
        if (!academicYear || !semester || !xlsxFile) {
            return NextResponse.json(
                { error: 'Missing required fields' },
                { status: 400 }
            )
        }

        console.log('\n=== ข้อมูลที่ได้รับ ===')
        console.log('ปีการศึกษา:', academicYear)
        console.log('ภาคเรียน:', semester)

        // ตรวจสอบว่ามีการส่งไฟล์ Excel มาหรือไม่
        if (xlsxFile) {
            console.log('✅ มีการส่งไฟล์ Excel มา:', xlsxFile.name, `(${(xlsxFile.size / 1024 / 1024).toFixed(2)} MB)`)
        } else {
            console.log('❌ ไม่มีการส่งไฟล์ Excel มา')
        }

        let pdfOcrResult = null
        if (pdfFile) {
            console.log('✅ มีการส่งไฟล์ PDF มา:', pdfFile.name, `(${(pdfFile.size / 1024 / 1024).toFixed(2)} MB)`)
            console.log('PDF File Type:', pdfFile.type)
            console.log('PDF File Size:', pdfFile.size, 'bytes')

            // ตรวจสอบว่าเป็นไฟล์ PDF จริง
            if (!pdfFile.type.includes('pdf') && !pdfFile.name.toLowerCase().endsWith('.pdf')) {
                console.error('❌ ไฟล์ที่ส่งมาไม่ใช่ไฟล์ PDF')
                console.log('ℹ️  จะข้ามการประมวลผล PDF')
            } else {
                // ส่ง PDF ไปยัง Gemini สำหรับ OCR
                try {
                    console.log('\n=== กำลังประมวลผล PDF ด้วย Gemini ===')
                    pdfOcrResult = await processPdfWithGemini(pdfFile)
                    console.log('✅ ผลการ OCR จาก PDF สำเร็จ:', pdfOcrResult)
                } catch (ocrError: any) {
                    console.error('❌ เกิดข้อผิดพลาดในการประมวลผล PDF:')
                    console.error('Error Message:', ocrError?.message || ocrError)
                    console.error('Error Stack:', ocrError?.stack)
                }
            }
        } else {
            console.log('ℹ️  ไม่มีการส่งไฟล์ PDF มา (ไม่บังคับ)')
            console.log('file_pdf value from FormData:', data.get('file_pdf'))
        }

        // อ่านไฟล์ Excel
        const arrayBuffer = await xlsxFile.arrayBuffer()
        const workbook = XLSX.read(arrayBuffer, { type: 'array' })

        console.log('\n=== รายชื่อชีท ===')
        console.log('ชีทที่มี:', workbook.SheetNames)

        let excelData: { [key: string]: any } = {}

        // ตรวจสอบว่ามีชีท "check" หรือไม่
        if (workbook.SheetNames.includes('check')) {
            console.log('\n=== ข้อมูลจากชีท "check" ===')
            const worksheet = workbook.Sheets['check']
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

            // แปลงข้อมูลให้อยู่ในรูป Object
            const convertedData: { [key: string]: any } = {}

            // กรองแถวที่มีข้อมูล (ไม่ใช่แถวว่าง)
            const validRows = jsonData.filter((row: any) =>
                Array.isArray(row) && row.length >= 2 &&
                row[0] !== undefined && row[0] !== '' &&
                row[0] !== 'list' // ข้าม header row
            )

            // แปลงจาก [key, value] เป็น object
            validRows.forEach((row: any) => {
                if (Array.isArray(row) && row.length >= 2) {
                    const key = row[0]
                    const value = row[1]
                    convertedData[key] = value
                }
            })

            console.log(convertedData)
            excelData = convertedData

        } else {
            console.log('\n❌ ไม่พบชีท "check" ในไฟล์ Excel')
            console.log('ชีทที่มี:', workbook.SheetNames)
        }

        // สร้างข้อมูลสำหรับรายงาน
        const reportData: ReportData = {
            // ข้อมูลจากฟอร์ม
            formData: {
                academicYear,
                semester,
                submittedAt: new Date().toISOString(),
                timestamp: new Date().toLocaleString('th-TH')
            },
            // ข้อมูลจากไฟล์ Excel
            excelData: excelData && Object.keys(excelData).length > 0 ? {
                hasData: true,
                sheetName: 'check',
                data: excelData,
                totalFields: Object.keys(excelData).length,
                // เพิ่มข้อมูลไฟล์ Excel
                fileName: xlsxFile.name,
                fileSize: xlsxFile.size,
                uploadedAt: new Date().toISOString()
            } : {
                hasData: false,
                message: 'ไม่พบข้อมูลในชีท "check" หรือชีทว่าง',
                // เพิ่มข้อมูลไฟล์ Excel แม้ข้อมูลจะว่าง
                fileName: xlsxFile.name,
                fileSize: xlsxFile.size,
                uploadedAt: new Date().toISOString()
            },
            // ข้อมูลจาก PDF OCR (Gemini)
            geminiOcrResult: pdfOcrResult ? {
                hasData: true,
                data: pdfOcrResult,
                processedAt: new Date().toLocaleString('th-TH')
            } : {
                hasData: false,
                message: 'ไม่มีการส่งไฟล์ PDF หรือไม่สามารถประมวลผลได้'
            },
            // สรุปผลการประมวลผล
            summary: {
                success: true,
                message: 'ประมวลผลไฟล์สำเร็จ',
                hasExcelData: excelData && Object.keys(excelData).length > 0,
                hasPdfData: !!pdfOcrResult,
                totalDataSources: (excelData && Object.keys(excelData).length > 0 ? 1 : 0) + (pdfOcrResult ? 1 : 0)
            }
        }

        // เปรียบเทียบข้อมูลจาก Gemini (PDF) กับข้อมูลจากฟอร์มและ Excel และเขียนหมายเหตุเมื่อไม่ตรง
        const remarks: string[] = []
        const gemData = reportData.geminiOcrResult?.data
        if (reportData.geminiOcrResult && reportData.geminiOcrResult.hasData && gemData) {
            const g = gemData
            const formYear = reportData.formData.academicYear
            const formSem = reportData.formData.semester
            const excel = reportData.excelData.data || {}

            // ปีการศึกษา
            if (g.academic_year) {
                const gemYear = String(g.academic_year).trim()
                const excelYear = excel.home_academic_year ? String(excel.home_academic_year).trim() : null
                if (String(formYear).trim() !== gemYear || (excelYear !== null && excelYear !== gemYear)) {
                    remarks.push(`ปีการศึกษาใน PDF (${gemYear}) ไม่ตรงกับ Form (${formYear})${excelYear ? `, Excel (${excelYear})` : ''}`)
                }
            }

            // ภาคเรียน
            if (g.semester) {
                const gemSem = String(g.semester).trim()
                const excelSem = excel.home_semester ? String(excel.home_semester).trim() : null
                if (String(formSem).trim() !== gemSem || (excelSem !== null && excelSem !== gemSem)) {
                    remarks.push(`ภาคเรียนใน PDF (${gemSem}) ไม่ตรงกับ Form (${formSem})${excelSem ? `, Excel (${excelSem})` : ''}`)
                }
            }

            // ระดับชั้น (เทียบกับ Excel หากมี)
            if (g.grade_level) {
                const gemGrade = String(g.grade_level).trim()
                const excelGrade = excel.home_grade_level ? String(excel.home_grade_level).trim() : null
                if (excelGrade && gemGrade !== excelGrade) {
                    remarks.push(`ระดับชั้นใน PDF (${gemGrade}) ไม่ตรงกับ Excel (${excelGrade})`)
                }
            }

            // ครูผู้สอน (เทียบกับ Excel หากมี)
            if (g.teacher) {
                const normalize = (s: any) => String(s || '').replace(/^(นาย|นาง|นางสาว)\\s*/, '').trim().toLowerCase()
                const gemTeacher = normalize(g.teacher)
                const excelTeacher = excel.home_teacher ? normalize(excel.home_teacher) : null
                if (excelTeacher && gemTeacher !== excelTeacher) {
                    remarks.push(`ครูผู้สอนใน PDF (${g.teacher}) ไม่ตรงกับ Excel (${excel.home_teacher})`)
                }
            }
        }

        reportData.remarks = remarks.length ? remarks.join('; ') : ''

        console.log('\n=== ข้อมูลรายงานที่จะส่งกลับ ===')
        console.log(reportData)

        // บันทึกข้อมูลลงฐานข้อมูลก่อนส่ง response
        let savedRecord = null
        try {
            console.log('\n=== กำลังบันทึกข้อมูลลงฐานข้อมูล ===')

            // ดึงชื่อครูที่รับผิดชอบจาก Excel (home_teacher) หรือใช้ค่าเริ่มต้น
            const uploaderName = (excelData && Object.keys(excelData).length > 0 && excelData.home_teacher)
                ? String(excelData.home_teacher)
                : 'Anonymous'

            console.log('🔍 Uploader Name:', uploaderName)
            console.log('🔍 Excel Data Keys:', Object.keys(excelData || {}))
            console.log('🔍 Home Teacher Value:', excelData?.home_teacher)

            // ดึง IP address และ User Agent
            const submitterIp = request.headers.get('x-forwarded-for') ||
                request.headers.get('x-real-ip') ||
                'unknown'
            const userAgent = request.headers.get('user-agent') || 'unknown'

            savedRecord = await prisma.ppkPp5Submit.create({
                data: {
                    uploaderName, // ใช้ค่าจาก home_teacher ใน Excel
                    academicYear,
                    semester,
                    xlsxFileName: xlsxFile.name,
                    xlsxFileSize: xlsxFile.size,
                    pdfFileName: pdfFile?.name || null,
                    pdfFileSize: pdfFile?.size || null,
                    status: 'PROCESSING', // เปลี่ยนเป็น PROCESSING ก่อน
                    backendResponse: JSON.parse(JSON.stringify(reportData)),
                    submitterIp,
                    userAgent
                }
            })

            console.log('✅ บันทึกข้อมูลลงฐานข้อมูลสำเร็จ - ID:', savedRecord.id)
            console.log('✅ UUID สำหรับ QR Code:', savedRecord.uuid)

            // อัปเดต status เป็น COMPLETED หลังจากบันทึกสำเร็จ
            await prisma.ppkPp5Submit.update({
                where: { id: savedRecord.id },
                data: {
                    status: 'COMPLETED',
                    processedAt: new Date()
                }
            })

        } catch (dbError) {
            console.error('❌ เกิดข้อผิดพลาดในการบันทึกลงฐานข้อมูล:', dbError)
            // throw error เพื่อไม่ให้ส่ง response หากบันทึกไม่สำเร็จ
            throw new Error('ไม่สามารถบันทึกข้อมูลลงฐานข้อมูลได้')
        }

        // เตรียม response data สำหรับส่งกลับพร้อม UUID สำหรับ QR Code
        const finalReportData = {
            ...reportData,
            // เพิ่มข้อมูล database record ID และ UUID สำหรับ QR Code
            database: {
                recordId: savedRecord.id,
                uuid: savedRecord.uuid, // สำคัญ: UUID สำหรับสร้าง QR Code
                savedAt: savedRecord.submittedAt.toISOString()
            }
        }

        console.log('✅ เตรียมส่ง UUID สำหรับ QR Code:', savedRecord.uuid)

        // ส่งข้อมูลกลับเป็น JSON
        return NextResponse.json(finalReportData, { status: 200 })

    } catch (error) {
        console.error('Upload error:', error)

        // พยายามบันทึกข้อผิดพลาดลงฐานข้อมูล
        try {
            const submitterIp = request.headers.get('x-forwarded-for') ||
                request.headers.get('x-real-ip') ||
                'unknown'
            const userAgent = request.headers.get('user-agent') || 'unknown'

            const data = await request.formData()
            const academicYear = data.get('academicYear') as string
            const semester = data.get('semester') as string
            const xlsxFile = data.get('file_xlsx') as File
            // ดึงชื่อครูที่รับผิดชอบจาก Excel (home_teacher) หากมี
            let uploaderName = 'Anonymous'
            try {
                // ลองอ่าน Excel เพื่อดึงข้อมูล home_teacher สำหรับ error logging
                if (xlsxFile) {
                    const xlsxBuffer = await xlsxFile.arrayBuffer()
                    const workbook = XLSX.read(xlsxBuffer, { type: 'buffer' })
                    const worksheet = workbook.Sheets['check']
                    if (worksheet) {
                        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
                        const headers = data[0] as string[]
                        const values = data[1] as any[]
                        const homeTeacherIndex = headers.findIndex(h =>
                            h && typeof h === 'string' && h.toLowerCase().includes('home_teacher')
                        )
                        if (homeTeacherIndex !== -1 && values[homeTeacherIndex]) {
                            uploaderName = String(values[homeTeacherIndex])
                        }
                    }
                }
            } catch (excelReadError) {
                console.warn('⚠️ ไม่สามารถอ่าน home_teacher จาก Excel สำหรับ error logging:', excelReadError)
            }

            const pdfFile = data.get('file_pdf') as File | null

            await prisma.ppkPp5Submit.create({
                data: {
                    uploaderName, // ใช้ค่าจาก home_teacher หรือ 'Anonymous'
                    academicYear: academicYear || 'unknown',
                    semester: semester || 'unknown',
                    xlsxFileName: xlsxFile?.name || 'unknown',
                    xlsxFileSize: xlsxFile?.size || 0,
                    pdfFileName: pdfFile?.name || null,
                    pdfFileSize: pdfFile?.size || null,
                    status: 'FAILED',
                    errorMessage: error instanceof Error ? error.message : 'Unknown error',
                    submitterIp,
                    userAgent
                }
            })

            console.log('✅ บันทึกข้อผิดพลาดลงฐานข้อมูลสำเร็จ')
        } catch (dbError) {
            console.error('❌ ไม่สามารถบันทึกข้อผิดพลาดลงฐานข้อมูลได้:', dbError)
        }

        return NextResponse.json(
            {
                error: 'Internal server error',
                message: error instanceof Error ? error.message : 'Unknown error'
            },
            { status: 500 }
        )
    }
}
