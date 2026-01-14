import jsPDF from 'jspdf'
import QRCode from 'qrcode'
import { ReportData } from './pdfTypes'
import { wrapTextInCell, generateQRCode, loadThaiFont, loadImageAsBase64 } from './pdfUtils'
import { checkPreMidtermItems, checkMidtermItems, checkFinalItems } from './reportCheckers'

// ฟังก์ชันสำหรับ render ตารางส่วนหัวและรายการไฟล์ Excel ในทุกหน้า
const renderHeaderTableAndFileInfo = async (pdf: jsPDF, data: ReportData, hasThaiFont: boolean, margins: any) => {
    const setFont = (style: 'normal' | 'bold' = 'normal') => {
        pdf.setFont(hasThaiFont ? 'THSarabun' : 'helvetica', style)
    }

    const pageWidth = pdf.internal.pageSize.getWidth()
    const tableStartX = margins.left
    const tableStartY = margins.top
    const tableWidth = pageWidth - margins.left - margins.right
    const tableHeight = 40

    // กำหนดความกว้างแต่ละคอลัมน์
    const col1Width = tableWidth * 0.2
    const col2Width = tableWidth * 0.6
    const col3Width = tableWidth * 0.2
    const cellHeight = tableHeight / 4

    // วาดกรอบตาราง
    pdf.setDrawColor(0, 0, 0)
    pdf.setLineWidth(0.2)

    // เพิ่ม background สีเทาให้ตาราง
    pdf.setFillColor(240, 240, 240)
    pdf.rect(tableStartX, tableStartY, tableWidth, tableHeight, 'F')

    // วาดเส้นแนวนอน
    for (let i = 0; i <= 4; i++) {
        const y = tableStartY + (i * cellHeight)
        if (i === 0 || i === 4) {
            pdf.line(tableStartX, y, tableStartX + tableWidth, y)
        } else {
            pdf.line(tableStartX + col1Width, y, tableStartX + tableWidth, y)
        }
    }

    // วาดเส้นแนวตั้ง
    pdf.line(tableStartX, tableStartY, tableStartX, tableStartY + tableHeight)
    const col1EndX = tableStartX + col1Width
    pdf.line(col1EndX, tableStartY, col1EndX, tableStartY + tableHeight)
    const col2EndX = tableStartX + col1Width + col2Width
    pdf.line(col2EndX, tableStartY + (2 * cellHeight), col2EndX, tableStartY + tableHeight)
    const tableEndX = tableStartX + tableWidth
    pdf.line(tableEndX, tableStartY, tableEndX, tableStartY + tableHeight)

    // เพิ่มโลโก้
    try {
        const logoResponse = await fetch('/logo-ppk-512x512-1.png')
        if (logoResponse.ok) {
            const logoArrayBuffer = await logoResponse.arrayBuffer()
            const logoBase64 = btoa(
                new Uint8Array(logoArrayBuffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
            )
            const logoDataURL = `data:image/png;base64,${logoBase64}`

            const logoSize = 25
            const logoX = tableStartX + (col1Width - logoSize) / 2
            const logoY = tableStartY + (tableHeight - logoSize) / 2

            pdf.addImage(logoDataURL, 'PNG', logoX, logoY, logoSize, logoSize)
        }
    } catch (logoError) {
        console.warn('⚠️ ข้อผิดพลาดในการโหลดโลโก้:', logoError)
    }

    // เพิ่มข้อความรายงาน
    setFont('bold')
    pdf.setFontSize(16)
    const mergedCellStartX = tableStartX + col1Width
    const mergedCellWidth = col2Width + col3Width
    const textX = mergedCellStartX + (mergedCellWidth / 2)
    const textY = tableStartY + (cellHeight / 2) + 3
    pdf.text('รายงานสรุปผลการตรวจสอบ ปพ.5', textX, textY, { align: 'center' })

    // ปีการศึกษาและภาคเรียน
    setFont('normal')
    pdf.setFontSize(16)
    const mergedCell22StartX = tableStartX + col1Width
    const mergedCell22Width = col2Width + col3Width
    const cell22CenterX = mergedCell22StartX + (mergedCell22Width / 2)
    const cell22Y = tableStartY + cellHeight + (cellHeight / 2) + 2

    const academicYear = data.excelData.data?.home_academic_year || data.formData.academicYear || 'ไม่มีข้อมูล'
    const semester = data.excelData.data?.home_semester || data.formData.semester || 'ไม่มีข้อมูล'

    pdf.text(`ปีการศึกษา ${academicYear} ภาคเรียนที่ ${semester}`, cell22CenterX, cell22Y, { align: 'center' })

    // วิชา รหัสวิชา และชื่อวิชา
    setFont('normal')
    pdf.setFontSize(14)
    const cell32X = tableStartX + col1Width + 5
    const cell32Y = tableStartY + (2 * cellHeight) + (cellHeight / 2) + 2

    const subjectCode = data.excelData.data?.home_subject_code || 'ไม่มีข้อมูล'
    const subject = data.excelData.data?.home_subject || 'ไม่มีข้อมูล'

    pdf.text(`วิชา ${subjectCode} ${subject}`, cell32X, cell32Y)

    // จำนวนชั่วโมงเรียนต่อสัปดาห์
    const cell33X = tableStartX + col1Width + col2Width + 5
    const cell33Y = tableStartY + (2 * cellHeight) + (cellHeight / 2) + 2

    const studyTime = data.excelData.data?.home_study_time || 'ไม่มีข้อมูล'

    pdf.text(`${studyTime} ชั่วโมง/สัปดาห์`, cell33X, cell33Y)

    // ผู้สอน
    const cell42X = tableStartX + col1Width + 5
    const cell42Y = tableStartY + (3 * cellHeight) + (cellHeight / 2) + 2

    const teacher = data.excelData.data?.home_teacher || 'ไม่มีข้อมูล'

    pdf.text(`ผู้สอน ${teacher}`, cell42X, cell42Y)

    // จำนวนหน่วยกิต
    const cell43X = tableStartX + col1Width + col2Width + 5
    const cell43Y = tableStartY + (3 * cellHeight) + (cellHeight / 2) + 2

    const credit = data.excelData.data?.home_credit || 'ไม่มีข้อมูล'

    pdf.text(`${credit} หน่วยกิต`, cell43X, cell43Y)

    // รายการไฟล์ Excel
    const contentStartY = tableStartY + tableHeight + 8
    let yPosition = contentStartY

    if (data.excelData.fileName || data.excelData.fileSize || data.excelData.uploadedAt) {
        setFont('bold')
        pdf.setFontSize(12)
        pdf.text('รายการไฟล์ Excel:', margins.left, yPosition)
        yPosition += 5

        setFont('normal')
        pdf.setFontSize(11)

        if (data.excelData.fileName) {
            pdf.text(`1. ชื่อไฟล์: ${data.excelData.fileName}`, margins.left + 5, yPosition)
            yPosition += 5
        }

        if (data.excelData.fileSize) {
            const fileSizeInKB = (data.excelData.fileSize / 1024).toFixed(2)
            const fileSizeInMB = (data.excelData.fileSize / (1024 * 1024)).toFixed(2)
            const displayFileSize = data.excelData.fileSize < 1024 * 1024
                ? `${fileSizeInKB} KB`
                : `${fileSizeInMB} MB`
            pdf.text(`2. ขนาดไฟล์: ${displayFileSize}`, margins.left + 5, yPosition)
            yPosition += 5
        }

        if (data.excelData.uploadedAt) {
            const uploadDate = new Date(data.excelData.uploadedAt).toLocaleString('th-TH', {
                year: 'numeric',
                month: 'long',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            })
            pdf.text(`3. อัพโหลดเมื่อ: ${uploadDate}`, margins.left + 5, yPosition)
            yPosition += 5
        }
    }

    return yPosition // คืนค่าตำแหน่ง Y สำหรับเนื้อหาต่อไป
}

// ฟังก์ชันสำหรับเพิ่ม QR Code ในทุกหน้าของ PDF
const addQRCodeToAllPages = async (pdf: jsPDF, uuid: string): Promise<void> => {
    try {
        // ดึง domain จาก .env (ฝั่ง client ต้องใช้ NEXT_PUBLIC_*
        const baseUrl =
            process.env.NEXT_PUBLIC_BASE_URL ||
            (typeof window !== "undefined" ? window.location.origin : "https://yourdomain");

        const qrUrl = `${baseUrl}/verify?uuid=${encodeURIComponent(uuid)}`;
        const qrCodeDataURL = await generateQRCode(qrUrl);
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const qrSize = 25;
        const qrX = pageWidth - qrSize - 10;
        const qrY = pageHeight - qrSize - 10;

        const totalPages = pdf.getNumberOfPages();
        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            pdf.setPage(pageNum);
            pdf.addImage(qrCodeDataURL, "PNG", qrX, qrY, qrSize, qrSize);
            pdf.setFont("helvetica");
            pdf.setFontSize(5);
            const textY = qrY + qrSize + 2;
            const centerX = qrX + qrSize / 2;
            pdf.text("Verify here", centerX, textY, { align: "center" });
        }
    } catch (qrError) {
        console.warn("⚠️ ไม่สามารถสร้าง QR Code ได้:", qrError);

        // ถ้าสร้าง QR Code ไม่ได้ ให้แสดง UUID เป็นข้อความในหน้าแรกแทน
        pdf.setPage(1);
        pdf.setFont("helvetica");
        pdf.setFontSize(8);
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        pdf.text("UUID:", pageWidth - 60, pageHeight - 15);
        pdf.text(uuid, pageWidth - 60, pageHeight - 10);
    }
}

// ===== ลบซ้ำ interface/type ที่ประกาศในไฟล์นี้ออก =====

/**
 * ฟังก์ชันสำหรับสร้าง PDF รายงาน ปพ.5 จากข้อมูลที่ backend ส่งกลับ
 * จะเปิด PDF ในแทบใหม่ และถ้าไม่สามารถเปิดได้จะดาวน์โหลดอัตโนมัติ
 * @param data ข้อมูลที่ได้รับจาก backend API
 */
export const generatePDF = async (data: ReportData): Promise<void> => {
    try {
        // โหลดไอคอนก่อนสร้าง PDF
        if (!correctIconDataURL || !cancelIconDataURL) {
            await loadIcons();
        }

        const pdf = new jsPDF()

        // โหลดและเพิ่มฟอนต์ TH Sarabun
        let hasThaiFont = false
        try {
            // โหลดฟอนต์ปกติ
            const fontBase64 = await loadThaiFont('/fonts/THSarabun.ttf')
            pdf.addFileToVFS('THSarabun.ttf', fontBase64)
            pdf.addFont('THSarabun.ttf', 'THSarabun', 'normal')
            hasThaiFont = true

            // พยายามโหลดฟอนต์หนัก
            try {
                const fontBoldBase64 = await loadThaiFont('/fonts/THSarabun Bold.ttf')
                pdf.addFileToVFS('THSarabun-Bold.ttf', fontBoldBase64)
                pdf.addFont('THSarabun-Bold.ttf', 'THSarabun', 'bold')
                console.log('✅ โหลดฟอนต์ TH Sarabun (ปกติและหนัก) สำเร็จ')
            } catch (boldFontError) {
                console.warn('⚠️ ไม่สามารถโหลดฟอนต์หนักได้ - ใช้ฟอนต์ปกติแทน:', boldFontError)
            }

            pdf.setFont('THSarabun')
        } catch (fontError) {
            console.warn('⚠️ ไม่สามารถโหลดฟอนต์ TH Sarabun ได้ - ใช้ฟอนต์เริ่มต้น:', fontError)
            pdf.setFont('helvetica')
        }

        // ฟังก์ชันช่วยในการตั้งค่าฟอนต์
        const setFont = (style: 'normal' | 'bold' = 'normal') => {
            pdf.setFont(hasThaiFont ? 'THSarabun' : 'helvetica', style)
        }

        // กำหนดขอบกระดาษทั้ง 4 ด้าน (หน่วย: mm)
        const margins = {
            top: 10,
            right: 15,
            bottom: 15,
            left: 15
        }

        // ใช้ฟังก์ชัน renderHeaderTableAndFileInfo สำหรับหน้าแรก
        let yPosition = await renderHeaderTableAndFileInfo(pdf, data, hasThaiFont, margins)
        const pageWidth = pdf.internal.pageSize.getWidth()
        // ======= หน้า 1 =======
        setFont('bold')
        pdf.setFontSize(14)
        pdf.text('รายการตรวจก่อนกลางภาค', pageWidth / 2, yPosition + 3, { align: 'center' })
        yPosition += 10

        // ตรวจสอบข้อมูลทุกรายการสำหรับหน้าก่อนกลางภาค
        const preMidtermResults = checkPreMidtermItems(data)
        console.log('✅ preMidtermResults:', preMidtermResults) // สำหรับ debug

        // แยก value กับ message ออกเป็น 2 array
        const resultValues = preMidtermResults.map(r => r.value)
        const resultMessages = preMidtermResults.map(r => r.message || '')

        yPosition = drawTable(
            pdf,
            margins.left,
            yPosition,
            pageWidth - margins.left - margins.right,
            8,
            [
                'ลำดับที่', 'รายการ', 'ผลการตรวจ', 'หมายเหตุ'
            ],
            [
                'ข้อมูลระดับชั้น (ปก)',
                'ข้อมูลห้องเรียน (ปก)',
                'ภาคเรียน (ปก)',
                'ปีการศึกษา (ปก)',
                'ข้อมูลรายวิชา (ปก)',
                'รหัสวิชา (ปก)',
                'ข้อมูลกลุ่มสาระ (ปก)',
                'หน่วยกิต (ปก)',
                'เวลาเรียน (ปก)',
                'ครูผู้สอน (ปก)',
                'ครูที่ปรึกษา (ปก)',
                'ความถูกต้องของ KPA (02)',
                'เวลาเรียนรวมสอดคล้องกับหน่วยกิต (03)',
                'คะแนนเต็ม (06)',
            ],
            setFont,
            resultValues,    // ส่ง value ไปช่องผลการตรวจ
            resultMessages   // ส่ง message ไปช่องหมายเหตุ
        )
        renderSignatureSection(pdf, margins.left, yPosition, pageWidth, hasThaiFont)
        // ======= หน้า 2 =======
        pdf.addPage()
        yPosition = await renderHeaderTableAndFileInfo(pdf, data, hasThaiFont, margins)
        setFont('bold')
        pdf.setFontSize(14)
        pdf.text('รายการตรวจกลางภาค', pageWidth / 2, yPosition + 3, { align: 'center' })
        yPosition += 10

        // ใช้ checkMidtermItems เพื่อดึงผลตรวจสอบกลางภาค
        const midtermResults = checkMidtermItems(data)
        const midtermResultValues = midtermResults.map(r => r.value)
        const midtermResultMessages = midtermResults.map(r => r.message || '')

        yPosition = drawTable(
            pdf,
            margins.left,
            yPosition,
            pageWidth - margins.left - margins.right,
            8,
            ['ลำดับที่', 'รายการ', 'ผลการตรวจ', 'หมายเหตุ'],
            [
                'บันทึกเวลาเรียน (03)',
                'คะแนนก่อนกลาง (04)',
                'คะแนนกลางภาค (04)'
            ],
            setFont,
            midtermResultValues,    // ส่ง value ไปช่องผลการตรวจ
            midtermResultMessages   // ส่ง message ไปช่องหมายเหตุ
        )
        renderSignatureSection(pdf, margins.left, yPosition, pageWidth, hasThaiFont)
        // ======= หน้า 3 =======
        pdf.addPage()
        yPosition = await renderHeaderTableAndFileInfo(pdf, data, hasThaiFont, margins)
        setFont('bold')
        pdf.setFontSize(14)
        pdf.text('รายการตรวจปลายภาค', pageWidth / 2, yPosition + 3, { align: 'center' })
        yPosition += 10

        // ใช้ checkFinalItems เพื่อดึงผลตรวจสอบปลายภาค
        const finalResults = checkFinalItems(data)
        const finalResultValues = finalResults.map(r => r.value)
        const finalResultMessages = finalResults.map(r => r.message || '')

        yPosition = drawTable(
            pdf,
            margins.left,
            yPosition,
            pageWidth - margins.left - margins.right,
            8,
            ['ลำดับที่', 'รายการ', 'ผลการตรวจ', 'หมายเหตุ'],
            [
                'บันทึกเวลาเรียน (03)',
                'คะแนนหลังกลางภาค (05)',
                'คะแนนสอบปลายภาค (05)',
                'ตรวจสอบการให้ระดับผลการเรียน (06)',
                'คะแนนสมรรถนะ (07)',
                'คุณลักษณะอันพึงประสงค์ (08)',
                'คะแนนการอ่าน คิดวิเคราะห์และเขียน (09)',
                'ปีการศึกษาตรงกับ SGS (ปก)',
                'ภาคเรียนตรงกับ SGS (ปก)',
                'ระดับชั้นตรงกับ SGS (ปก)',
                'ชื่อวิชาตรงกับ SGS (ปก)',
                'ครูผู้สอนตรงกับ SGS (ปก)',
                'สรุปผลการประเมินคุณลักษณะอันพึงประสงค์ (ปพ.5 SGS)',
                'สรุปการประเมินการอ่าน คิด วิเคราะห์ และเขียน (ปพ.5 SGS)',
            ],
            setFont,
            finalResultValues,
            finalResultMessages
        )
        renderSignatureSection(pdf, margins.left, yPosition, pageWidth, hasThaiFont)

        // สร้างและเพิ่ม QR Code ลงในทุกหน้า PDF (ถ้ามี UUID)
        if (data.database?.uuid) {
            try {
                console.log('🔍 กำลังเพิ่ม QR Code ลงในทุกหน้า PDF สำหรับ UUID:', data.database.uuid)
                await addQRCodeToAllPages(pdf, data.database.uuid)
                console.log('✅ เพิ่ม QR Code ลงในทุกหน้า PDF สำเร็จ')

            } catch (qrError) {
                console.warn('⚠️ ไม่สามารถสร้าง QR Code ได้:', qrError)

                // ถ้าสร้าง QR Code ไม่ได้ ให้แสดง UUID เป็นข้อความที่หน้าสุดท้ายแทน
                const totalPages = pdf.getNumberOfPages()
                pdf.setPage(totalPages)
                setFont('normal')
                pdf.setFontSize(8)
                const pageWidth = pdf.internal.pageSize.getWidth()
                const pageHeight = pdf.internal.pageSize.getHeight()
                pdf.text('UUID:', pageWidth - 60, pageHeight - 15)
                pdf.text(data.database.uuid, pageWidth - 60, pageHeight - 10)
            }
        }

        // เพิ่มตารางหัวรายงานและรายการไฟล์ Excel ลงในทุกหน้า (ยกเว้นหน้าแรกที่มีอยู่แล้ว)
        try {
            await addHeaderToAllPages(pdf, data, hasThaiFont, margins)
        } catch (headerError) {
            console.warn('⚠️ ไม่สามารถเพิ่มตารางหัวรายงานในทุกหน้าได้:', headerError)
        }        // สร้างชื่อไฟล์
        const filename = `report-pp5-${data.formData.academicYear}-${data.formData.semester}-${Date.now()}.pdf`

        // สร้าง Blob จาก PDF และเปิดในแทบใหม่
        const pdfBlob = pdf.output('blob')
        const pdfUrl = URL.createObjectURL(pdfBlob)

        // เปิด PDF ในแทบใหม่
        const newWindow = window.open(pdfUrl, '_blank')

        if (newWindow) {
            // ตั้งชื่อหน้าต่างใหม่
            newWindow.document.title = filename

            // ล้าง URL เมื่อหน้าต่างถูกปิด (เพื่อประหยัด memory)
            newWindow.addEventListener('beforeunload', () => {
                URL.revokeObjectURL(pdfUrl)
            })

            console.log('✅ PDF สร้างและเปิดในแทบใหม่สำเร็จ:', filename)
        } else {
            // Fallback: ดาวน์โหลดถ้าไม่สามารถเปิดแทบใหม่ได้ (popup blocked) - ใช้ jsPDF save
            console.warn('⚠️ ไม่สามารถเปิดแทบใหม่ได้ (อาจถูก popup blocker บล็อก) - ดาวน์โหลดลงเครื่องแทน')
            pdf.save(filename)
            URL.revokeObjectURL(pdfUrl)
            console.log('✅ PDF ดาวน์โหลดสำเร็จ (fallback):', filename)
        }

    } catch (error) {
        console.error('❌ เกิดข้อผิดพลาดในการสร้าง PDF:', error)
        throw new Error('ไม่สามารถสร้าง PDF ได้ กรุณาลองใหม่อีกครั้ง')
    }
}

/**
 * ฟังก์ชันช่วยในการตรวจสอบโครงสร้างข้อมูล
 * @param data ข้อมูลที่ต้องการตรวจสอบ
 * @returns true หากข้อมูลมีโครงสร้างที่ถูกต้อง
 */
export const validateReportData = (data: any): data is ReportData => {
    return (
        data &&
        typeof data === 'object' &&
        data.formData &&
        data.excelData &&
        data.geminiOcrResult &&
        data.summary &&
        typeof data.formData.academicYear === 'string' &&
        typeof data.formData.semester === 'string' &&
        typeof data.formData.timestamp === 'string'
    )
}

// ฟังก์ชันสำหรับเพิ่มตารางหัวรายงานและรายการไฟล์ Excel ในทุกหน้าของ PDF
const addHeaderToAllPages = async (pdf: jsPDF, data: ReportData, hasThaiFont: boolean, margins: any): Promise<void> => {
    try {
        console.log('🔍 กำลังเพิ่มตารางหัวรายงานและรายการไฟล์ Excel ลงในทุกหน้า PDF')

        // ได้จำนวนหน้าทั้งหมด
        const totalPages = (pdf as any).internal.pages.length - 1 // หักหน้าแรกที่เป็น template

        // วนลูปเพิ่มตารางหัวใน Página 3 เป็นต้นไป (หน้าแรกและหน้าที่ 2 มีการ render ไว้แล้วใน generatePDF)
        for (let pageNum = 3; pageNum <= totalPages; pageNum++) {
            console.log(`🔍 กำลังเพิ่มตารางหัวรายงานในหน้า ${pageNum}`)

            // ไปที่หน้าที่ต้องการ
            pdf.setPage(pageNum)

            // เรียกใช้ฟังก์ชัน render header สำหรับหน้านี้
            await renderHeaderTableAndFileInfo(pdf, data, hasThaiFont, margins)
        }

        console.log('✅ เพิ่มตารางหัวรายงานและรายการไฟล์ Excel ลงในทุกหน้า PDF สำเร็จ')

    } catch (error) {
        console.warn('⚠️ ไม่สามารถเพิ่มตารางหัวรายงานในทุกหน้าได้:', error)
        throw error
    }
}


// ====== Helper: Draw Table (Refactored) ======
const drawTable = (
    pdf: jsPDF,
    startX: number,
    startY: number,
    tableWidth: number,
    cellHeight: number,
    headers: string[],
    data: string[],
    setFont: (style?: 'normal' | 'bold') => void,
    results: string[] = [],
    notes: string[] = []
) => {
    const colPercents = [0.1, 0.4, 0.1, 0.4]
    const colWidths = colPercents.map(p => tableWidth * p)
    const colPositions = colWidths.reduce((acc, w, i) => {
        acc.push((acc[i - 1] || startX) + (i > 0 ? colWidths[i - 1] : 0))
        return acc
    }, [] as number[])
    colPositions.push(startX + tableWidth)
    const tableHeight = (data.length + 1) * cellHeight

    pdf.setFillColor(250, 250, 250)
    pdf.rect(startX, startY, tableWidth, tableHeight, 'F')
    pdf.setDrawColor(0, 0, 0)
    pdf.setLineWidth(0.1)
    for (let i = 0; i <= data.length + 1; i++) {
        const y = startY + (i * cellHeight)
        pdf.line(startX, y, startX + tableWidth, y)
    }
    colPositions.forEach(x => {
        pdf.line(x, startY, x, startY + tableHeight)
    })

    // เพิ่มหัวตาราง - ใช้ wrapTextInCell เพื่อจัดการข้อความยาว
    setFont('bold')
    pdf.setFontSize(12)

    const headerTexts = ['ลำดับที่', 'รายการ', 'ผลการตรวจ', 'หมายเหตุ']
    const headerWidths = [colWidths[0], colWidths[1], colWidths[2], colWidths[3]]
    const headerPositions = [colPositions[0], colPositions[1], colPositions[2], colPositions[3]]

    // วาดหัวข้อตารางแต่ละคอลัมน์
    headerTexts.forEach((headerText, idx) => {
        const maxHeaderWidth = headerWidths[idx] - 4 // ลบ margin ซ้าย-ขวา

        // สำหรับ "ผลการตรวจสอบ" ให้ใช้ความกว้างแคบลงเพื่อบังคับให้แบ่งเป็น 2 บรรทัด
        let adjustedMaxWidth = maxHeaderWidth
        if (headerText === 'ผลการตรวจ') {
            adjustedMaxWidth = maxHeaderWidth * 0.6 // ลดความกว้างลง 40% เพื่อบังคับให้แบ่งบรรทัด
        }

        wrapTextInCell(
            pdf,
            headerText,
            headerPositions[idx] + 2, // เพิ่ม margin ซ้าย
            startY + 3, // เริ่มต้นข้อความที่ด้านบนของเซลล์
            adjustedMaxWidth,
            4, // line height เล็กลงสำหรับหัวข้อ
            headerWidths[idx],
            cellHeight,
            true, // จัดกึ่งกลางสำหรับบรรทัดเดียว (ใช้สำหรับทุกหัวข้อ)
            false // ไม่ใช้การจัดชิดซ้าย-กึ่งกลางแนวตั้ง
        )
    })

    // เพิ่มข้อมูลในตาราง
    setFont('normal')
    pdf.setFontSize(12)

    for (let i = 0; i < data.length; i++) {
        const rowY = startY + ((i + 1) * cellHeight);

        // กึ่งกลางแนวตั้ง
        const textBaselineY = rowY + (cellHeight / 2);

        // คอลัมน์ที่ 1: หมายเลขลำดับ
        const seqX = colPositions[0] + (colWidths[0] / 2);
        pdf.text((i + 1).toString(), seqX, textBaselineY, { align: 'center', baseline: 'middle' });

        // คอลัมน์ที่ 2: รายการ
        const maxTextWidth = colWidths[1] - 6;
        wrapTextInCell(
            pdf,
            data[i],
            colPositions[1] + 3,
            textBaselineY, // ใช้กึ่งกลางเซลล์
            maxTextWidth,
            4,
            colWidths[1],
            cellHeight,
            false,
            true
        );

        // คอลัมน์ที่ 3: ผลการตรวจสอบ
        const resultValue = results[i];
        const iconSize = cellHeight * 0.6; // ลดขนาดไอคอนลงประมาณ 10%
        const iconX = colPositions[2] + (colWidths[2] - iconSize) / 2; // จัดกึ่งกลางไอคอนในคอลัมน์
        const iconY = rowY + (cellHeight - iconSize) / 2; // จัดกึ่งกลางไอคอนในแถว

        if (resultValue === '1' && correctIconDataURL) {
            pdf.addImage(correctIconDataURL, 'PNG', iconX, iconY, iconSize, iconSize);
        } else if (resultValue === '0' && cancelIconDataURL) {
            pdf.addImage(cancelIconDataURL, 'PNG', iconX, iconY, iconSize, iconSize);
        } else if (resultValue !== '') {
            // ถ้าไม่ใช่ 0 หรือ 1 แต่มีค่า ให้แสดงข้อความเดิม (เผื่อกรณีอื่นในอนาคต)
            const resultText = resultValue || '';
            const resultX = colPositions[2] + (colWidths[2] / 2);
            pdf.text(resultText, resultX, textBaselineY, { align: 'center' });
        }

        // คอลัมน์ที่ 4: หมายเหตุ
        const noteText = notes[i] || '';
        const noteX = colPositions[3] + (colWidths[3] / 2);
        pdf.text(noteText, noteX, textBaselineY, { align: 'center', baseline: 'middle' });
    }

    return startY + tableHeight + 10
}

// ========== Object เก็บชื่อผู้ลงนาม (ดึงจาก Environment Variables) ========== //
const signatureNames = {
    registrar: process.env.NEXT_PUBLIC_SIGNATURE_REGISTRAR_NAME || 'นายทะเบียน (ยังไม่ได้ตั้งค่า)',
    academicHead: process.env.NEXT_PUBLIC_SIGNATURE_ACADEMIC_HEAD_NAME || 'หัวหน้าวิชาการ (ยังไม่ได้ตั้งค่า)',
    director: process.env.NEXT_PUBLIC_SIGNATURE_DIRECTOR_NAME || 'ผู้อำนวยการ (ยังไม่ได้ตั้งค่า)'
}

/**
 * ฟังก์ชันสำหรับวาดลายเซ็นต์ 3 คน (ไม่มีเส้นตาราง)
 * @param pdf jsPDF instance
 * @param x จุดเริ่มต้นแนวนอน (เช่น margins.left)
 * @param y จุดเริ่มต้นแนวตั้ง (หลังตาราง)
 * @param pageWidth ความกว้างหน้ากระดาษ
 * @param hasThaiFont ใช้ฟอนต์ไทยหรือไม่
 */
const renderSignatureSection = (
    pdf: jsPDF,
    x: number,
    y: number,
    pageWidth: number,
    hasThaiFont: boolean
) => {
    const setFont = (style: 'normal' | 'bold' = 'normal') => {
        pdf.setFont(hasThaiFont ? 'THSarabun' : 'helvetica', style)
    }
    setFont('normal')
    pdf.setFontSize(14)

    // ข้อมูลผู้เซ็นต์
    const signers = [
        { label: 'นายทะเบียน', name: signatureNames.registrar },
        { label: 'หัวหน้าวิชาการ', name: signatureNames.academicHead },
        { label: 'ผู้อำนวยการโรงเรียน', name: signatureNames.director }
    ]

    // กำหนดขนาดพื้นที่สำหรับแต่ละช่อง (ไม่มีเส้นตาราง)
    const tableWidth = pageWidth - x * 2
    const colWidth = tableWidth / 3
    const startY = y + 5 // เริ่มต้นวาดลายเซ็นต์ 5mm หลังจากจุดเริ่มต้นแนวตั้ง

    // วาดข้อความในแต่ละช่อง (จัดกึ่งกลางแนวตั้งและแนวฮอน)
    signers.forEach((signer, idx) => {
        const colX = x + idx * colWidth
        const centerX = colX + colWidth / 2
        let textY = startY + 10
        pdf.text('ลงชื่อ....................................................', centerX, textY, { align: 'center' })
        textY += 8
        pdf.text(`(${signer.name})`, centerX, textY, { align: 'center' })
        textY += 8
        pdf.text(signer.label, centerX, textY, { align: 'center' })
    })
}

// ตัวแปรสำหรับเก็บ Data URL ของรูปภาพ
let correctIconDataURL: string | null = null;
let cancelIconDataURL: string | null = null;

// ฟังก์ชันสำหรับโหลดไอคอน
const loadIcons = async () => {
    // ใช้ path แบบ relative เพื่อให้ทำงานทั้ง dev/prod
    correctIconDataURL = await loadImageAsBase64('/correct.png');
    cancelIconDataURL = await loadImageAsBase64('/cancel.png');
};