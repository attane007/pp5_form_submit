// Type definitions สำหรับข้อมูลที่ backend ส่งกลับ
export interface FormData {
    academicYear: string
    semester: string
    submittedAt: string
    timestamp: string
}

export interface ExcelData {
    hasData: boolean
    sheetName?: string
    data?: { [key: string]: any }
    totalFields?: number
    message?: string
    fileName?: string
    fileSize?: number
    uploadedAt?: string
}

export interface GeminiOcrData {
    course_id: string
    course_name: string
    academic_year: string
    semester: string
    grade_level: string
    section: number
    teacher: string
    grade_valid: boolean
    attitude_valid: boolean
    read_analyze_write_valid: boolean
}

export interface GeminiOcrResult {
    hasData: boolean
    data?: GeminiOcrData
    processedAt?: string
    message?: string
}

export interface Summary {
    success: boolean
    message: string
    hasExcelData: boolean
    hasPdfData: boolean
    totalDataSources: number
}

export interface DatabaseInfo {
    recordId: string
    uuid: string
    savedAt: string
}

export interface ReportData {
    formData: FormData
    excelData: ExcelData
    geminiOcrResult: GeminiOcrResult
    summary: Summary
    remarks?: string
    database?: DatabaseInfo
}
