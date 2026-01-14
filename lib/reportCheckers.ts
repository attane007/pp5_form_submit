import { ReportData } from './pdfTypes'

export type CheckResult = { value: '1' | '0' | ''; message?: string }

export function checkPreMidtermItems(data: ReportData): CheckResult[] {
    return [
        checkGradeLevel(data),
        checkRoom(data),
        checkSemester(data),
        checkAcademicYear(data),
        checkSubject(data),
        checkSubjectCode(data),
        checkLearningArea(data),
        checkCredit(data),
        checkStudyTime(data),
        checkTeacher(data),
        checkAdvisor(data),
        checkKPA(data),
        checkHourCredit(data),
        checkMidtermScore(data),
    ]
}

// ตรวจสอบว่ามีข้อมูลระดับชั้นหรือไม่
function checkGradeLevel(data: ReportData): CheckResult {
    return data.excelData.data?.home_grade_level ? { value: '1' } : { value: '0', message: 'ไม่มีข้อมูลระดับชั้น' }
}

// ตรวจสอบว่ามีข้อมูลห้องเรียนหรือไม่
function checkRoom(data: ReportData): CheckResult {
    return data.excelData.data?.home_room ? { value: '1' } : { value: '0', message: 'ไม่มีข้อมูลห้องเรียน' }
}

// ตรวจสอบว่าข้อมูลภาคเรียนในฟอร์ม ตรงกับข้อมูลใน excel หรือไม่
function checkSemester(data: ReportData): CheckResult {
    const formSemester = data.formData.semester
    const excelSemester = data.excelData.data?.home_semester
    if (formSemester == null || excelSemester == null) {
        return { value: '0', message: 'ข้อมูลภาคเรียนไม่ครบถ้วน' }
    }
    if (String(formSemester) === String(excelSemester)) {
        return { value: '1' }
    }
    return {
        value: '0',
        message: `ภาคเรียนในฟอร์ม (${formSemester}) ไม่ตรงกับไฟล์ Excel (${excelSemester})`
    }
}

// ตรวจสอบว่าข้อมูลปีการศึกษาในฟอร์ม ตรงกับข้อมูลใน excel หรือไม่
function checkAcademicYear(data: ReportData): CheckResult {
    const formYear = data.formData.academicYear
    const excelYear = data.excelData.data?.home_academic_year
    if (formYear == null || excelYear == null) {
        return { value: '0', message: 'ข้อมูลปีการศึกษาไม่ครบถ้วน' }
    }
    if (String(formYear) === String(excelYear)) {
        return { value: '1' }
    }
    return {
        value: '0',
        message: `ปีการศึกษาในฟอร์ม (${formYear}) ไม่ตรงกับไฟล์ Excel (${excelYear})`
    }
}

// ตรวจสอบว่ามีข้อมูลรายวิชาหรือไม่
function checkSubject(data: ReportData): CheckResult {
    return data.excelData.data?.home_subject ? { value: '1' } : { value: '0', message: 'ไม่มีข้อมูลรายวิชา' }
}

// ตัวอย่าง: ตรวจสอบว่ารหัสวิชาสอดคล้องกับกลุ่มสาระและระดับชั้นหรือไม่
function checkSubjectCode(data: ReportData): CheckResult {
    const code = data.excelData.data?.home_subject_code
    const learningArea = data.excelData.data?.home_learning_area
    const gradeLevel = data.excelData.data?.home_grade_level

    if (!code || typeof code !== 'string' || code.length < 2 || !learningArea || !gradeLevel)
        return { value: '0', message: 'ไม่มีข้อมูลรหัสวิชา/กลุ่มสาระ/ระดับชั้น' }

    const areaMap: Record<string, string> = {
        'ว': 'วิทยาศาสตร์และเทคโนโลยี',
        'ค': 'คณิตศาสตร์',
        'ส': 'สังคมศึกษา ศาสนา และวัฒนธรรม',
        'อ': 'ภาษาต่างประเทศ',
        'ท': 'ภาษาไทย',
        'ศ': 'ศิลปะ',
        'พ': 'สุขศึกษาและพลศึกษา',
        'ง': 'การงานอาชีพ',
    }
    const areaFromCode = areaMap[code[0]]
    if (areaFromCode && areaFromCode !== learningArea)
        return { value: '0', message: 'กลุ่มสาระไม่ตรงกับรหัสวิชา' }

    const levelDigit = code[1]
    if (levelDigit === '2' && !/^ม\.([123])$/.test(gradeLevel))
        return { value: '0', message: 'รหัสวิชาระบุช่วงชั้น ม.1-3 แต่ระดับชั้นไม่ตรง' }
    if (levelDigit === '3' && !/^ม\.([456])$/.test(gradeLevel))
        return { value: '0', message: 'รหัสวิชาระบุช่วงชั้น ม.4-6 แต่ระดับชั้นไม่ตรง' }

    return { value: '1' }
}

// ตรวจสอบว่ามีข้อมูลกลุ่มสาระหรือไม่
function checkLearningArea(data: ReportData): CheckResult {
    return data.excelData.data?.home_learning_area ? { value: '1' } : { value: '0', message: 'ไม่มีข้อมูลกลุ่มสาระ' }
}

// ตรวจสอบว่ามีข้อมูลหน่วยกิตหรือไม่ (ต้องเป็นตัวเลข)
function checkCredit(data: ReportData): CheckResult {
    return typeof data.excelData.data?.home_credit === 'number' ? { value: '1' } : { value: '0', message: 'หน่วยกิตต้องเป็นตัวเลข' }
}

// ตรวจสอบว่าข้อมูลเวลาเรียนเป็น 2 เท่าของหน่วยกิตหรือไม่ (และต้องเป็นตัวเลขทั้งคู่)
function checkStudyTime(data: ReportData): CheckResult {
    const studyTime = data.excelData.data?.home_study_time
    const credit = data.excelData.data?.home_credit
    if (typeof studyTime !== 'number' || typeof credit !== 'number') return { value: '0', message: 'เวลาเรียนและหน่วยกิตต้องเป็นตัวเลข' }
    return studyTime === credit * 2 ? { value: '1' } : { value: '0', message: 'เวลาเรียนต้องเป็น 2 เท่าของหน่วยกิต' }
}

// ตรวจสอบว่ามีข้อมูลครูผู้สอนหรือไม่
function checkTeacher(data: ReportData): CheckResult {
    return data.excelData.data?.home_teacher ? { value: '1' } : { value: '0', message: 'ไม่มีข้อมูลครูผู้สอน' }
}

// ตรวจสอบว่ามีข้อมูลครูที่ปรึกษาหรือไม่
function checkAdvisor(data: ReportData): CheckResult {
    return data.excelData.data?.home_advisor ? { value: '1' } : { value: '0', message: 'ไม่มีข้อมูลครูที่ปรึกษา' }
}

// ตรวจสอบว่าข้อมูล KPA (02_k, 02_p, 02_a, 02_midterm, 02_final) ต้องไม่ว่างและเป็นตัวเลข และ 02_total ต้องเท่ากับ 100
function checkKPA(data: ReportData): CheckResult {
    const d = data.excelData.data

    // ตรวจสอบแต่ละช่องและแยกข้อความ error
    const fieldLabels: Record<string, string> = {
        '02_k': 'K',
        '02_p': 'P',
        '02_a': 'A',
        '02_midterm': 'คะแนนกลางภาค',
        '02_final': 'คะแนนปลายภาค'
    }
    for (const key of Object.keys(fieldLabels)) {
        const val = d?.[key]
        if (val === undefined || val === null || val === '') {
            return { value: '0', message: `ไม่มีข้อมูล ${fieldLabels[key]}` }
        }
        if (isNaN(Number(val))) {
            return { value: '0', message: `${fieldLabels[key]} ต้องเป็นตัวเลข` }
        }
    }

    // ตรวจสอบว่า 02_total ต้องเท่ากับ 100
    const total = Number(d?.['02_total'])
    if (isNaN(total) || total !== 100) {
        return { value: '0', message: 'คะแนนรวมในหน้า 02 ต้องเท่ากับ 100' }
    }

    return { value: '1' }
}

// ตรวจสอบว่าเวลาเรียนรวม (03_total_hour) ต้องเป็น 40 เท่าของหน่วยกิต (home_credit) และทั้งคู่ต้องเป็นตัวเลข
function checkHourCredit(data: ReportData): CheckResult {
    const d = data.excelData.data;
    if (typeof d?.['03_total_hour'] !== 'number' || typeof d?.home_credit !== 'number') {
        return { value: '0', message: 'เวลาเรียนรวมและหน่วยกิตต้องเป็นตัวเลข' };
    }
    if (d['03_total_hour'] !== d.home_credit * 40) {
        return { value: '0', message: `เวลาเรียนรวม (${d['03_total_hour']}) ต้องเป็น 40 เท่าของหน่วยกิต (${d.home_credit})` };
    }
    return { value: '1' }
}

// ตรวจสอบว่ามีข้อมูลคะแนนเต็มก่อนกลางภาค (06_total_before_modterm, 06_total_midterm, 06_total_after_modterm, 06_total_final, 06_total_score) หรือไม่
function checkMidtermScore(data: ReportData): CheckResult {
    const d = data.excelData.data
    const fields = [
        '06_total_before_modterm',
        '06_total_midterm',
        '06_total_after_modterm',
        '06_total_final',
        '06_total_score'
    ]
    const fieldLabels: Record<string, string> = {
        '06_total_before_modterm': 'คะแนนเต็มก่อนกลางภาค',
        '06_total_midterm': 'คะแนนเต็มกลางภาค',
        '06_total_after_modterm': 'คะแนนเต็มหลังกลางภาค',
        '06_total_final': 'คะแนนเต็มปลายภาค',
        '06_total_score': 'คะแนนเต็มรวม'
    }
    const missing: string[] = []
    for (const key of fields) {
        if (!d?.[key]) {
            missing.push(fieldLabels[key])
        }
    }
    if (missing.length > 0) {
        return { value: '0', message: `ไม่มีข้อมูลหรือข้อมูลไม่ถูกต้องในช่อง: ${missing.join(', ')}` }
    }
    return { value: '1' }
}

// ตรวจสอบข้อมูลกลางภาค
export function checkMidtermItems(data: ReportData): CheckResult[] {
    const d = data.excelData.data;

    // 1. บันทึกเวลาเรียน (03)
    const hasStudyRecord = !!d?.['03_check_valid_hour_midterm'];
    const studyRecordResult: CheckResult = hasStudyRecord
        ? { value: '1' }
        : { value: '0', message: 'กรุณาบันทึกเวลาเรียนให้ครบจำนวน 10 สัปดาห์' };

    // 2. คะแนนก่อนกลาง (04)
    const beforeMidtermResult: CheckResult = (() => {
        if (!d?.['06_before_midterm_percent_valid']) {
            return { value: '0', message: 'กรุณากรอกคะแนนก่อนกลางภาคให้ครบถ้วน' };
        }
        if (d?.['06_compare_before_midterm_with_total_before_midterm'] === false) {
            return { value: '0', message: 'คะแนนก่อนกลางภาคมีบางคนกรอกเกินคะแนนเต็ม' };
        }
        return { value: '1' };
    })();

    // 3. คะแนนกลางภาค (04)
    const midtermResult: CheckResult = (() => {
        if (!d?.['06_midterm_percent_valid']) {
            return { value: '0', message: 'กรุณากรอกคะแนนกลางภาคให้ครบถ้วน' };
        }
        if (d?.['06_compare_midterm_with_total_midterm'] === false) {
            return { value: '0', message: 'คะแนนกลางภาคมีบางคนกรอกเกินคะแนนเต็ม' };
        }
        return { value: '1' };
    })();

    return [studyRecordResult, beforeMidtermResult, midtermResult];
}

// ฟังก์ชันคำนวณเปอร์เซ็นต์ความเหมือนของข้อความ (Levenshtein similarity)
function textSimilarity(a: string, b: string): number {
    const alen = a.length;
    const blen = b.length;
    if (alen === 0) return blen === 0 ? 1 : 0;
    if (blen === 0) return 0;

    const dp: number[] = Array(blen + 1).fill(0);
    for (let j = 0; j <= blen; j++) dp[j] = j;

    for (let i = 1; i <= alen; i++) {
        let prev = dp[0];
        dp[0] = i;
        for (let j = 1; j <= blen; j++) {
            const temp = dp[j];
            if (a[i - 1] === b[j - 1]) {
                dp[j] = prev;
            } else {
                dp[j] = Math.min(prev, dp[j - 1], dp[j]) + 1;
            }
            prev = temp;
        }
    }
    const distance = dp[blen];
    return 1 - distance / Math.max(alen, blen);
}

// ตรวจสอบว่าข้อมูลรายวิชาจาก Excel ตรงกับข้อมูลจาก SGS PDF หรือไม่ (>=80% ถือว่าเหมือน)
function checkSgsSubject(data: ReportData): CheckResult {
    if (!data.geminiOcrResult.hasData) {
        return { value: '', message: 'ไม่มีข้อมูล SGS' };
    }
    const excelSubject = data.excelData.data?.home_subject;
    const sgsSubject = data.geminiOcrResult.data?.course_name;
    if (!excelSubject || !sgsSubject) {
        return { value: '0', message: 'ไม่มีข้อมูลชื่อวิชาในไฟล์ใดไฟล์หนึ่ง' };
    }
    const s1 = String(excelSubject).trim();
    const s2 = String(sgsSubject).trim();
    const similarity = textSimilarity(s1, s2);
    if (similarity >= 0.5) {
        return { value: '1' };
    }
    return { value: '0', message: `ชื่อวิชาไม่ตรงกัน (Excel: ${excelSubject}, SGS: ${sgsSubject})` };
}

// ตรวจสอบว่าข้อมูลครูผู้สอนจาก Excel ตรงกับข้อมูลจาก SGS PDF หรือไม่ (>=80% ถือว่าเหมือน)
function checkSgsTeacher(data: ReportData): CheckResult {
    if (!data.geminiOcrResult.hasData) {
        return { value: '', message: 'ไม่มีข้อมูล SGS' };
    }
    const excelTeacher = data.excelData.data?.home_teacher;
    const sgsTeacher = data.geminiOcrResult.data?.teacher;
    if (!excelTeacher || !sgsTeacher) {
        return { value: '0', message: 'ไม่มีข้อมูลครูผู้สอนในไฟล์ใดไฟล์หนึ่ง' };
    }
    // เปรียบเทียบโดยไม่สนใจคำนำหน้าชื่อและช่องว่าง
    const normalize = (name: string) => name.replace(/^(นาย|นาง|นางสาว)\s*/, '').trim();
    const t1 = normalize(String(excelTeacher));
    const t2 = normalize(String(sgsTeacher));
    const similarity = textSimilarity(t1, t2);
    if (similarity >= 0.7) {
        return { value: '1' };
    }
    return { value: '0', message: `ครูผู้สอนไม่ตรงกัน (Excel: ${excelTeacher}, SGS: ${sgsTeacher})` };
}

// ตรวจสอบปีการศึกษาจาก SGS เทียบกับ Form/Excel
function checkSgsAcademicYear(data: ReportData): CheckResult {
    if (!data.geminiOcrResult.hasData) return { value: '', message: 'ไม่มีข้อมูล SGS' }
    const sgsYear = data.geminiOcrResult.data?.academic_year
    const formYear = data.formData.academicYear
    const excelYear = data.excelData.data?.home_academic_year
    if (!sgsYear) return { value: '0', message: 'ไม่มีข้อมูลปีการศึกษาใน SGS' }
    const s = String(sgsYear).trim()
    if (String(formYear).trim() === s || (excelYear && String(excelYear).trim() === s)) return { value: '1' }
    return { value: '0', message: `ปีการศึกษาใน SGS (${s}) ไม่ตรงกับ Form (${formYear})${excelYear ? `, Excel (${excelYear})` : ''}` }
}

// ตรวจสอบภาคเรียนจาก SGS เทียบกับ Form/Excel
function checkSgsSemester(data: ReportData): CheckResult {
    if (!data.geminiOcrResult.hasData) return { value: '', message: 'ไม่มีข้อมูล SGS' }
    const sgsSem = data.geminiOcrResult.data?.semester
    const formSem = data.formData.semester
    const excelSem = data.excelData.data?.home_semester
    if (!sgsSem) return { value: '0', message: 'ไม่มีข้อมูลภาคเรียนใน SGS' }
    const s = String(sgsSem).trim()
    if (String(formSem).trim() === s || (excelSem && String(excelSem).trim() === s)) return { value: '1' }
    return { value: '0', message: `ภาคเรียนใน SGS (${s}) ไม่ตรงกับ Form (${formSem})${excelSem ? `, Excel (${excelSem})` : ''}` }
}

// ตรวจสอบระดับชั้นจาก SGS เทียบกับ Excel
function checkSgsGradeLevel(data: ReportData): CheckResult {
    if (!data.geminiOcrResult.hasData) return { value: '', message: 'ไม่มีข้อมูล SGS' }
    const sgsGrade = data.geminiOcrResult.data?.grade_level
    const excelGrade = data.excelData.data?.home_grade_level
    if (!sgsGrade) return { value: '0', message: 'ไม่มีข้อมูลระดับชั้นใน SGS' }
    const s = String(sgsGrade).trim()
    if (excelGrade && String(excelGrade).trim() === s) return { value: '1' }
    return { value: '0', message: `ระดับชั้นใน SGS (${s}) ไม่ตรงกับ Excel (${excelGrade || 'ไม่มีข้อมูล'})` }
}

// Helper function for creating SGS check results
function createSgsCheckResult(
    geminiOcrResult: ReportData['geminiOcrResult'],
    key: 'attitude_valid' | 'read_analyze_write_valid',
    failMessage: string = 'ไม่ผ่านเกณฑ์'
): CheckResult {
    if (!geminiOcrResult.hasData) {
        return { value: '', message: 'ไม่มีข้อมูล SGS' };
    }
    return geminiOcrResult.data?.[key]
        ? { value: '1' }
        : { value: '0', message: failMessage };
}

// ตรวจสอบข้อมูลปลายภาค
export function checkFinalItems(data: ReportData): CheckResult[] {
    const d = data.excelData.data;

    // 1. บันทึกเวลาเรียน (03)
    const studyRecordResult: CheckResult = d?.['03_check_valid_hour_final']
        ? { value: '1' }
        : { value: '0', message: 'กรุณาบันทึกเวลาเรียนหลังกลางภาคให้ครบ' };

    // 2. คะแนนหลังกลางภาค (05)
    const afterMidtermResult: CheckResult = (() => {
        if (!d?.['06_after_midterm_percent_valid']) {
            return { value: '0', message: 'กรุณากรอกคะแนนหลังกลางภาคให้ครบถ้วน' };
        }
        if (d?.['06_compare_after_midterm_with_total_after_midterm'] === false) {
            return { value: '0', message: 'คะแนนหลังกลางภาคมีบางคนกรอกเกินคะแนนเต็ม' };
        }
        return { value: '1' };
    })();

    // 3. คะแนนสอบปลายภาค (05)
    const finalScoreResult: CheckResult = (() => {
        if (!d?.['06_final_percent_valid']) {
            return { value: '0', message: 'กรุณากรอกคะแนนปลายภาคให้ครบถ้วน' };
        }
        if (d?.['06_compare_final_with_total_final'] === false) {
            return { value: '0', message: 'คะแนนปลายภาคมีบางคนกรอกเกินคะแนนเต็ม' };
        }
        return { value: '1' };
    })();

    // 4. ตรวจสอบการให้ระดับผลการเรียน (06)
    const gradeResult: CheckResult = d?.['06_grade_percent_valid']
        ? { value: '1' }
        : { value: '0', message: 'กรุณาตรวจสอบการให้ระดับผลการเรียน' };

    // 5. คะแนนสมรรถนะ (07)
    const competencyResult: CheckResult = d?.['07_competency_percent_valid']
        ? { value: '1' }
        : { value: '0', message: 'กรุณากรอกคะแนนสมรรถนะให้ครบถ้วน' };

    // 6. คะแนนคุณลักษณะอันพึงประสงค์ (08)
    const attitudeResult: CheckResult = d?.['08_attitude_count_percent_valid']
        ? { value: '1' }
        : { value: '0', message: 'กรุณากรอกคะแนนคุณลักษณะฯ ให้ครบถ้วน' };

    // 7. คะแนนการอ่าน คิดวิเคราะห์และเขียน (09)
    const readAnalyzeWriteResult: CheckResult = d?.['09_read_analyze_write_percent_valid']
        ? { value: '1' }
        : { value: '0', message: 'กรุณากรอกคะแนนการอ่านฯ ให้ครบถ้วน' };

    return [
        studyRecordResult,
        afterMidtermResult,
        finalScoreResult,
        gradeResult,
        competencyResult,
        attitudeResult,
        readAnalyzeWriteResult,
        // การตรวจสอบจาก SGS: ปีการศึกษา, ภาคเรียน, ระดับชั้น
        checkSgsAcademicYear(data),
        checkSgsSemester(data),
        checkSgsGradeLevel(data),
        checkSgsSubject(data),
        checkSgsTeacher(data),
        createSgsCheckResult(data.geminiOcrResult, 'attitude_valid'),
        createSgsCheckResult(data.geminiOcrResult, 'read_analyze_write_valid')
    ];
}