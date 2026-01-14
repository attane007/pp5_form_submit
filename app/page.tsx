'use client'

import { useState, useRef, DragEvent, useEffect } from 'react'
import Image from 'next/image'
import { generatePDF, validateReportData } from '@/lib/pdfGenerator'

export default function Home() {
  const [academicYear, setAcademicYear] = useState('')
  const [semester, setSemester] = useState('')
  const [file, setFile] = useState<File | null>(null)
  const [pdfFile, setPdfFile] = useState<File | null>(null)
  const [isDragOver, setIsDragOver] = useState(false)
  const [isPdfDragOver, setIsPdfDragOver] = useState(false)
  const [error, setError] = useState('')
  const [isLoading, setIsLoading] = useState(false)
  const [isUploading, setIsUploading] = useState(false)
  const [success, setSuccess] = useState(false)
  const [panelContent, setPanelContent] = useState({
    type: 'info' as 'info' | 'error' | 'warning' | 'success',
    title: 'คำแนะนำการใช้งาน',
    messages: [
      'เลือกปีการศึกษาและภาคเรียนที่ต้องการ',
      'อัปโหลดไฟล์ Excel (.xlsx) ที่มีข้อมูล ปพ.5',
      'อัปโหลดรายงาน ปพ.5 จาก SGS (.pdf) - ไม่บังคับ',
      'ไฟล์ต้องมีขนาดไม่เกิน 10MB',
      'กดปุ่ม "ส่งข้อมูลเพื่อตรวจสอบ" เมื่อพร้อม'
    ]
  })
  const fileInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)

  // Generate academic years (10 years back from current year in Buddhist Era)
  const currentYear = new Date().getFullYear()
  const currentBuddhistYear = currentYear + 543
  const academicYears = Array.from({ length: 10 }, (_, i) => currentBuddhistYear - i)

  // Auto-calculate academic year and semester when NEXT_PUBLIC_AUTO_ACADEMIC is set in .env
  const autoAcademic = Boolean(process.env.NEXT_PUBLIC_AUTO_ACADEMIC)

  const computeAutoAcademic = () => {
    const today = new Date()
    const y = today.getFullYear()
    const m = today.getMonth() // 0-based (0=Jan)
    const d = today.getDate()
    let sem = ''
    let acadYear = y + 543

    // Semester 1: May 15 - Oct 31 (May is month 4)
    if ((m === 4 && d >= 15) || (m >= 5 && m <= 9)) {
      sem = '1'
      acadYear = y + 543
    } else if (m === 10 || m === 11 || (m >= 0 && m <= 2) || (m === 2 && d <= 30)) {
      // Semester 2: Nov 1 (10) current year - Mar 30 next year (2)
      sem = '2'
      if (m >= 0 && m <= 2) {
        acadYear = (y - 1) + 543
      } else {
        acadYear = y + 543
      }
    }

    setAcademicYear(String(acadYear))
    setSemester(sem)
  }

  useEffect(() => {
    if (!autoAcademic) return
    computeAutoAcademic()
  }, [autoAcademic])

  const updatePanelContent = (type: 'info' | 'error' | 'warning' | 'success', title: string, messages: string[]) => {
    setPanelContent({ type, title, messages })
  }

  const handleFileChange = async (selectedFile: File | null) => {
    setError('')
    setSuccess(false)
    setIsUploading(false)  // Reset uploading state

    if (!selectedFile) {
      setFile(null)
      return
    }

    // Reset panel to info state when starting new file selection
    if (panelContent.type === 'success') {
      updatePanelContent('info', 'คำแนะนำการใช้งาน', [
        'เลือกปีการศึกษาและภาคเรียนที่ต้องการ',
        'อัปโหลดไฟล์ Excel (.xlsx) ที่มีข้อมูล ปพ.5',
        'อัปโหลดรายงาน ปพ.5 จาก SGS (.pdf) - ไม่บังคับ',
        'ไฟล์ต้องมีขนาดไม่เกิน 10MB',
        'กดปุ่ม "ส่งข้อมูลเพื่อตรวจสอบ" เมื่อพร้อม'
      ])
    }

    setIsUploading(true)

    // Simulate file validation delay for better UX
    await new Promise(resolve => setTimeout(resolve, 500))

    // Check file type
    if (selectedFile.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' &&
      !selectedFile.name.toLowerCase().endsWith('.xlsx')) {
      setError('กรุณาเลือกไฟล์ .xlsx เท่านั้น')
      updatePanelContent('error', 'ข้อผิดพลาดในการเลือกไฟล์', [
        'ไฟล์ที่เลือกไม่ใช่ไฟล์ Excel (.xlsx)',
        'กรุณาเลือกไฟล์ที่มีนามสกุล .xlsx เท่านั้น',
        'ตรวจสอบให้แน่ใจว่าไฟล์เป็น Excel ที่บันทึกในรูปแบบ .xlsx'
      ])
      setFile(null)
      setIsUploading(false)
      if (fileInputRef.current) {
        fileInputRef.current.value = ''
      }
      return
    }

    // Check file size (max 10MB)
    if (selectedFile.size > 10 * 1024 * 1024) {
      setError('ไฟล์มีขนาดใหญ่เกินไป (สูงสุด 10MB)')
      updatePanelContent('error', 'ไฟล์มีขนาดใหญ่เกินไป', [
        `ขนาดไฟล์ปัจจุบัน: ${(selectedFile.size / 1024 / 1024).toFixed(2)} MB`,
        'ขนาดสูงสุดที่อนุญาต: 10 MB',
        'กรุณาลดขนาดไฟล์หรือเลือกไฟล์อื่น'
      ])
      setFile(null)
      setIsUploading(false)
      if (fileInputRef.current) {
        fileInputRef.current.value = ''
      }
      return
    }

    setFile(selectedFile)
    setIsUploading(false)
    setSuccess(true)
    updatePanelContent('success', 'ไฟล์ถูกเลือกเรียบร้อยแล้ว', [
      `ชื่อไฟล์: ${selectedFile.name}`,
      `ขนาดไฟล์: ${(selectedFile.size / 1024 / 1024).toFixed(2)} MB`,
      'ไฟล์พร้อมสำหรับการส่งข้อมูล',
      'กรอกปีการศึกษาและภาคเรียน แล้วกดส่งข้อมูล'
    ])

    // Auto-hide success message
    setTimeout(() => setSuccess(false), 3000)
  }

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsDragOver(true)
  }

  const handleDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsDragOver(false)
  }

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsDragOver(false)

    const droppedFile = e.dataTransfer.files[0]
    handleFileChange(droppedFile)
  }

  const handleBrowseClick = () => {
    fileInputRef.current?.click()
  }

  const handlePdfBrowseClick = () => {
    pdfInputRef.current?.click()
  }

  const handlePdfFileChange = async (selectedFile: File | null) => {
    setError('')
    setSuccess(false)

    if (!selectedFile) {
      setPdfFile(null)
      return
    }

    // Reset panel to info state when starting new file selection
    if (panelContent.type === 'success') {
      updatePanelContent('info', 'คำแนะนำการใช้งาน', [
        'เลือกปีการศึกษาและภาคเรียนที่ต้องการ',
        'อัปโหลดไฟล์ Excel (.xlsx) ที่มีข้อมูล ปพ.5',
        'อัปโหลดรายงาน ปพ.5 จาก SGS (.pdf) - ไม่บังคับ',
        'ไฟล์ต้องมีขนาดไม่เกิน 10MB',
        'กดปุ่ม "ส่งข้อมูลเพื่อตรวจสอบ" เมื่อพร้อม'
      ])
    }

    // Check file type
    if (selectedFile.type !== 'application/pdf' &&
      !selectedFile.name.toLowerCase().endsWith('.pdf')) {
      setError('กรุณาเลือกไฟล์ .pdf เท่านั้น')
      updatePanelContent('error', 'ข้อผิดพลาดในการเลือกไฟล์ PDF', [
        'ไฟล์ที่เลือกไม่ใช่ไฟล์ PDF (.pdf)',
        'กรุณาเลือกไฟล์ที่มีนามสกุล .pdf เท่านั้น',
        'ตรวจสอบให้แน่ใจว่าไฟล์เป็น PDF'
      ])
      setPdfFile(null)
      if (pdfInputRef.current) {
        pdfInputRef.current.value = ''
      }
      return
    }

    // Check file size (max 10MB)
    if (selectedFile.size > 10 * 1024 * 1024) {
      setError('ไฟล์ PDF มีขนาดใหญ่เกินไป (สูงสุด 10MB)')
      updatePanelContent('error', 'ไฟล์ PDF มีขนาดใหญ่เกินไป', [
        `ขนาดไฟล์ปัจจุบัน: ${(selectedFile.size / 1024 / 1024).toFixed(2)} MB`,
        'ขนาดสูงสุดที่อนุญาต: 10 MB',
        'กรุณาลดขนาดไฟล์หรือเลือกไฟล์อื่น'
      ])
      setPdfFile(null)
      if (pdfInputRef.current) {
        pdfInputRef.current.value = ''
      }
      return
    }

    setPdfFile(selectedFile)
    updatePanelContent('success', 'ไฟล์ PDF ถูกเลือกเรียบร้อยแล้ว', [
      `ชื่อไฟล์: ${selectedFile.name}`,
      `ขนาดไฟล์: ${(selectedFile.size / 1024 / 1024).toFixed(2)} MB`,
      'ไฟล์รายงาน SGS พร้อมสำหรับการส่งข้อมูล'
    ])
  }

  const handlePdfDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsPdfDragOver(true)
  }

  const handlePdfDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsPdfDragOver(false)
  }

  const handlePdfDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault()
    setIsPdfDragOver(false)

    const droppedFile = e.dataTransfer.files[0]
    handlePdfFileChange(droppedFile)
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()

    if (!academicYear || !semester || !file) {
      setError('กรุณากรอกข้อมูลให้ครบถ้วน')
      updatePanelContent('error', 'ข้อมูลไม่ครบถ้วน', [
        'กรุณาเลือกปีการศึกษา',
        'กรุณาเลือกภาคเรียน',
        'กรุณาเลือกไฟล์ Excel (.xlsx)',
        'ตรวจสอบข้อมูลและลองอีกครั้ง'
      ])
      return
    }

    setIsLoading(true)
    setError('')

    try {
      // Create FormData
      const formData = new FormData()
      formData.append('academicYear', academicYear)
      formData.append('semester', semester)
      formData.append('file_xlsx', file)

      if (pdfFile) {
        console.log('📄 เพิ่ม PDF file เข้า FormData:', pdfFile.name, `(${(pdfFile.size / 1024 / 1024).toFixed(2)} MB)`)
        formData.append('file_pdf', pdfFile)
      } else {
        console.log('ℹ️  ไม่มี PDF file ที่จะส่ง')
      }

      const backendUrl = '/api/upload'

      updatePanelContent('info', 'กำลังประมวลผล...', [
        'กำลังส่งไฟล์ไปยังเซิร์ฟเวอร์',
        'กำลังอ่านข้อมูลจากไฟล์ Excel',
        pdfFile ? 'กำลังประมวลผล PDF ด้วย Gemini AI' : '',
        'กำลังเตรียมข้อมูลสำหรับรายงาน...'
      ].filter(Boolean))

      // Send POST request to local API
      const response = await fetch(backendUrl, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}))
        throw new Error(errorData.message || `HTTP ${response.status}: ${response.statusText}`)
      }

      // รับข้อมูล JSON จาก backend
      const result = await response.json()
      console.log('Upload successful:', result)

      // ตรวจสอบโครงสร้างข้อมูลและสร้าง PDF
      if (validateReportData(result)) {
        try {
          await generatePDF(result)

          // Reset form on success
          if (autoAcademic) {
            computeAutoAcademic()
          } else {
            setAcademicYear('')
            setSemester('')
          }
          setFile(null)
          setPdfFile(null)
          setSuccess(true)
          setIsUploading(false)  // Reset uploading state

          // Clear file input values
          if (fileInputRef.current) {
            fileInputRef.current.value = ''
          }
          if (pdfInputRef.current) {
            pdfInputRef.current.value = ''
          }

          // Show success message in panel
          updatePanelContent('success', 'ส่งข้อมูลสำเร็จ!', [
            'ข้อมูลถูกส่งไปยังเซิร์ฟเวอร์เรียบร้อยแล้ว',
            'รายงาน PDF ได้ถูกสร้างด้วยฟอนต์ไทยและเปิดในแทบใหม่แล้ว',
            'หากไม่สามารถเปิดแทบใหม่ได้ PDF จะถูกดาวน์โหลดอัตโนมัติ',
            'คุณสามารถส่งไฟล์ใหม่ได้อีกครั้ง'
          ])

          // Reset panel content to initial state after 10 seconds
          setTimeout(() => {
            setSuccess(false)  // Reset success state
            updatePanelContent('info', 'คำแนะนำการใช้งาน', [
              'เลือกปีการศึกษาและภาคเรียนที่ต้องการ',
              'อัปโหลดไฟล์ Excel (.xlsx) ที่มีข้อมูล ปพ.5',
              'อัปโหลดรายงาน ปพ.5 จาก SGS (.pdf) - ไม่บังคับ',
              'ไฟล์ต้องมีขนาดไม่เกิน 10MB',
              'กดปุ่ม "ส่งข้อมูลเพื่อตรวจสอบ" เมื่อพร้อม'
            ])
          }, 10000)
        } catch (pdfError) {
          console.error('ข้อผิดพลาดในการสร้าง PDF:', pdfError)
          throw new Error(pdfError instanceof Error ? pdfError.message : 'ไม่สามารถสร้าง PDF ได้')
        }
      } else {
        console.error('โครงสร้างข้อมูลจาก backend ไม่ถูกต้อง:', result)
        throw new Error('โครงสร้างข้อมูลจากเซิร์ฟเวอร์ไม่ถูกต้อง')
      } // เพิ่มเวลาเป็น 10 วินาที เพื่อให้ผู้ใช้ได้เห็นผลลัพธ์

    } catch (error) {
      console.error('Upload error:', error)

      let errorTitle = 'เกิดข้อผิดพลาดในการส่งข้อมูล'
      let errorMessages: string[] = []

      if (error instanceof TypeError && error.message.includes('Failed to fetch')) {
        errorTitle = 'ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้'
        errorMessages = [
          'ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต',
          'ตรวจสอบว่าเซิร์ฟเวอร์กำลังทำงาน',
          'ติดต่อผู้ดูแลระบบหากปัญหายังคงอยู่'
        ]
      } else if (error instanceof Error) {
        errorMessages = [
          `สาเหตุ: ${error.message}`,
          'ตรวจสอบข้อมูลที่กรอกให้ถูกต้อง',
          'ตรวจสอบขนาดไฟล์ไม่เกิน 10MB',
          'ลองส่งข้อมูลอีกครั้งในภายหลัง'
        ]
      } else {
        errorMessages = [
          'เกิดข้อผิดพลาดที่ไม่ทราบสาเหตุ',
          'ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต',
          'ลองส่งข้อมูลอีกครั้งในภายหลัง',
          'ติดต่อผู้ดูแลระบบหากปัญหายังคงอยู่'
        ]
      }

      const errorMessage = error instanceof Error ? error.message : 'เกิดข้อผิดพลาดในการส่งข้อมูล'
      setError(`ข้อผิดพลาด: ${errorMessage}`)
      updatePanelContent('error', errorTitle, errorMessages)
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-blue-50 py-5 px-4 sm:px-6 lg:px-8"> {/* Changed background color */}
      <div className="max-w-5xl mx-auto">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 lg:gap-8 lg:items-center lg:min-h-[calc(100vh-5rem)]">
          {/* Main Form Section */}
          <div className="flex justify-center">
            <div className="max-w-lg w-full">
              {/* Header */}
              <div className="text-center mb-7">
                <div className="flex justify-center mb-3">
                  <div className="p-1.5 bg-white rounded-xl shadow-md">
                    <Image
                      src="/logo-ppk-512x512-1.png"
                      alt="PPK Logo"
                      width={50}
                      height={50}
                      priority
                      className="rounded-lg"
                    />
                  </div>
                </div>
                <h1 className="text-3xl font-bold text-indigo-800 mb-2"> {/* Changed text color */}
                  ส่งไฟล์ ปพ.5
                </h1>
                <p className="text-indigo-600">อัปโหลดไฟล์เพื่อตรวจสอบข้อมูล</p> {/* Changed text color */}
                <div className="w-16 h-1 bg-blue-600 mx-auto mt-2 rounded-full"> {/* Changed background color */}
                </div>
              </div>

              {/* Form Card */}
              <div className="bg-white rounded-xl shadow-lg border border-slate-200 p-5">
                <form onSubmit={handleSubmit} className="space-y-5">
                  {/* Academic Year Selection */}
                  <div className="space-y-2">
                    <label htmlFor="academicYear" className="block text-sm font-semibold text-slate-700 mb-1.5">
                      ปีการศึกษา <span className="text-red-500">*</span>
                    </label>
                    <div className="relative">
                      <select
                        id="academicYear"
                        value={academicYear}
                        onChange={(e) => {
                          setAcademicYear(e.target.value)
                          setError('')
                          setSuccess(false)
                        }}
                        className={`w-full px-4 py-2.5 bg-slate-50 border-2 border-slate-200 rounded-lg focus:outline-none focus:border-blue-500 text-slate-700 font-medium appearance-none ${autoAcademic ? 'cursor-not-allowed opacity-70' : 'cursor-pointer'}`}
                        required
                        disabled={autoAcademic}
                      >
                        <option value="">เลือกปีการศึกษา</option>
                        {academicYears.map((year) => (
                          <option key={year} value={year}>
                            {year}
                          </option>
                        ))}
                      </select>
                      <div className="absolute inset-y-0 right-0 flex items-center pr-4 pointer-events-none">
                        <svg className="h-5 w-5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                        </svg>
                      </div>
                    </div>
                  </div>

                  {/* Semester Selection */}
                  <div className="space-y-2">
                    <label htmlFor="semester" className="block text-sm font-semibold text-slate-700 mb-1.5">
                      ภาคเรียน <span className="text-red-500">*</span>
                    </label>
                    <div className="relative">
                      <select
                        id="semester"
                        value={semester}
                        onChange={(e) => {
                          setSemester(e.target.value)
                          setError('')
                          setSuccess(false)
                        }}
                        className={`w-full px-4 py-2.5 bg-slate-50 border-2 border-slate-200 rounded-lg focus:outline-none focus:border-blue-500 text-slate-700 font-medium appearance-none ${autoAcademic ? 'cursor-not-allowed opacity-70' : 'cursor-pointer'}`}
                        required
                        disabled={autoAcademic}
                      >
                        <option value="">เลือกภาคเรียน</option>
                        <option value="1">ภาคเรียนที่ 1</option>
                        <option value="2">ภาคเรียนที่ 2</option>
                        <option value="3">ภาคเรียนที่ 3</option>
                      </select>
                      <div className="absolute inset-y-0 right-0 flex items-center pr-4 pointer-events-none">
                        <svg className="h-5 w-5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                        </svg>
                      </div>
                    </div>
                  </div>

                  {/* File Upload Section - Side by Side */}
                  <div className="space-y-4">
                    <div className="text-sm font-semibold text-slate-700">
                      อัปโหลดไฟล์เอกสาร
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Excel File Upload */}
                      <div className="space-y-2">
                        <label className="block text-sm font-medium text-slate-700 mb-1.5">
                          ไฟล์ ปพ.5 (.xlsx) <span className="text-red-500">*</span>
                        </label>
                        <div
                          className={`border-2 border-dashed rounded-lg p-4 text-center cursor-pointer ${isDragOver
                            ? 'border-blue-400 bg-blue-50'
                            : file
                              ? 'border-green-300 bg-green-50'
                              : 'border-slate-300 bg-slate-50'
                            }`}
                          onDragOver={handleDragOver}
                          onDragLeave={handleDragLeave}
                          onDrop={handleDrop}
                          onClick={() => !file && !isUploading && handleBrowseClick()}
                        >
                          <input
                            type="file"
                            ref={fileInputRef}
                            onChange={(e) => handleFileChange(e.target.files?.[0] || null)}
                            accept=".xlsx"
                            className="hidden"
                          />

                          {isUploading ? (
                            <div className="text-blue-600">
                              <div className="inline-flex items-center justify-center w-10 h-10 bg-blue-100 rounded-full mb-2">
                                <svg className="animate-spin h-4 w-4" fill="none" viewBox="0 0 24 24">
                                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                                </svg>
                              </div>
                              <p className="font-medium text-sm">กำลังตรวจสอบ...</p>
                            </div>
                          ) : file ? (
                            <div className="text-green-700">
                              <div className="inline-flex items-center justify-center w-10 h-10 bg-green-100 rounded-full mb-2">
                                <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                                </svg>
                              </div>
                              <p className="font-medium text-sm mb-1 break-words">{file.name}</p>
                              <p className="text-xs text-slate-500 mb-2">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation()
                                  setFile(null)
                                  setError('')
                                  setSuccess(false)
                                  if (fileInputRef.current) {
                                    fileInputRef.current.value = ''
                                  }
                                }}
                                className="inline-flex items-center px-2 py-1 text-xs font-medium text-red-600 bg-red-50 rounded hover:bg-red-100"
                              >
                                <svg className="w-3 h-3 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                                </svg>
                                ลบ
                              </button>
                            </div>
                          ) : (
                            <div className="text-slate-600">
                              <div className="inline-flex items-center justify-center w-10 h-10 bg-slate-100 rounded-full mb-2">
                                <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10" />
                                </svg>
                              </div>
                              <p className="font-medium text-sm mb-1">ลากไฟล์ Excel</p>
                              <p className="text-xs mb-2">หรือคลิกเพื่อเลือก</p>
                              <div className="inline-flex items-center px-3 py-1.5 text-xs text-blue-600 bg-blue-50 rounded font-medium">
                                <svg className="w-3 h-3 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8l-8-8-8 8" />
                                </svg>
                                เลือกไฟล์
                              </div>
                              <p className="text-xs text-slate-400 mt-2">รองรับ .xlsx (สูงสุด 10MB)</p>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* PDF File Upload */}
                      <div className="space-y-2">
                        <label className="block text-sm font-medium text-slate-700 mb-1.5">
                          รายงาน SGS (.pdf)
                        </label>
                        <div
                          className={`border-2 border-dashed rounded-lg p-4 text-center cursor-pointer ${isPdfDragOver
                            ? 'border-purple-400 bg-purple-50'
                            : pdfFile
                              ? 'border-green-300 bg-green-50'
                              : 'border-slate-300 bg-slate-50'
                            }`}
                          onDragOver={handlePdfDragOver}
                          onDragLeave={handlePdfDragLeave}
                          onDrop={handlePdfDrop}
                          onClick={() => !pdfFile && handlePdfBrowseClick()}
                        >
                          <input
                            type="file"
                            ref={pdfInputRef}
                            onChange={(e) => handlePdfFileChange(e.target.files?.[0] || null)}
                            accept=".pdf"
                            className="hidden"
                          />

                          {pdfFile ? (
                            <div className="text-green-700">
                              <div className="inline-flex items-center justify-center w-10 h-10 bg-green-100 rounded-full mb-2">
                                <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                                </svg>
                              </div>
                              <p className="font-medium text-sm mb-1 break-words">{pdfFile.name}</p>
                              <p className="text-xs text-slate-500 mb-2">{(pdfFile.size / 1024 / 1024).toFixed(2)} MB</p>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation()
                                  setPdfFile(null)
                                  setError('')
                                  setSuccess(false)
                                  if (pdfInputRef.current) {
                                    pdfInputRef.current.value = ''
                                  }
                                }}
                                className="inline-flex items-center px-2 py-1 text-xs font-medium text-red-600 bg-red-50 rounded hover:bg-red-100"
                              >
                                <svg className="w-3 h-3 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                                </svg>
                                ลบ
                              </button>
                            </div>
                          ) : (
                            <div className="text-slate-600">
                              <div className="inline-flex items-center justify-center w-10 h-10 bg-slate-100 rounded-full mb-2">
                                <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
                                </svg>
                              </div>
                              <p className="font-medium text-sm mb-1">ลากไฟล์ PDF</p>
                              <p className="text-xs mb-2">หรือคลิกเพื่อเลือก</p>
                              <div className="inline-flex items-center px-3 py-1.5 text-xs text-purple-600 bg-purple-50 rounded font-medium">
                                <svg className="w-3 h-3 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8l-8-8-8 8" />
                                </svg>
                                เลือกไฟล์
                              </div>
                              <p className="text-xs text-slate-400 mt-2">รองรับ .pdf (สูงสุด 10MB)</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Success Message */}
                  {success && (
                    <div className="flex items-center space-x-3 text-green-700 bg-green-50 border border-green-200 rounded-lg p-2.5">
                      <svg className="h-5 w-5 text-green-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                      <span className="font-medium">ส่งข้อมูลเรียบร้อยแล้ว!</span>
                    </div>
                  )}

                  {/* Error Message */}
                  {error && (
                    <div className="flex items-center space-x-3 text-red-700 bg-red-50 border border-red-200 rounded-lg p-2.5">
                      <svg className="h-5 w-5 text-red-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                      <span className="font-medium">{error}</span>
                    </div>
                  )}

                  {/* Submit Button */}
                  <button
                    type="submit"
                    disabled={isLoading || isUploading}
                    className={`w-full py-2.5 px-6 rounded-lg font-semibold ${isLoading || isUploading
                      ? 'bg-slate-400 text-white cursor-not-allowed'
                      : 'bg-blue-600 text-white hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2'
                      }`}
                  >
                    {isLoading ? (
                      <div className="flex items-center justify-center space-x-2">
                        <svg className="animate-spin h-5 w-5" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                        </svg>
                        <span>กำลังส่งข้อมูล...</span>
                      </div>
                    ) : (
                      <div className="flex items-center justify-center space-x-2">
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
                        </svg>
                        <span>ส่งข้อมูลเพื่อตรวจสอบ</span>
                      </div>
                    )}
                  </button>
                </form>
              </div>

              {/* Footer */}
              <div className="text-center mt-5 text-slate-500 text-sm">
                <p>โรงเรียนโพนงามพิทยานุกูล</p>
              </div>
            </div>
          </div>

          {/* Side Panel */}
          <div className="flex justify-center">
            <div className="w-full max-w-lg">
              {/* Panel Header */}
              <div className="mb-4">
                <h2 className="text-lg font-semibold text-slate-800">ข้อมูลและคำแนะนำ</h2>
              </div>

              {/* Panel Content */}
              <div className={`bg-white rounded-xl shadow-lg border-2 p-5 ${panelContent.type === 'error' ? 'border-red-200 bg-red-50' :
                panelContent.type === 'warning' ? 'border-yellow-200 bg-yellow-50' :
                  panelContent.type === 'success' ? 'border-green-200 bg-green-50' :
                    'border-blue-200 bg-blue-50'
                }`}>
                {/* Panel Icon and Title */}
                <div className="flex items-center mb-4">
                  <div className={`flex-shrink-0 w-8 h-8 rounded-full flex items-center justify-center ${panelContent.type === 'error' ? 'bg-red-100' :
                    panelContent.type === 'warning' ? 'bg-yellow-100' :
                      panelContent.type === 'success' ? 'bg-green-100' :
                        'bg-blue-100'
                    }`}>
                    {panelContent.type === 'error' && (
                      <svg className="w-5 h-5 text-red-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                    )}
                    {panelContent.type === 'warning' && (
                      <svg className="w-5 h-5 text-yellow-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L3.732 16.5c-.77.833.192 2.5 1.732 2.5z" />
                      </svg>
                    )}
                    {panelContent.type === 'success' && (
                      <svg className="w-5 h-5 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                    )}
                    {panelContent.type === 'info' && (
                      <svg className="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                      </svg>
                    )}
                  </div>
                  <h3 className={`ml-3 font-semibold ${panelContent.type === 'error' ? 'text-red-800' :
                    panelContent.type === 'warning' ? 'text-yellow-800' :
                      panelContent.type === 'success' ? 'text-green-800' :
                        'text-blue-800'
                    }`}>
                    {panelContent.title}
                  </h3>
                </div>

                {/* Panel Messages */}
                <div className="space-y-2">
                  {panelContent.messages.map((message, index) => (
                    <div key={index} className={`flex items-start space-x-2 text-sm ${panelContent.type === 'error' ? 'text-red-700' :
                      panelContent.type === 'warning' ? 'text-yellow-700' :
                        panelContent.type === 'success' ? 'text-green-700' :
                          'text-blue-700'
                      }`}>
                      <div className="flex-shrink-0 w-1.5 h-1.5 rounded-full bg-current mt-2"></div>
                      <span>{message}</span>
                    </div>
                  ))}
                </div>

                {/* Quick Actions */}
                {panelContent.type === 'info' && (
                  <div className="mt-4 pt-4 border-t border-blue-200">
                    <div className="text-xs text-blue-600 font-medium mb-2">การตรวจสอบไฟล์</div>
                    <div className="space-y-1 text-xs text-blue-700">
                      <div>✓ ไฟล์ Excel: .xlsx เท่านั้น</div>
                      <div>✓ ไฟล์ PDF: .pdf เท่านั้น</div>
                      <div>✓ ขนาดไฟล์: สูงสุด 10 MB</div>
                      <div>✓ เนื้อหา: ข้อมูล ปพ.5 และรายงาน SGS</div>
                    </div>
                  </div>
                )}

                {/* ลิงก์ไปยังหน้าตรวจสอบรายงาน */}
                <div className="mt-4 pt-4 border-t border-blue-200">
                  <div className="text-xs text-blue-600 font-medium mb-2">ตรวจสอบรายงาน</div>
                  <a
                    href="/verify"
                    className="inline-flex items-center px-3 py-2 bg-purple-600 text-white rounded-lg text-sm font-medium hover:bg-purple-700 transition"
                  >
                    <svg className="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4m-4-2l4-4m0 0l-4-4m4 4H9" />
                    </svg>
                    ไปที่หน้าตรวจสอบรายงาน
                  </a>
                </div>

                {/* ตัวอย่างไฟล์ Excel */}
                <div className="mt-6">
                  <a
                    href="/ปพ.5_1_2568_ตัวอย่าง.xlsx"
                    download
                    className="inline-flex items-center px-3 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 transition"
                  >
                    {/* ไอคอนดาวน์โหลด (ลูกศรลง + ก้อนเมฆ) */}
                    <svg className="w-5 h-5 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M9 19l3 3m0 0l3-3m-3 3V10" />
                    </svg>
                    ดาวน์โหลดไฟล์ Excel ตัวอย่าง
                  </a>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
