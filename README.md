# ระบบส่งไฟล์ ปพ.5 และตรวจสอบรายงาน

ระบบอัปโหลดไฟล์ Excel และ PDF สำหรับข้อมูล ปพ.5 พร้อมการตรวจสอบรายงานด้วย AI และระบบตรวจสอบสถานะผ่าน QR Code

## 🚀 ฟีเจอร์หลัก

- ✅ **อัปโหลดไฟล์ Excel (.xlsx)** - สำหรับข้อมูล ปพ.5
- ✅ **อัปโหลดไฟล์ PDF รายงาน SGS** - สำหรับการตรวจสอบด้วย AI
- ✅ **การประมวลผลด้วย AI (Gemini)** - วิเคราะห์รายงาน PDF อัตโนมัติ
- ✅ **ระบบตรวจสอบรายงาน** - ตรวจสอบความถูกต้องของข้อมูล
- ✅ **QR Code สำหรับติดตาม** - ตรวจสอบสถานะการประมวลผล
- ✅ **Drag & Drop Interface** - อัปโหลดไฟล์แบบลากวาง
- ✅ **Responsive Design** - รองรับทุกอุปกรณ์
- ✅ **ฐานข้อมูล MongoDB** - เก็บข้อมูลการส่งและผลการประมวลผล

## 🛠️ เทคโนโลยีที่ใช้

- **Frontend:** Next.js 15, React 19, TypeScript, Tailwind CSS
- **Backend:** Next.js API Routes
- **Database:** MongoDB (ผ่าน Prisma ORM)
- **AI Processing:** Google Gemini AI
- **File Processing:** XLSX, jsPDF, QRCode
- **Deployment:** Docker & Docker Compose

## 📋 การติดตั้งและรัน

### 1. ติดตั้ง Dependencies

```bash
npm install
```

### 2. ตั้งค่า Environment Variables

สร้างไฟล์ `.env` ในโฟลเดอร์หลัก:

```bash
# คัดลอกไฟล์ตัวอย่าง
cp .env.example .env
```

แก้ไขไฟล์ `.env` ให้ตรงกับ environment ของคุณ:

```env
# ========================================
# Environment Variables สำหรับระบบ ปพ.5
# ========================================

# Database Configuration
DATABASE_URL="mongodb+srv://<username>:<password>@<cluster>/<database>?retryWrites=true&w=majority"

# AI Processing Configuration
GEMINI_API_KEY="<your-gemini-api-key>"

# Application Configuration
NEXT_PUBLIC_BASE_URL="http://localhost:3000"
PORT=3000

# ========================================
# Optional Configuration
# ========================================

# Signature Configuration (สำหรับ PDF)
NEXT_PUBLIC_SIGNATURE_REGISTRAR_NAME="นายทะเบียน (ยังไม่ได้ตั้งค่า)"
NEXT_PUBLIC_SIGNATURE_ACADEMIC_HEAD_NAME="หัวหน้าวิชาการ (ยังไม่ได้ตั้งค่า)"
NEXT_PUBLIC_SIGNATURE_DIRECTOR_NAME="ผู้อำนวยการ (ยังไม่ได้ตั้งค่า)"

# Docker Configuration (ถ้าใช้ Docker)
NEXT_PUBLIC_SITE_URL="http://localhost:3000"

# Node Environment
NODE_ENV="development"
```

### คำแนะนำการตั้งค่า Environment Variables

#### 1. **DATABASE_URL**
- สำหรับ MongoDB Atlas: `mongodb+srv://<username>:<password>@<cluster>/<database>?retryWrites=true&w=majority`
- สำหรับ MongoDB Local: `mongodb://localhost:27017/<database>`
- ตรวจสอบให้แน่ใจว่า username และ password ถูกต้อง

#### 2. **GEMINI_API_KEY**
- รับ API key จาก [Google AI Studio](https://makersuite.google.com/app/apikey)
- ตรวจสอบสิทธิ์การใช้งาน Gemini API
- API key ต้องไม่หมดอายุ

#### 3. **NEXT_PUBLIC_BASE_URL**
- Development: `http://localhost:3000`
- Production: `https://your-domain.com`
- ต้องไม่มี trailing slash

#### 4. **Signature Configuration**
- ตั้งค่าชื่อผู้ลงนามใน PDF
- ใช้สำหรับการสร้างรายงาน PDF
- สามารถปล่อยเป็นค่า default ได้

### 3. ตั้งค่าฐานข้อมูล

```bash
# Generate Prisma client
npx prisma generate

# Push schema to database
npx prisma db push
```

### 4. รันโปรเจค (Development Mode)

```bash
npm run dev
```

### 5. Deploy ด้วย Docker

```bash
# Build และรันด้วย Docker Compose
docker compose up --build
```

## 📁 โครงสร้างโปรเจค

```
pp5_form_submit/
├── app/
│   ├── api/
│   │   ├── upload/route.ts      # API สำหรับอัปโหลดไฟล์
│   │   └── verify/route.ts      # API สำหรับตรวจสอบสถานะ
│   ├── verify/
│   │   ├── page.tsx             # หน้าตรวจสอบสถานะ
│   │   └── VerifyClient.tsx     # Component ตรวจสอบ
│   ├── page.tsx                 # หน้าหลักอัปโหลด
│   ├── layout.tsx               # Layout หลัก
│   └── globals.css              # Styles
├── lib/
│   ├── pdfGenerator.ts          # ฟังก์ชันสร้าง PDF
│   ├── pdfTypes.ts              # Type definitions
│   ├── pdfUtils.ts              # Utility functions
│   ├── prisma.ts                # Prisma client
│   └── reportCheckers.ts        # ฟังก์ชันตรวจสอบรายงาน
├── prisma/
│   └── schema.prisma            # Database schema
├── public/                      # Static files
├── Dockerfile                   # Docker configuration
├── docker-compose.yml           # Docker Compose
└── package.json                 # Dependencies
```

## 🔧 การใช้งาน

### การอัปโหลดไฟล์

1. **เลือกปีการศึกษาและภาคเรียน**
2. **อัปโหลดไฟล์ Excel (.xlsx)** ที่มีข้อมูล ปพ.5
3. **อัปโหลดไฟล์ PDF รายงาน SGS** (ไม่บังคับ)
4. **กดปุ่ม "ส่งข้อมูลเพื่อตรวจสอบ"**

### การตรวจสอบสถานะ

1. **สแกน QR Code** ที่ได้รับหลังการส่ง
2. **หรือกรอก UUID** ในหน้า `/verify`
3. **ดูผลการตรวจสอบ** และสถานะการประมวลผล

## 🔍 การตรวจสอบรายงาน

ระบบจะตรวจสอบรายงานใน 3 ช่วง:

### ก่อนกลางภาค
- ข้อมูลระดับชั้น, ห้องเรียน, ภาคเรียน, ปีการศึกษา
- ข้อมูลรายวิชา, รหัสวิชา, กลุ่มสาระ
- หน่วยกิต, เวลาเรียน, ครูผู้สอน
- ความถูกต้องของ KPA, คะแนนเต็ม

### กลางภาค
- บันทึกเวลาเรียน
- คะแนนก่อนกลางภาค
- คะแนนกลางภาค

### ปลายภาค
- คะแนนหลังกลางภาค, คะแนนสอบปลายภาค
- การให้ระดับผลการเรียน
- คะแนนสมรรถนะ, คุณลักษณะอันพึงประสงค์
- คะแนนการอ่าน คิดวิเคราะห์และเขียน
- ความตรงกันกับข้อมูล SGS

## 🚨 การแก้ไขปัญหา

### ข้อผิดพลาดทั่วไป

1. **DATABASE_URL ไม่ถูกต้อง**
   - ตรวจสอบการเชื่อมต่อ MongoDB
   - ตรวจสอบ username, password, และ database name
   - ตรวจสอบ format ของ connection string

2. **GEMINI_API_KEY ไม่ถูกต้อง**
   - ตรวจสอบ API key จาก Google AI Studio
   - ตรวจสอบสิทธิ์การใช้งาน
   - ตรวจสอบว่า API key ยังไม่หมดอายุ

3. **NEXT_PUBLIC_BASE_URL ไม่ถูกต้อง**
   - ตรวจสอบ URL ของแอปพลิเคชัน
   - ตรวจสอบ protocol (http/https)
   - ตรวจสอบ port number

4. **ไฟล์มีขนาดใหญ่เกินไป**
   - ขนาดไฟล์สูงสุด: 10MB
   - ลดขนาดไฟล์หรือบีบอัด

5. **Signature Configuration ไม่ถูกต้อง**
   - ตรวจสอบชื่อใน NEXT_PUBLIC_SIGNATURE_* variables
   - ตรวจสอบการตั้งค่าสำหรับ PDF generation

### การ Debug

```bash
# ตรวจสอบ logs
docker compose logs

# ตรวจสอบ database
npx prisma studio

# ตรวจสอบ environment variables
echo $DATABASE_URL
echo $GEMINI_API_KEY
echo $NEXT_PUBLIC_BASE_URL
echo $NEXT_PUBLIC_SIGNATURE_REGISTRAR_NAME

# ตรวจสอบ Prisma connection
npx prisma db pull

# ตรวจสอบ API endpoints
curl http://localhost:3000/api/health
```

## 📝 การพัฒนา

### Scripts ที่มี

```bash
npm run dev          # รันในโหมด development
npm run build        # Build สำหรับ production
npm run start        # รัน production server
npm run lint         # ตรวจสอบ code quality
```

### การเพิ่มฟีเจอร์ใหม่

1. **เพิ่ม API Route** ใน `app/api/`
2. **อัปเดต Database Schema** ใน `prisma/schema.prisma`
3. **เพิ่ม Type Definitions** ใน `lib/`
4. **อัปเดต UI Components** ใน `app/`

## 🤝 การสนับสนุน

หากพบปัญหาในการใช้งาน กรุณาติดต่อทีมพัฒนา หรือสร้าง Issue ใน repository

## 📄 License

โปรเจคนี้เป็นส่วนหนึ่งของระบบจัดการข้อมูลการศึกษา
