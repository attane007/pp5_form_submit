# ระบบส่งไฟล์ ปพ.5

ระบบอัปโหลดไฟล์ Excel และ PDF สำหรับข้อมูล ปพ.5 และรายงาน SGS

## 🚀 การติดตั้งและรัน

1. **ติดตั้ง dependencies:**

   ```bash
   npm install
   ```

2. **ตั้งค่า Environment Variables:**

   - สร้างไฟล์ `.env` หรือคัดลอกจาก `.env.example` ที่มีตัวอย่างค่าตัวแปรที่จำเป็น:

   ```bash
   cp .env.example .env
   ```

   - แก้ไขค่าใน `.env` ให้ตรงกับ environment ของคุณ เช่น

   ```env
   DATABASE_URL="mongodb+srv://<user>:<password>@<host>/<database>"
   GEMINI_API_KEY="<your-gemini-api-key>"
   NEXT_PUBLIC_BASE_URL=http://localhost:3000
   PORT=3000
   ```

3. **รันโปรเจ็กต์ (โหมดพัฒนา):**

   ```bash
   npm run dev
   ```

4. **Deploy ด้วย Docker Compose:**
   - ตรวจสอบว่าได้ตั้งค่าไฟล์ `.env` แล้ว
   - รันคำสั่ง:
   ```bash
   docker compose up --build
   ```
   - แอปจะรันที่พอร์ทตามที่กำหนดใน `.env` เช่น `http://localhost:3000`

## ⚙️ การตั้งค่า Backend URL

แก้ไขไฟล์ `.env` และเปลี่ยน `NEXT_PUBLIC_BASE_URL` เป็น URL ของ backend server:

- **Development:** `http://localhost:3000`
- **Production:** `https://your-backend-domain.com`

## 🛠️ การแก้ไขปัญหา

### ❌ TypeError: Failed to fetch

**สาเหตุที่เป็นไปได้:**

1. ไม่มีการตั้งค่า `NEXT_PUBLIC_BASE_URL`
2. Backend server ไม่ได้รัน
3. URL ไม่ถูกต้อง
4. ปัญหา CORS

**วิธีแก้ไข:**

1. **ตรวจสอบไฟล์ `.env`:**

   ```env
   NEXT_PUBLIC_BASE_URL=http://localhost:3000
   ```

2. **ทดสอบการเชื่อมต่อ:**

   - คลิกปุ่ม "🔍 ทดสอบการเชื่อมต่อเซิร์ฟเวอร์" ใน panel

3. **ตรวจสอบ Backend Server:**

   ```bash
   # ตรวจสอบว่า server รันที่ port ที่ตั้งไว้ใน .env (เช่น 3000)
   curl http://localhost:3000/api/health
   ```

4. **ตรวจสอบ Console Log:**
   - เปิด Developer Tools (F12)
   - ดูใน Console tab สำหรับข้อผิดพลาด

### 🌐 การตั้งค่า CORS (สำหรับ Backend)

หาก backend เป็น Express.js:

```javascript
const cors = require("cors");
app.use(
  cors({
    origin: "http://localhost:3000", // URL ของ frontend
    credentials: true,
  })
);
```

## 📋 ฟีเจอร์

- ✅ อัปโหลดไฟล์ Excel (.xlsx)
- ✅ อัปโหลดไฟล์ PDF รายงาน SGS
- ✅ Drag & Drop
- ✅ การตรวจสอบไฟล์
- ✅ ทดสอบการเชื่อมต่อ
- ✅ แสดงข้อผิดพลาดแบบละเอียด
- ✅ Responsive Design
- ✅ รองรับการ deploy ด้วย Docker Compose
- ✅ ตั้งค่า environment variables ผ่านไฟล์ .env

## 🏗️ โครงสร้างโปรเจ็กต์

```
pp5_form_submit/
├── app/
│   ├── page.tsx          # หน้าหลัก
│   ├── layout.tsx        # Layout
│   └── globals.css       # Styles
├── public/               # Static files
├── .env                  # Environment variables
├── .env.example          # ตัวอย่าง environment variables
├── Dockerfile            # สำหรับ build Docker image
├── docker-compose.yml    # สำหรับ orchestrate ด้วย Docker Compose
└── package.json          # Dependencies
```

## 📚 Technologies

- **Frontend:** Next.js 14, React, TypeScript, Tailwind CSS
- **File Upload:** FormData API
- **State Management:** React Hooks
