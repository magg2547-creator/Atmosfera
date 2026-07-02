# Atmosfera

> แดชบอร์ดแสดงผลคุณภาพอากาศและการใช้พลังงานแบบ Real-time — เร็ว ลื่นไหล ไร้รอยต่อ

![License](https://img.shields.io/badge/license-Private-blue)
![Stack](https://img.shields.io/badge/stack-Vanilla%20JS%20%7C%20CSS3%20%7C%20Vercel-lightgrey)
![Status](https://img.shields.io/badge/status-Production%20Ready-green)

---

## ภาพรวม

Atmosfera คือ Web Application ประเภท **Display-Only Dashboard** ที่แสดงผลข้อมูลคุณภาพอากาศ (Air Quality) และการใช้พลังงานไฟฟ้า (Energy Consumption) ดึงข้อมูลจาก Google Sheets ผ่าน Google Apps Script โดยเน้นประสบการณ์การแสดงผลที่:

- **Reactive** — UI อัปเดตอัตโนมัติเมื่อข้อมูลเปลี่ยนแปลง ไม่ต้องสั่ง Render เอง
- **Fluid** — เลื่อนหน้าจอลื่นไหลตามธรรมชาติ รองรับทั้ง Mouse Wheel และ Trackpad
- **Resilient** — ป้องกัน Memory Leak, Layout Shift และ Scroll Hijacking ตั้งแต่ระดับสถาปัตยกรรม

---

## สถาปัตยกรรมภายใน

### Frontend Stack

| Layer | เทคโนโลยี |
|---|---|
| Markup | HTML5 (Semantic) |
| Styling | CSS3 — Custom Properties, `backdrop-filter`, `color-mix()`, CSS Grid/Flexbox |
| Logic | Vanilla JavaScript ES6+ Modules (No Frameworks) |
| Charts | Chart.js 4.x |

### Reactive State Store

หัวใจของ Atmosfera คือ **Reactive State Store** (`src/store/reactiveState.js`) ที่เปลี่ยนระบบจาก Imperative Render มาเป็น **Data-Driven Observer Pattern**:

```
┌─────────────┐     setState()      ┌───────────────┐
│  Data Layer │ ──────────────────► │  State Store  │
│  (Fetcher)  │                     │  (Observer)   │
└─────────────┘                     └───────┬───────┘
                                            │ queueMicrotask
                                            ▼
                                   ┌─────────────────┐
                                   │   Subscribers   │
                                   │  (UI Bindings)  │
                                   └─────────────────┘
```

- **`setState(updater)`** — อัปเดตเฉพาะ Path ที่เปลี่ยน ใช้ Path-targeted comparison แทนการ Clone ทั้ง Tree
- **`subscribe(path, cb)`** — ลงทะเบียน Listener ที่ทำงานเฉพาะเมื่อ Path ที่สนใจเปลี่ยนแปลง
- **`batch(fn)`** — รวมการอัปเดตหลายครั้งเป็นการ Notify ครั้งเดียว ป้องกัน Re-render ซ้ำซ้อน
- **`queueMicrotask`** — ใช้ Microtask Queue แทน `setTimeout` เพื่อความแม่นยำสูงสุด
- **`structuredClone`** — ใช้ Native API ของ Browser แทน Custom Clone เพื่อประสิทธิภาพที่ดีขึ้น

### Modularization — Separation of Concerns

| โมดูล | หน้าที่ |
|---|---|
| `src/config.js` | Shared configuration — Single source of truth สำหรับค่าคงที่ทั้งโปรเจกต์ |
| `src/utils/aqiHelper.js` | ถอด Logic การคำนวณเกณฑ์ค่า AQI และ Meta ข้อมูลออกจากไฟล์หลัก |
| `src/utils/dateHelper.js` | จัดการรูปแบบวันที่และเวลาสำหรับแสดงผล |
| `src/utils/escapeHtml.js` | ป้องกัน XSS ด้วย HTML Escaping |
| `src/charts/dashboardCharts.js` | จัดการ Chart.js instances, อัปเดตกราฟด้วยแอนิเมชันนุ่มนวล |
| `src/charts/plugins.js` | Chart.js custom plugins (tooltip, annotation) |
| `src/services/dashboard.js` | ศูนย์กลาง UI orchestration — event binding, scroll spy, render pipeline |
| `src/services/fetcher.js` | ดึงข้อมูลจาก API พร้อมจัดการ retry, timeout, error state |
| `src/services/pdf.js` | ระบบ Export รายงานเป็น PDF |
| `src/services/pdfDatePicker.js` | ปฏิทินเลือกช่วงวันที่สำหรับ PDF Export พร้อม Drag & Drop |
| `src/store/reactiveState.js` | Reactive State Store core (Observer Pattern) |

### Backend & Cloud Deployment

```
Browser ──► Vercel Edge (/api/data) ──► Google Apps Script ──► Google Sheets
              │                              │
              │  Cache: SWR                   │  Token protected
              │  s-maxage=300                 │  (server-to-server only)
              │  stale-while-revalidate=600   │
```

- **`api/data.js`** — Vercel Serverless Function ทำหน้าที่เป็น Middleware ซ่อน `GOOGLE_SCRIPT_URL` ไว้ที่ฝั่ง Server ผู้ใช้ไม่เห็น URL ต้นทางโดยตรง
- **`vercel.json`** — จัดการ Routing ผ่าน `rewrites` และ `headers` พร้อม Edge Caching แบบ **Stale-While-Revalidate** ป้องกัน Quota Exceeded ของ Google Sheets

---

## UI/UX Micro-Tuning

### Native & Fluid Scroll
ใช้ Custom Lerp-based Scroll Engine สำหรับ Mouse Wheel และ Native Scroll สำหรับ Trackpad ทำให้การเลื่อนหน้าจอลื่นไหลตามธรรมชาติในทุกอุปกรณ์

### Zero Layout Shift Skeleton Loader
Skeleton Layer ใช้ `position: absolute; inset: 0` ซ้อนทับเนื้อหาเดิม พร้อมเอฟเฟกต์ **Shimmer** ทำให้กล่องข้อมูลไม่ยุบตัวหรือขยับขณะรอข้อมูล

### Smooth Sidebar Navigation & Scroll Spy
- **`isScrollingFromNav`** — ตัวแปรล็อกสถานะชั่วคราวขณะกดเมนู ป้องกัน Scroll Spy ทำงานซ้ำซ้อน
- **`will-change: top, height`** — สั่ง GPU รันล่วงหน้าสำหรับ `.nav-indicator`
- **`requestAnimationFrame`** — เกลี่ยการคำนวณ `getBoundingClientRect()` ให้ตรงกับ Refresh Rate ของจอ

### Page Visibility Aware Countdown
ใช้ Page Visibility API หยุด Countdown Timer อัตโนมัติเมื่อ Tab ซ่อนอยู่ และ Resume อย่างถูกต้องเมื่อ Tab กลับมา Active — ป้องกัน CPU Waste และการ Fetch ที่ไม่จำเป็น

### Pre-computed Search Index
แต่ละ Row ใน Dataset มี `_search` field ที่คำนวณล่วงหน้าตอน Normalize ทำให้การกรองข้อมูล Table เป็น O(1) String Lookup แทนการสร้าง String ใหม่ทุกครั้ง

---

## โครงสร้างโปรเจกต์

```
Atmosfera/
├── api/
│   └── data.js                 # Vercel Serverless — Data Proxy Middleware
├── assets/
│   └── favicon.png             # Favicon
├── src/
│   ├── index.html              # Entry Point
│   ├── app.js                  # Bootstrap — DOMContentLoaded → initApp()
│   ├── config.js               # Shared configuration constants
│   ├── style.css               # Design System — Custom Properties, Animations, Layout
│   ├── charts/
│   │   ├── dashboardCharts.js  # Chart.js initialization & update pipeline
│   │   └── plugins.js          # Custom Chart.js plugins
│   ├── services/
│   │   ├── dashboard.js        # UI orchestration — events, scroll spy, render
│   │   ├── fetcher.js          # Data fetching with retry & error handling
│   │   ├── pdf.js              # PDF export engine
│   │   └── pdfDatePicker.js    # Calendar picker for PDF date range
│   ├── store/
│   │   └── reactiveState.js    # Reactive State Store (Observer Pattern)
│   └── utils/
│       ├── aqiHelper.js        # AQI calculation & threshold metadata
│       ├── dateHelper.js       # Date/time formatting utilities
│       └── escapeHtml.js       # XSS-safe HTML escaping
├── vercel.json                 # Vercel configuration — rewrites, headers, caching
└── README.md                   # This file
```

---

## การตั้งค่าเพื่อใช้งานจริง

### 1. Deploy บน Vercel

#### ตั้งค่า Environment Variables

ไปที่ **Vercel Dashboard → Project → Settings → Environment Variables** แล้วเพิ่ม:

| Variable | คำอธิบาย | ตัวอย่าง |
|---|---|---|
| `GOOGLE_SCRIPT_URL` | URL ของ Google Apps Script Web App | `https://script.google.com/macros/s/.../exec` |
| `API_TOKEN` | Secret Token สำหรับ Authenticate กับ Apps Script (ไม่จำเป็น ถ้า Apps Script ไม่ได้ตรวจ token) | `your-secret-token-here` |

> **หมายเหตุ:** ถ้า Apps Script ไม่ได้ตรวจ token ให้ตั้งเฉพาะ `GOOGLE_SCRIPT_URL` ได้เลย

#### Deploy

```bash
# ติดตั้ง Vercel CLI (ถ้ายังไม่มี)
npm i -g vercel

# Login
vercel login

# Deploy
vercel --prod
```

### 2. รันในเครื่อง (Local Development)

```bash
# Clone โปรเจกต์
git clone https://github.com/magg2547-creator/Atmosfera.git
cd Atmosfera

# สร้างไฟล์ .env.local สำหรับตัวแปรสภาพแวดล้อม
echo "GOOGLE_SCRIPT_URL=https://script.google.com/macros/s/.../exec" >> .env.local
# Optional: ตั้งเฉพาะกรณี Apps Script ตรวจ token
# echo "API_TOKEN=your-secret-token-here" >> .env.local

# รัน development server
vercel dev
```

จากนั้นเปิด Browser ไปที่ `http://localhost:3000`

---

## License

Private — All rights reserved
