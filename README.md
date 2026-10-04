# Teacher Community

เว็บลงทะเบียนชุมชนคุณครู สร้างด้วย React + TypeScript พร้อมหน้าครูและแดชบอร์ดแอดมินสำหรับส่งออก Excel/ภาพ PNG

## Run locally
```bash
npm install
cp .env.example .env.local
npm run dev
```
เมื่อไม่มีค่า Supabase เว็บจะทำงานในโหมดตัวอย่าง (เก็บข้อมูลในเบราว์เซอร์เท่านั้น)

## Connect Supabase
1. สร้างโปรเจกต์ Supabase ชื่อ `teacher-community`
2. รัน `supabase/schema.sql` ใน SQL Editor
3. คัดลอก Project URL และ anon key ลงใน `.env.local`
4. เพิ่ม URL ของ Vercel ใน Authentication > URL Configuration > Redirect URL
5. เปิด Email provider เพื่อใช้ Magic Link สำหรับครูในรุ่นแรก

รหัสชุมชนสร้างที่ฐานข้อมูลแบบ atomic เริ่มจาก `กก026` จึงไม่ซ้ำแม้มีการลงทะเบียนพร้อมกันจำนวนมาก

## LINE Login
LINE Login ใช้แทน Magic Link ได้ แต่ Supabase Auth ไม่มี provider LINE สำเร็จรูป. ทางที่ปลอดภัยคือใช้ LINE Login/LIFF ผ่าน Edge Function หรือ backend ของเราเพื่อยืนยัน LINE ID แล้วออก session ของระบบเอง. อย่าเก็บ channel secret, service-role key หรือ Supabase password ใน Vite environment.

## Vercel
เพิ่ม `VITE_SUPABASE_URL` และ `VITE_SUPABASE_ANON_KEY` ใน Environment Variables ของ Vercel แล้ว deploy. ใช้ anon key เท่านั้น ไม่ใช่ service role key หรือ database password.
