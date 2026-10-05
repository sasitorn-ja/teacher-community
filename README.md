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
4. สร้างผู้ใช้แอดมิน 1 คนใน Authentication > Users แล้วรันคำสั่ง bootstrap ที่ท้ายไฟล์ SQL (แทน `AUTH_USER_UUID` ด้วย UUID ของแอดมิน)
5. Deploy Edge Functions `admin-accounts` และ `passwordless-login` เพื่อให้แอดมินจัดการบัญชีครู และให้ครูเข้าสู่ระบบด้วย Username โดยไม่ต้องใช้ Password
   ```bash
   supabase functions deploy passwordless-login --no-verify-jwt
   supabase functions deploy admin-accounts
   ```

ถ้าโปรเจกต์ Supabase ถูกสร้างไปแล้ว ให้รันไฟล์ใน `supabase/migrations/` ตามลำดับ โดยเฉพาะ `20261004_system_settings.sql` เพื่อเปิดการโหลดรายชื่อครูและค่าระบบจาก Supabase

ครูเข้าสู่ระบบด้วย **ชื่อ-นามสกุลครู** อย่างเดียว ส่วน Username `admin` จะเข้าสู่ role ผู้ดูแลโดยไม่ต้องใช้ Password ระบบจะสร้าง one-time Auth session ผ่าน Edge Function เพื่อให้สิทธิ์ RLS เดิมยังทำงานได้

รหัสชุมชนสร้างที่ฐานข้อมูลแบบ atomic เริ่มจาก `กก026` จึงไม่ซ้ำแม้มีการลงทะเบียนพร้อมกันจำนวนมาก

## LINE Login
LINE Login ใช้แทนการล็อกอินด้วยรหัสชุมชนได้ในอนาคต แต่เวอร์ชันนี้ใช้รูปแบบบัญชีที่แอดมินจัดการตามโจทย์. อย่าเก็บ channel secret, service-role key หรือ Supabase password ใน Vite environment.

## Vercel
เพิ่ม `VITE_SUPABASE_URL` และ `VITE_SUPABASE_ANON_KEY` ใน Environment Variables ของ Vercel (เลือก Production และ Preview ตามที่ใช้งาน) แล้วกด Redeploy ใหม่ทุกครั้งหลังเพิ่มหรือแก้ค่า ใช้ anon key เท่านั้น ไม่ใช่ service role key หรือ database password.

ถ้าไม่ได้ตั้งค่าครบใน Production ระบบจะหยุดพร้อมข้อความแจ้งเตือนแทนการเก็บข้อมูลใน `localStorage` เพื่อป้องกันข้อมูลบน Vercel แยกจากฐานข้อมูลโดยไม่รู้ตัว
