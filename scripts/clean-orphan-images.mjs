// List (and optionally delete) files in the community-images bucket that no community record points to.
// Usage:
//   SUPABASE_SERVICE_ROLE_KEY=... node --env-file=.env.local scripts/clean-orphan-images.mjs           # dry run
//   SUPABASE_SERVICE_ROLE_KEY=... node --env-file=.env.local scripts/clean-orphan-images.mjs --delete  # delete
import { createClient } from '@supabase/supabase-js'

const BUCKET = 'community-images'
// Skip very recent files: a teacher may have uploaded but not finished saving yet.
const MIN_AGE_MS = 60 * 60 * 1000

const url = process.env.SUPABASE_URL ?? process.env.VITE_SUPABASE_URL
const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY
if (!url || !serviceKey) {
  console.error('ต้องตั้งค่า VITE_SUPABASE_URL และ SUPABASE_SERVICE_ROLE_KEY ก่อน')
  process.exit(1)
}
const shouldDelete = process.argv.includes('--delete')
const supabase = createClient(url, serviceKey, { auth: { persistSession: false } })

async function listAll(prefix = '') {
  const files = []
  for (let offset = 0; ; offset += 1000) {
    const { data, error } = await supabase.storage.from(BUCKET).list(prefix, { limit: 1000, offset })
    if (error) throw new Error(`อ่านรายการไฟล์ไม่สำเร็จ: ${error.message}`)
    for (const item of data) {
      const path = prefix ? `${prefix}/${item.name}` : item.name
      if (item.id === null) files.push(...await listAll(path)) // folder
      else files.push({ path, size: item.metadata?.size ?? 0, createdAt: item.created_at })
    }
    if (data.length < 1000) return files
  }
}

const { data: communities, error } = await supabase.from('communities').select('image_url')
if (error) throw new Error(`อ่านข้อมูลชุมชนไม่สำเร็จ: ${error.message}`)
const usedPaths = new Set()
for (const { image_url } of communities) {
  const match = image_url?.match(new RegExp(`/storage/v1/object/(?:public|sign|authenticated)/${BUCKET}/([^?]+)`))
  if (match) usedPaths.add(decodeURIComponent(match[1]))
}

const files = await listAll()
const now = Date.now()
const orphans = files.filter((file) => !usedPaths.has(file.path) && now - new Date(file.createdAt).getTime() > MIN_AGE_MS)
const kb = (bytes) => `${(bytes / 1024).toFixed(1)} KB`

console.log(`ไฟล์ทั้งหมด ${files.length} ไฟล์ · ถูกใช้งาน ${usedPaths.size} ไฟล์ · ไม่มีข้อมูลใดใช้ ${orphans.length} ไฟล์`)
for (const file of orphans) console.log(`  ${file.path}  (${kb(file.size)})`)
console.log(`พื้นที่ที่คืนได้ ${kb(orphans.reduce((sum, file) => sum + file.size, 0))}`)

if (!orphans.length) process.exit(0)
if (!shouldDelete) {
  console.log('\nนี่คือการตรวจสอบเท่านั้น ยังไม่ได้ลบไฟล์ · รันซ้ำพร้อม --delete เพื่อลบจริง')
  process.exit(0)
}
for (let i = 0; i < orphans.length; i += 100) {
  const { error: removeError } = await supabase.storage.from(BUCKET).remove(orphans.slice(i, i + 100).map((file) => file.path))
  if (removeError) throw new Error(`ลบไฟล์ไม่สำเร็จ: ${removeError.message}`)
}
console.log(`ลบไฟล์ที่ไม่ได้ใช้ ${orphans.length} ไฟล์แล้ว`)
