import { createClient } from '@supabase/supabase-js'

const url = String(import.meta.env.VITE_SUPABASE_URL ?? '').trim()
const anonKey = String(import.meta.env.VITE_SUPABASE_ANON_KEY ?? '').trim()
const validUrl = (() => {
  try {
    const parsed = new URL(url)
    return parsed.protocol === 'https:' && !url.includes('YOUR_PROJECT')
  } catch {
    return false
  }
})()

export const hasSupabaseConfig = Boolean(validUrl && anonKey && !anonKey.includes('YOUR_SUPABASE'))
export const supabaseConfigError = hasSupabaseConfig
  ? ''
  : 'ยังไม่ได้ตั้งค่า VITE_SUPABASE_URL และ VITE_SUPABASE_ANON_KEY ใน Vercel หรือค่าที่ตั้งไว้ไม่ถูกต้อง'

async function fetchWithTimeout(input: RequestInfo | URL, init?: RequestInit) {
  const controller = new AbortController()
  const timeoutId = window.setTimeout(() => controller.abort(), 15000)
  const parentSignal = init?.signal
  const abortFromParent = () => controller.abort()
  if (parentSignal?.aborted) abortFromParent()
  else parentSignal?.addEventListener('abort', abortFromParent, { once: true })
  try {
    return await fetch(input, { ...init, signal: controller.signal })
  } finally {
    window.clearTimeout(timeoutId)
    parentSignal?.removeEventListener('abort', abortFromParent)
  }
}

export const supabase = hasSupabaseConfig
  ? createClient(url, anonKey, { global: { fetch: fetchWithTimeout } })
  : null
