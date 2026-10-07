import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Max-Age': '86400',
}

function getServiceKey() {
  const legacyKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')
  const configuredKeys = Deno.env.get('SUPABASE_SECRET_KEYS')
  if (!configuredKeys) return legacyKey!
  try {
    const keys = JSON.parse(configuredKeys) as Record<string, unknown>
    const candidates = [...Object.keys(keys), ...Object.values(keys).filter((value): value is string => typeof value === 'string')]
    return candidates.find((value) => value.startsWith('sb_secret_')) ?? legacyKey!
  } catch {
    return legacyKey!
  }
}

// Use this project's public API host explicitly. In the hosted function
// runtime, request.url can be an internal gateway URL rather than this host.
const projectUrl = 'https://kksdtpygqjhrscvfgugw.supabase.co'
// Publishable keys are intentionally public. Keeping the project's current
// key here makes the directory RPC independent of function-gateway headers.
const projectPublishableKey = 'sb_publishable_iShvkD-CTsnmPtLFMWjXsg_23pgEqPh'
const service = createClient(
  projectUrl,
  getServiceKey(),
)
type LoginProfile = {
  id: string
  teacher_name: string
  role: string
  login_email: string
}

function response(body: Record<string, unknown>, status = 200) {
  return Response.json(body, { status, headers: cors })
}

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: cors })
  if (request.method !== 'POST') return response({ error: 'Method not allowed' }, 405)

  try {
    const body = await request.json() as { username?: string; adminOnly?: boolean }
    const username = String(body.username ?? '').normalize('NFKC').trim()
    const adminOnly = body.adminOnly === true
    if (!username) return response({ error: 'กรุณากรอก Username' }, 400)

    const profileQuery = service
      .from('teacher_profiles')
      .select('id, teacher_name, role, login_email')

    const { data: matchedProfile, error: profileError } = username.toLowerCase() === 'admin'
      ? await profileQuery.eq('role', 'admin').limit(1).maybeSingle()
      : await profileQuery.eq('teacher_name', username).maybeSingle()
    let profile = matchedProfile as LoginProfile | null

    if (profileError) return response({ error: profileError.message }, 400)
    // The public RPC is a safe fallback for teacher lookups: it is already
    // security-definer and exposes only the login address needed below.
    // This keeps existing accounts usable if a deployed function runs with a
    // stale service-role context while the database record itself is present.
    if (!profile && !adminOnly && username.toLowerCase() !== 'admin') {
      const directory = createClient(projectUrl, projectPublishableKey)
      const { data: loginEmail, error: lookupError } = await directory
        .rpc('lookup_login_email', { p_teacher_name: username })
      if (lookupError) return response({ error: lookupError.message }, 400)
      if (loginEmail) {
        profile = { id: '', teacher_name: username, role: 'teacher', login_email: loginEmail }
      }
    }
    if (!profile) return response({ error: 'ไม่พบ Username นี้ในระบบ' }, 404)
    if (adminOnly && profile.role !== 'admin') return response({ error: 'บัญชีนี้ไม่มีสิทธิ์ผู้ดูแล' }, 403)
    if (!profile.login_email) return response({ error: 'บัญชีนี้ยังไม่มี login_email' }, 400)

    // Supabase Auth still needs a real user session for RLS. The one-time
    // magic-link token is generated server-side, so the user never enters a password.
    const { data, error } = await service.auth.admin.generateLink({
      type: 'magiclink',
      email: profile.login_email,
    })
    const tokenHash = data?.properties?.hashed_token
    if (error || !tokenHash) return response({ error: error?.message ?? 'สร้าง passwordless session ไม่สำเร็จ' }, 400)

    return response({ token_hash: tokenHash })
  } catch (error) {
    return response({ error: error instanceof Error ? error.message : 'เข้าสู่ระบบไม่สำเร็จ' }, 400)
  }
})
