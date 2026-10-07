const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type, x-supabase-api-version',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Max-Age': '86400',
}

function getServiceKey() {
  const explicitKey = Deno.env.get('PASSWORDLESS_SERVICE_KEY')
  if (explicitKey) return explicitKey
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

const projectUrl = 'https://kksdtpygqjhrscvfgugw.supabase.co'
const projectPublishableKey = 'sb_publishable_iShvkD-CTsnmPtLFMWjXsg_23pgEqPh'
type LoginProfile = {
  id: string
  teacher_name: string
  role: string
  login_email: string
}

function response(body: Record<string, unknown>, status = 200) {
  return Response.json(body, { status, headers: cors })
}

function serviceHeaders(serviceKey: string) {
  return {
    apikey: serviceKey,
    Authorization: `Bearer ${serviceKey}`,
    'Content-Type': 'application/json',
  }
}

function publicHeaders() {
  return {
    apikey: projectPublishableKey,
    'Content-Type': 'application/json',
  }
}

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: cors })
  if (request.method !== 'POST') return response({ error: 'Method not allowed' }, 405)

  try {
    const body = await request.json() as { username?: string; adminOnly?: boolean }
    // Thai names can contain Sara Am (ำ); Unicode NFKC decomposes it and
    // makes an exact PostgreSQL name lookup miss the stored account.
    const username = String(body.username ?? '').trim()
    const adminOnly = body.adminOnly === true
    if (!username) return response({ error: 'กรุณากรอก Username' }, 400)

    const serviceKey = getServiceKey()
    if (!serviceKey) return response({ error: 'ยังไม่ได้ตั้งค่า service key ของระบบ' }, 500)

    const filters = new URLSearchParams({ select: 'id,teacher_name,role,login_email', limit: '1' })
    if (username.toLowerCase() === 'admin') filters.set('role', 'eq.admin')
    else filters.set('teacher_name', `eq.${username}`)

    // Use the REST endpoint directly. This avoids the hosted runtime's stale
    // Supabase client context while still authenticating every request with the
    // project's server-only secret.
    const profileResponse = await fetch(`${projectUrl}/rest/v1/teacher_profiles?${filters}`, {
      headers: serviceHeaders(serviceKey),
    })
    const profileBody = await profileResponse.json()
    if (!profileResponse.ok) {
      return response({ error: profileBody.message ?? 'ค้นหาบัญชีไม่สำเร็จ' }, profileResponse.status)
    }
    let profile = Array.isArray(profileBody) ? profileBody[0] as LoginProfile | undefined : undefined
    if (!profile && !adminOnly && username.toLowerCase() !== 'admin') {
      const lookupResponse = await fetch(projectUrl + '/rest/v1/rpc/lookup_login_email', {
        method: 'POST',
        headers: publicHeaders(),
        body: JSON.stringify({ p_teacher_name: username }),
      })
      const loginEmail = await lookupResponse.json()
      if (!lookupResponse.ok) {
        return response({ error: loginEmail.message ?? 'ค้นหาบัญชีไม่สำเร็จ' }, lookupResponse.status)
      }
      if (typeof loginEmail === 'string' && loginEmail) {
        profile = { id: '', teacher_name: username, role: 'teacher', login_email: loginEmail }
      }
    }
    if (!profile) return response({ error: 'ไม่พบ Username นี้ในระบบ' }, 404)
    if (adminOnly && profile.role !== 'admin') return response({ error: 'บัญชีนี้ไม่มีสิทธิ์ผู้ดูแล' }, 403)
    if (!profile.login_email) return response({ error: 'บัญชีนี้ยังไม่มี login_email' }, 400)

    const linkResponse = await fetch(`${projectUrl}/auth/v1/admin/generate_link`, {
      method: 'POST',
      headers: serviceHeaders(serviceKey),
      body: JSON.stringify({ type: 'magiclink', email: profile.login_email }),
    })
    const linkBody = await linkResponse.json()
    const tokenHash = linkBody.hashed_token
    if (!linkResponse.ok || !tokenHash) {
      return response({ error: linkBody.message ?? 'สร้าง passwordless session ไม่สำเร็จ' }, linkResponse.status || 400)
    }

    return response({ token_hash: tokenHash })
  } catch (error) {
    return response({ error: error instanceof Error ? error.message : 'เข้าสู่ระบบไม่สำเร็จ' }, 400)
  }
})
