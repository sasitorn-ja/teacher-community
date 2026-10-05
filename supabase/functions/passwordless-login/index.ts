import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Max-Age': '86400',
}

const service = createClient(
  Deno.env.get('SUPABASE_URL')!,
  Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
)

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

    const { data: profile, error: profileError } = username.toLowerCase() === 'admin'
      ? await profileQuery.eq('role', 'admin').limit(1).maybeSingle()
      : await profileQuery.eq('teacher_name', username).maybeSingle()

    if (profileError) return response({ error: profileError.message }, 400)
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
