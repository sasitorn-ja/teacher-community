import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const cors = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Max-Age': '86400',
}
const service = createClient(Deno.env.get('SUPABASE_URL')!, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!)

type TeacherInput = { teacher_name?: string; community_code?: string }
const normalizeCode = (value: string) => value.normalize('NFKC').replace(/\s+/g, '').trim()

async function assertValidAndAvailableCode(community_code: string, excludeId?: string) {
  const code = normalizeCode(community_code)
  if (!/^กก\d{3,}$/.test(code) && code !== 'admin026') {
    throw new Error('รูปแบบรหัสชุมชนไม่ถูกต้อง กรุณาใช้ เช่น กก126')
  }

  const { data: existing, error } = await service
    .from('teacher_profiles')
    .select('id,teacher_name')
    .eq('community_code', code)
    .maybeSingle()
  if (error) throw new Error(error.message)
  if (existing && existing.id !== excludeId) {
    throw new Error(`รหัสชุมชน ${code} ถูกใช้โดย ${existing.teacher_name} แล้ว`)
  }
}

async function createTeacher(teacher_name: string, community_code: string) {
  community_code = normalizeCode(community_code)
  if (!teacher_name || !community_code) throw new Error('กรุณากรอกชื่อครูและรหัสชุมชน')
  await assertValidAndAvailableCode(community_code)
  const login_email = `teacher-${crypto.randomUUID()}@teacher-community.local`
  const { data, error } = await service.auth.admin.createUser({
    email: login_email,
    password: community_code,
    email_confirm: true,
  })
  if (error || !data.user) throw new Error(error?.message ?? 'Create failed')

  const { error: profileError } = await service.from('teacher_profiles').insert({
    id: data.user.id,
    teacher_name,
    community_code,
    login_email,
    role: 'teacher',
  })
  if (profileError) {
    await service.auth.admin.deleteUser(data.user.id)
    throw new Error(profileError.message)
  }
}

async function updateTeacher(id: string, teacher_name: string, community_code: string) {
  community_code = normalizeCode(community_code)
  if (!id || !teacher_name || !community_code) throw new Error('กรุณากรอกชื่อครูและรหัสชุมชน')
  await assertValidAndAvailableCode(community_code, id)
  const auth = await service.auth.admin.updateUserById(id, { password: community_code })
  if (auth.error) throw new Error(auth.error.message)

  const { error } = await service.from('teacher_profiles')
    .update({ teacher_name, community_code })
    .eq('id', id)
  if (error) throw new Error(error.message)
}

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: cors })

  const token = (request.headers.get('Authorization') ?? '').replace('Bearer ', '')
  const { data: { user } } = await service.auth.getUser(token)
  const { data: caller } = user
    ? await service.from('teacher_profiles').select('role').eq('id', user.id).single()
    : { data: null }

  if (caller?.role !== 'admin') {
    return Response.json({ error: 'Admin only' }, { status: 403, headers: cors })
  }

  const body = await request.json()
  const { action, id, teacher_name, community_code } = body

  if (action === 'create') {
    try {
      await createTeacher(String(teacher_name ?? '').trim(), String(community_code ?? '').trim())
    } catch (error) {
      return Response.json(
        { error: error instanceof Error ? error.message : 'Create failed' },
        { status: 400, headers: cors },
      )
    }
  } else if (action === 'update') {
    try {
      await updateTeacher(id, String(teacher_name ?? '').trim(), String(community_code ?? '').trim())
    } catch (error) {
      return Response.json(
        { error: error instanceof Error ? error.message : 'Update failed' },
        { status: 400, headers: cors },
      )
    }
  } else if (action === 'delete') {
    const { error } = await service.auth.admin.deleteUser(id)
    if (error) return Response.json({ error: error.message }, { status: 400, headers: cors })
  } else if (action === 'bulk-create') {
    const rows = Array.isArray(body.accounts) ? body.accounts as TeacherInput[] : []
    let created = 0
    let updated = 0
    const errors: string[] = []
    const batchCodes = new Map<string, string>()

    for (const row of rows) {
      const name = String(row.teacher_name ?? '').trim()
      const code = normalizeCode(String(row.community_code ?? ''))
      if (!name || !code) {
        errors.push(`${name || 'ไม่ทราบชื่อ'}: ข้อมูลไม่ครบ`)
        continue
      }
      const codeKey = code.toLowerCase()
      const previousName = batchCodes.get(codeKey)
      if (previousName && previousName !== name) {
        errors.push(`${name}: รหัสชุมชน ${code} ซ้ำกับ ${previousName} ในไฟล์นำเข้า`)
        continue
      }
      batchCodes.set(codeKey, name)

      try {
        const { data: existing, error: lookupError } = await service
          .from('teacher_profiles')
          .select('id,role')
          .eq('teacher_name', name)
          .maybeSingle()
        if (lookupError) throw new Error(lookupError.message)

        if (existing) {
          if (existing.role !== 'teacher') throw new Error('เป็นบัญชีผู้ดูแล ไม่สามารถนำเข้าทับได้')
          await updateTeacher(existing.id, name, code)
          updated += 1
        } else {
          await createTeacher(name, code)
          created += 1
        }
      } catch (error) {
        errors.push(`${name}: ${error instanceof Error ? error.message : 'นำเข้าไม่สำเร็จ'}`)
      }
    }

    return Response.json(
      { ok: true, created, updated, failed: errors.length, errors: errors.slice(0, 20) },
      { headers: cors },
    )
  } else {
    return Response.json({ error: 'Unknown action' }, { status: 400, headers: cors })
  }

  return Response.json({ ok: true }, { headers: cors })
})
