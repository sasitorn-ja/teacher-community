import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
const cors = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type' }
Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: cors })
  const service = createClient(Deno.env.get('SUPABASE_URL')!, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!)
  const token = (request.headers.get('Authorization') ?? '').replace('Bearer ', '')
  const { data: { user } } = await service.auth.getUser(token)
  const { data: caller } = user ? await service.from('teacher_profiles').select('role').eq('id', user.id).single() : { data: null }
  if (caller?.role !== 'admin') return Response.json({ error: 'Admin only' }, { status: 403, headers: cors })
  const { action, id, teacher_name, community_code } = await request.json()
  if (action === 'create') {
    const login_email = `teacher-${crypto.randomUUID()}@teacher-community.local`
    const { data, error } = await service.auth.admin.createUser({ email: login_email, password: community_code, email_confirm: true })
    if (error || !data.user) return Response.json({ error: error?.message ?? 'Create failed' }, { status: 400, headers: cors })
    const { error: profileError } = await service.from('teacher_profiles').insert({ id:data.user.id, teacher_name, community_code, login_email, role:'teacher' })
    if (profileError) { await service.auth.admin.deleteUser(data.user.id); return Response.json({ error:profileError.message }, { status:400, headers:cors }) }
  } else if (action === 'update') {
    const auth = await service.auth.admin.updateUserById(id, { password:community_code })
    const { error } = auth.error ? { error:auth.error } : await service.from('teacher_profiles').update({ teacher_name, community_code }).eq('id',id)
    if (error) return Response.json({ error:error.message }, { status:400, headers:cors })
  } else if (action === 'delete') {
    const { error } = await service.auth.admin.deleteUser(id)
    if (error) return Response.json({ error:error.message }, { status:400, headers:cors })
  } else return Response.json({ error:'Unknown action' }, { status:400, headers:cors })
  return Response.json({ ok:true }, { headers:cors })
})
