-- Public display settings are stored in Supabase instead of localStorage.
create or replace function public.list_teacher_directory()
returns table(id uuid, teacher_name text)
language sql stable security definer set search_path = public as $$
  select id, teacher_name from public.teacher_profiles where role = 'teacher' order by teacher_name;
$$;
revoke all on function public.list_teacher_directory() from public;
grant execute on function public.list_teacher_directory() to anon, authenticated;

drop policy if exists "admin updates own profile" on public.teacher_profiles;
create policy "admin updates own profile" on public.teacher_profiles for update to authenticated using (id = auth.uid() and role = 'admin') with check (id = auth.uid() and role = 'admin');

create table if not exists public.system_settings (
  id text primary key default 'main' check (id = 'main'),
  title text not null default 'ระบบลงทะเบียนชุมนุมคุณครู',
  term text not null default 'ภาคเรียนที่ 2',
  year text not null default 'ปีการศึกษา 2569',
  school text not null default 'โรงเรียนวิเชียรมาตุ',
  updated_at timestamptz not null default now()
);
insert into public.system_settings (id) values ('main') on conflict (id) do nothing;
alter table public.system_settings enable row level security;
drop policy if exists "public system settings" on public.system_settings;
create policy "public system settings" on public.system_settings for select using (true);
drop policy if exists "admin manages system settings" on public.system_settings;
create policy "admin manages system settings" on public.system_settings for all to authenticated using (public.is_admin()) with check (public.is_admin());
