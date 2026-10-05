-- Teacher Community: account-based schema
-- Teacher name is the username; login uses the passwordless-login Edge Function.
create table if not exists public.communities (
  id uuid primary key default gen_random_uuid(),
  activity_date date not null default current_date,
  community_code text not null check (community_code ~ '^กก[0-9]{3,}$'),
  community_name text not null, advisor_name text not null, school_name text not null,
  location text not null default '', member_count integer not null default 22 check (member_count >= 22),
  description text not null default '', image_url text,
  owner_id uuid references auth.users(id) on delete set null,
  unique (owner_id, activity_date),
  created_at timestamptz not null default now(), updated_at timestamptz not null default now()
);
create table if not exists public.teacher_profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  teacher_name text unique not null,
  community_code text unique not null check (community_code ~ '^กก[0-9]{3,}$' or community_code = 'admin026'),
  login_email text unique not null,
  role text not null default 'teacher' check (role in ('teacher', 'admin')),
  community_id uuid unique references public.communities(id) on delete set null,
  created_at timestamptz not null default now()
);
create or replace function public.is_admin() returns boolean language sql stable security definer set search_path = public as $$
  select exists(select 1 from public.teacher_profiles where id = auth.uid() and role = 'admin');
$$;
create or replace function public.lookup_login_email(p_teacher_name text) returns text language sql stable security definer set search_path = public as $$
  select login_email from public.teacher_profiles where teacher_name = trim(p_teacher_name) limit 1;
$$;
revoke all on function public.lookup_login_email(text) from public;
grant execute on function public.lookup_login_email(text) to anon, authenticated;
create or replace function public.list_teacher_directory()
returns table(id uuid, teacher_name text)
language sql stable security definer set search_path = public as $$
  select id, teacher_name from public.teacher_profiles where role = 'teacher' order by teacher_name;
$$;
revoke all on function public.list_teacher_directory() from public;
grant execute on function public.list_teacher_directory() to anon, authenticated;

create or replace function public.get_database_usage()
returns table(object_name text, object_type text, row_count bigint, table_bytes bigint, index_bytes bigint, total_bytes bigint)
language plpgsql security definer set search_path = public, pg_catalog as $$
begin
  if not public.is_admin() then raise exception 'Admin only'; end if;
  return query
  select c.relname::text, 'database table'::text, coalesce(stats.n_live_tup::bigint,0), pg_table_size(c.oid::regclass), pg_indexes_size(c.oid::regclass), pg_total_relation_size(c.oid::regclass)
  from pg_class c join pg_namespace n on n.oid=c.relnamespace left join pg_stat_user_tables stats on stats.relid=c.oid
  where n.nspname='public' and c.relkind='r'
  union all
  select 'community-images'::text, 'storage bucket'::text, count(*)::bigint, 0::bigint, 0::bigint, coalesce(sum(nullif((metadata->>'size'), '')::bigint),0)::bigint
  from storage.objects where bucket_id='community-images';
end; $$;
revoke all on function public.get_database_usage() from public;
grant execute on function public.get_database_usage() to authenticated;

create table if not exists public.system_settings (
  id text primary key default 'main' check (id = 'main'),
  title text not null default 'ระบบลงทะเบียนชุมนุมคุณครู',
  term text not null default 'ภาคเรียนที่ 2',
  year text not null default 'ปีการศึกษา 2569',
  school text not null default 'โรงเรียนวิเชียรมาตุ',
  updated_at timestamptz not null default now()
);
insert into public.system_settings (id) values ('main') on conflict (id) do nothing;

create table if not exists public.community_submission_events (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  community_id uuid references public.communities(id) on delete cascade,
  activity_date date not null default current_date,
  submitted_at timestamptz not null default now()
);
create index if not exists community_submission_events_owner_date_idx
  on public.community_submission_events (owner_id, activity_date, submitted_at desc);

alter table public.system_settings enable row level security;
alter table public.community_submission_events enable row level security;
drop policy if exists "public system settings" on public.system_settings;
create policy "public system settings" on public.system_settings for select using (true);
drop policy if exists "admin manages system settings" on public.system_settings;
create policy "admin manages system settings" on public.system_settings for all to authenticated using (public.is_admin()) with check (public.is_admin());
create policy "teachers read own submission events" on public.community_submission_events for select to authenticated using (owner_id = auth.uid() or public.is_admin());
create policy "teachers create submission events" on public.community_submission_events for insert to authenticated with check (owner_id = auth.uid() or public.is_admin());

alter table public.communities enable row level security;
alter table public.teacher_profiles enable row level security;
create policy "public community cards" on public.communities for select using (true);
create policy "teacher reads profile" on public.teacher_profiles for select using (id = auth.uid() or public.is_admin());
create policy "admin updates own profile" on public.teacher_profiles for update to authenticated using (id = auth.uid() and role = 'admin') with check (id = auth.uid() and role = 'admin');
create policy "teacher creates community" on public.communities for insert to authenticated with check (owner_id = auth.uid() or public.is_admin());
create policy "teacher edits community" on public.communities for update to authenticated using (owner_id = auth.uid() or public.is_admin()) with check (owner_id = auth.uid() or public.is_admin());
create policy "admin deletes communities" on public.communities for delete to authenticated using (public.is_admin());
-- Bootstrap first admin after creating that person in Authentication > Users:
-- insert into public.teacher_profiles (id,teacher_name,community_code,login_email,role)
-- values ('AUTH_USER_UUID','ชื่อแอดมิน','admin026','admin@teacher-community.local','admin');

-- Community image uploads. Files are uploaded only when the form is saved.
insert into storage.buckets (id, name, public)
values ('community-images', 'community-images', true)
on conflict (id) do update set public = true;

drop policy if exists "community image uploads" on storage.objects;
create policy "community image uploads"
on storage.objects for insert to authenticated
with check (
  bucket_id = 'community-images'
  and ((storage.foldername(name))[1] = auth.uid()::text or public.is_admin())
);

drop policy if exists "community image deletes" on storage.objects;
create policy "community image deletes"
on storage.objects for delete to authenticated
using (
  bucket_id = 'community-images'
  and ((storage.foldername(name))[1] = auth.uid()::text or public.is_admin())
);

drop policy if exists "admin manages community images" on storage.objects;
create policy "admin manages community images"
on storage.objects for all to authenticated
using (bucket_id = 'community-images' and public.is_admin())
with check (bucket_id = 'community-images' and public.is_admin());
