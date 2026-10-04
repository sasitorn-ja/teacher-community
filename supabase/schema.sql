-- Teacher Community: account-based schema
-- Teacher name is the username; community code is the Supabase Auth password.
create table if not exists public.communities (
  id uuid primary key default gen_random_uuid(),
  activity_date date not null default current_date,
  community_code text not null check (community_code ~ '^กก[0-9]{3,}$'),
  community_name text not null, advisor_name text not null, school_name text not null,
  location text not null default '', member_count integer not null default 0 check (member_count >= 0),
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
alter table public.communities enable row level security;
alter table public.teacher_profiles enable row level security;
create policy "public community cards" on public.communities for select using (true);
create policy "teacher reads profile" on public.teacher_profiles for select using (id = auth.uid() or public.is_admin());
create policy "teacher creates community" on public.communities for insert to authenticated with check (owner_id = auth.uid());
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
  and (storage.foldername(name))[1] = auth.uid()::text
);

drop policy if exists "community image deletes" on storage.objects;
create policy "community image deletes"
on storage.objects for delete to authenticated
using (
  bucket_id = 'community-images'
  and (storage.foldername(name))[1] = auth.uid()::text
);
