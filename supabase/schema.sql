-- Run once in Supabase SQL Editor after creating the project.
create table if not exists public.communities (
  id uuid primary key default gen_random_uuid(),
  community_code text unique not null,
  community_name text not null check (char_length(community_name) <= 160),
  advisor_name text not null check (char_length(advisor_name) <= 160),
  school_name text not null check (char_length(school_name) <= 180),
  location text not null default '', member_count integer not null default 0 check (member_count >= 0),
  description text not null default '', image_url text, owner_id uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(), updated_at timestamptz not null default now()
);
create or replace function public.assign_community_code() returns trigger language plpgsql as $$
declare next_number integer;
begin
  if new.community_code is null or new.community_code = '' then
    perform pg_advisory_xact_lock(26026);
    select coalesce(max(nullif(regexp_replace(community_code, '\\D', '', 'g'), '')::integer), 25) + 1 into next_number from public.communities;
    new.community_code := 'กก' || lpad(next_number::text, 3, '0');
  end if;
  new.updated_at := now(); return new;
end; $$;
drop trigger if exists communities_assign_code on public.communities;
create trigger communities_assign_code before insert or update on public.communities for each row execute function public.assign_community_code();
alter table public.communities enable row level security;
create policy "Community cards are public" on public.communities for select using (true);
create policy "Authenticated teachers can create" on public.communities for insert to authenticated with check (auth.uid() = owner_id);
create policy "Teachers edit their own community" on public.communities for update to authenticated using (auth.uid() = owner_id) with check (auth.uid() = owner_id);
-- Add admin user IDs below after creating their accounts.
-- create policy "Admins manage all communities" on public.communities for all to authenticated using (auth.uid() in ('ADMIN_USER_UUID')) with check (auth.uid() in ('ADMIN_USER_UUID'));
