-- Keep a durable count of every successful save, including edits on the same day.
create table if not exists public.community_submission_events (
  id uuid primary key default gen_random_uuid(),
  owner_id uuid not null references auth.users(id) on delete cascade,
  community_id uuid references public.communities(id) on delete cascade,
  activity_date date not null default current_date,
  submitted_at timestamptz not null default now()
);

create index if not exists community_submission_events_owner_date_idx
  on public.community_submission_events (owner_id, activity_date, submitted_at desc);

alter table public.community_submission_events enable row level security;

drop policy if exists "teachers read own submission events" on public.community_submission_events;
create policy "teachers read own submission events"
  on public.community_submission_events for select to authenticated
  using (owner_id = auth.uid() or public.is_admin());

drop policy if exists "teachers create submission events" on public.community_submission_events;
create policy "teachers create submission events"
  on public.community_submission_events for insert to authenticated
  with check (owner_id = auth.uid() or public.is_admin());
