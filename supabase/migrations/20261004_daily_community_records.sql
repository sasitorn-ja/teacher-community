-- Keep one editable community record per teacher and Bangkok calendar day.
alter table public.communities
  add column if not exists activity_date date not null default current_date;

alter table public.communities
  drop constraint if exists communities_owner_id_key;

alter table public.communities
  drop constraint if exists communities_community_code_key;

create unique index if not exists communities_owner_activity_date_key
  on public.communities (owner_id, activity_date)
  where owner_id is not null;
