-- New and updated community records must accept at least 22 students.
-- NOT VALID keeps existing historical rows readable without silently changing them;
-- PostgreSQL still enforces the rule for every new insert or update.
do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conrelid = 'public.communities'::regclass
      and conname = 'communities_member_count_min_check'
  ) then
    alter table public.communities
      add constraint communities_member_count_min_check
      check (member_count >= 22) not valid;
  end if;
end $$;
