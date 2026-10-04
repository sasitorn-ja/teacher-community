-- Read-only usage report for the admin storage dashboard.
-- It reports PostgreSQL table/index sizes and Supabase Storage object sizes;
-- plan-level billing/egress quotas are not exposed to browser clients.
create or replace function public.get_database_usage()
returns table(
  object_name text,
  object_type text,
  row_count bigint,
  table_bytes bigint,
  index_bytes bigint,
  total_bytes bigint
)
language plpgsql
security definer
set search_path = public, pg_catalog
as $$
begin
  if not public.is_admin() then
    raise exception 'Admin only';
  end if;

  return query
  select
    c.relname::text,
    'database table'::text,
    coalesce(stats.n_live_tup::bigint, 0),
    pg_table_size(c.oid::regclass),
    pg_indexes_size(c.oid::regclass),
    pg_total_relation_size(c.oid::regclass)
  from pg_class c
  join pg_namespace n on n.oid = c.relnamespace
  left join pg_stat_user_tables stats on stats.relid = c.oid
  where n.nspname = 'public' and c.relkind = 'r'

  union all

  select
    'community-images'::text,
    'storage bucket'::text,
    count(*)::bigint,
    0::bigint,
    0::bigint,
    coalesce(sum(nullif((metadata->>'size'), '')::bigint), 0)::bigint
  from storage.objects
  where bucket_id = 'community-images';
end;
$$;

revoke all on function public.get_database_usage() from public;
grant execute on function public.get_database_usage() to authenticated;
