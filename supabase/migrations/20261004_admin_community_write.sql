-- Admins can save community records on behalf of a selected teacher from
-- the admin entry screen. Teachers remain limited to their own records.
drop policy if exists "teacher creates community" on public.communities;
create policy "teacher creates community"
on public.communities for insert to authenticated
with check (owner_id = auth.uid() or public.is_admin());

drop policy if exists "teacher edits community" on public.communities;
create policy "teacher edits community"
on public.communities for update to authenticated
using (owner_id = auth.uid() or public.is_admin())
with check (owner_id = auth.uid() or public.is_admin());
