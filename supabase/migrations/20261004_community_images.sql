-- Store community images in Supabase Storage only when the teacher saves the form.
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
