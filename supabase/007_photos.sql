-- PVH operator photos -- run once in the Supabase SQL editor, after schema.sql.
-- Safe to run again.
--
-- A private storage bucket holding one small face thumbnail per operator.
--
--   * Read:   only people with a `profiles` row (the same rule as notes and checks).
--             Being able to sign in is not enough on its own. There is no public
--             or anonymous access: the bucket is private and no policy names `anon`.
--   * Write:  only people listed in `photo_admins` -- the administrator, never an
--             officer. Officers cannot add, replace or delete a photo from the app.
--   * Names:  a file must be named like  0123456789abcdef01234567.jpg  (24 hex
--             characters). The name is a one-way hash of the record key, so a
--             licence number never appears in a file name or URL.
--   * Limits: JPEG only, 200 KB at most per file.
--
-- Nothing here touches the existing tables.

-- ------------------------------------------------------------------ bucket
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('operator-photos', 'operator-photos', false, 204800, array['image/jpeg'])
on conflict (id) do update
  set public = false,
      file_size_limit = 204800,
      allowed_mime_types = array['image/jpeg'];

-- ------------------------------------------------------------ photo admins
create table if not exists public.photo_admins (
  id uuid primary key references auth.users(id) on delete cascade
);
alter table public.photo_admins enable row level security;
-- No policy on purpose: nobody can read or change this list through the API.
-- The administrator adds themselves in the SQL editor (see below).

create or replace function public.is_photo_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.photo_admins where id = auth.uid());
$$;
revoke all on function public.is_photo_admin() from public, anon;
grant execute on function public.is_photo_admin() to authenticated;

-- ---------------------------------------------------------------- policies
drop policy if exists photos_read   on storage.objects;
drop policy if exists photos_insert on storage.objects;
drop policy if exists photos_update on storage.objects;
drop policy if exists photos_delete on storage.objects;

create policy photos_read on storage.objects for select to authenticated
  using (bucket_id = 'operator-photos' and (public.is_pvh_user() or public.is_photo_admin()));

create policy photos_insert on storage.objects for insert to authenticated
  with check (bucket_id = 'operator-photos' and public.is_photo_admin()
              and name ~ '^[0-9a-f]{24}\.jpg$');

create policy photos_update on storage.objects for update to authenticated
  using (bucket_id = 'operator-photos' and public.is_photo_admin())
  with check (bucket_id = 'operator-photos' and public.is_photo_admin()
              and name ~ '^[0-9a-f]{24}\.jpg$');

create policy photos_delete on storage.objects for delete to authenticated
  using (bucket_id = 'operator-photos' and public.is_photo_admin());

-- ------------------------------------------------- make yourself the admin
-- Run this once, with your own sign-in email (the account you will upload with):
--
--   insert into public.photo_admins (id)
--   select id from auth.users where email = 'you@example.com'
--   on conflict do nothing;
--
-- To remove an administrator:  delete from public.photo_admins where id = '...';

-- ------------------------------------------------------------------ checks
-- Run these after the above. Each should return what the comment says.
--
--   select id, public, file_size_limit, allowed_mime_types
--     from storage.buckets where id = 'operator-photos';
--   -- one row, public = false
--
--   select policyname, cmd from pg_policies
--     where schemaname = 'storage' and tablename = 'objects' and policyname like 'photos_%'
--     order by policyname;
--   -- four rows: photos_delete, photos_insert, photos_read, photos_update
--
--   select count(*) from public.photo_admins;
--   -- 1 (you)
