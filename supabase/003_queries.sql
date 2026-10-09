-- Shared "Queried" log: who opened a record, and when.
-- Run once in the Supabase SQL editor on a project that already has schema.sql
-- applied. (A fresh project gets all of this from schema.sql.) Safe to re-run.
--
-- The app writes a row each time an officer opens a vehicle, owner or operator
-- record (at most once per officer per record every 30 minutes). Other officers
-- then see "Queried <date> @ <time> · <name>" alongside the "Stopped" marks,
-- which stay in the checks table.
--
-- Same rules as checks: the name and receive time are stamped by the database
-- from the signed-in account, never taken from the app, and rows can only be
-- added -- nobody can change or erase one from the app.

create table if not exists public.queries (
  id          uuid primary key default gen_random_uuid(),
  rec_key     text not null check (length(rec_key) between 1 and 200),
  rec_label   text check (length(rec_label) <= 200),
  written_at  timestamptz not null default now(),
  created_at  timestamptz not null default now(),
  author_id   uuid,
  author_name text not null default ''
);
create index if not exists queries_rec_key_idx on public.queries (rec_key);
create index if not exists queries_created_idx on public.queries (created_at);
alter table public.queries enable row level security;

drop policy if exists queries_read   on public.queries;
drop policy if exists queries_insert on public.queries;
create policy queries_read   on public.queries for select to authenticated using (public.is_pvh_user());
create policy queries_insert on public.queries for insert to authenticated with check (public.is_pvh_user());

create or replace function public.queries_guard() returns trigger
language plpgsql security definer set search_path = public as $$
declare who text;
begin
  if auth.uid() is null then return new; end if;
  select display_name into who from public.profiles where id = auth.uid();
  new.author_id   := auth.uid();
  new.author_name := coalesce(who, '');
  new.created_at  := now();
  new.written_at  := least(coalesce(new.written_at, now()), now());
  return new;
end $$;
drop trigger if exists queries_guard_trg on public.queries;
create trigger queries_guard_trg before insert on public.queries
  for each row execute function public.queries_guard();

revoke all on public.queries from anon;
grant select, insert on public.queries to authenticated;
