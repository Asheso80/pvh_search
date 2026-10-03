-- PVH shared notes and checks -- run once in the Supabase SQL editor.
--
-- Access model: only people with a row in `profiles` can read or write
-- anything. Being able to sign in is not enough on its own, so a stray
-- account (or a sign-up that slipped through) sees nothing.
--
-- Authorship and timestamps are set by the database, not the app, so an
-- officer cannot post as someone else or back-date a note beyond "now".

-- ---------------------------------------------------------------- profiles
create table if not exists public.profiles (
  id           uuid primary key references auth.users(id) on delete cascade,
  display_name text not null check (length(trim(display_name)) between 1 and 60)
);
alter table public.profiles enable row level security;

drop policy if exists profiles_read_own on public.profiles;
create policy profiles_read_own on public.profiles
  for select to authenticated using (id = auth.uid());
-- No insert/update/delete policy on purpose: people are added and removed by
-- the administrator in the SQL editor, never from the app.

create or replace function public.is_pvh_user() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.profiles where id = auth.uid());
$$;
revoke all on function public.is_pvh_user() from public, anon;
grant execute on function public.is_pvh_user() to authenticated;

-- ------------------------------------------------------------------- notes
create table if not exists public.notes (
  id            uuid primary key default gen_random_uuid(),
  rec_key       text not null check (length(rec_key) between 1 and 200),
  rec_type      text not null check (rec_type in ('vehicle','operator','owner')),
  rec_label     text check (length(rec_label) <= 200),
  body          text not null check (length(trim(body)) between 1 and 2000),
  action_needed boolean not null default false,
  written_at    timestamptz not null default now(),  -- when the officer wrote it
  created_at    timestamptz not null default now(),  -- when the server got it
  updated_at    timestamptz not null default now(),
  author_id     uuid,
  author_name   text not null default '',
  done_at       timestamptz,
  done_by       text
);
create index if not exists notes_rec_key_idx on public.notes (rec_key);
create index if not exists notes_updated_idx on public.notes (updated_at);
alter table public.notes enable row level security;

drop policy if exists notes_read   on public.notes;
drop policy if exists notes_insert on public.notes;
drop policy if exists notes_update on public.notes;
create policy notes_read   on public.notes for select to authenticated using (public.is_pvh_user());
create policy notes_insert on public.notes for insert to authenticated with check (public.is_pvh_user());
create policy notes_update on public.notes for update to authenticated
  using (public.is_pvh_user()) with check (public.is_pvh_user());
-- No delete policy: a note is a record. The administrator can remove one in
-- the dashboard if it was posted by mistake.

create or replace function public.notes_guard() returns trigger
language plpgsql security definer set search_path = public as $$
declare who text;
begin
  -- auth.uid() is null for the administrator working in the SQL editor,
  -- who is allowed to correct anything.
  if auth.uid() is null then
    if tg_op = 'UPDATE' then new.updated_at := now(); end if;
    return new;
  end if;
  select display_name into who from public.profiles where id = auth.uid();
  if tg_op = 'INSERT' then
    new.author_id   := auth.uid();
    new.author_name := coalesce(who, '');
    new.created_at  := now();
    new.updated_at  := now();
    new.written_at  := least(coalesce(new.written_at, now()), now());
    new.done_at     := null;
    new.done_by     := null;
  else
    -- Officers may only tick a note done or reopen it.
    if new.id <> old.id or new.rec_key <> old.rec_key or new.rec_type <> old.rec_type
       or new.rec_label is distinct from old.rec_label or new.body <> old.body
       or new.action_needed <> old.action_needed or new.written_at <> old.written_at
       or new.created_at <> old.created_at or new.author_id is distinct from old.author_id
       or new.author_name <> old.author_name then
      raise exception 'Only the done flag of a note can be changed';
    end if;
    new.done_by    := case when new.done_at is not null then coalesce(who, '') else null end;
    new.updated_at := now();
  end if;
  return new;
end $$;
drop trigger if exists notes_guard_trg on public.notes;
create trigger notes_guard_trg before insert or update on public.notes
  for each row execute function public.notes_guard();

-- ------------------------------------------------------------------ checks
create table if not exists public.checks (
  id          uuid primary key default gen_random_uuid(),
  rec_key     text not null check (length(rec_key) between 1 and 200),
  rec_label   text check (length(rec_label) <= 200),
  written_at  timestamptz not null default now(),
  created_at  timestamptz not null default now(),
  author_id   uuid,
  author_name text not null default ''
);
create index if not exists checks_rec_key_idx on public.checks (rec_key);
create index if not exists checks_created_idx on public.checks (created_at);
alter table public.checks enable row level security;

drop policy if exists checks_read   on public.checks;
drop policy if exists checks_insert on public.checks;
create policy checks_read   on public.checks for select to authenticated using (public.is_pvh_user());
create policy checks_insert on public.checks for insert to authenticated with check (public.is_pvh_user());

create or replace function public.checks_guard() returns trigger
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
drop trigger if exists checks_guard_trg on public.checks;
create trigger checks_guard_trg before insert on public.checks
  for each row execute function public.checks_guard();

-- ------------------------------------------------------------------ grants
revoke all on public.profiles, public.notes, public.checks from anon;
grant select on public.profiles to authenticated;
grant select, insert, update on public.notes to authenticated;
grant select, insert on public.checks to authenticated;
