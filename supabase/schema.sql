-- PVH shared notes and checks -- run once in the Supabase SQL editor.
--
-- Access model: only people with a row in `profiles` can read or write
-- anything. Being able to sign in is not enough on its own, so a stray
-- account (or a sign-up that slipped through) sees nothing.
--
-- Authorship and timestamps are set by the database, not the app, so an
-- officer cannot post as someone else. The time a note, stop or query was
-- written comes from the phone (it may have been written with no signal) and
-- is held to the 72 hours before the server receives it -- never later, never
-- earlier -- see 004_time_guards.sql.

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
  done_by       text,
  edited_at     timestamptz,
  edited_by     text,
  removed_at    timestamptz,
  removed_by    text
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

create table if not exists public.note_events (
  id         uuid primary key default gen_random_uuid(),
  note_id    uuid not null references public.notes(id),
  kind       text not null check (kind in ('edit','remove')),
  old_body   text,
  new_body   text,
  by_name    text not null default '',
  author_id  uuid,
  created_at timestamptz not null default now()
);
create index if not exists note_events_note_idx on public.note_events (note_id);
create index if not exists note_events_created_idx on public.note_events (created_at);
alter table public.note_events enable row level security;
drop policy if exists note_events_read on public.note_events;
create policy note_events_read on public.note_events
  for select to authenticated using (public.is_pvh_user());
-- No insert/update/delete policy: events are written only by the trigger below,
-- so nobody can add, change or erase an entry from the app.

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
  who := coalesce(who, '');
  if tg_op = 'INSERT' then
    new.author_id   := auth.uid();
    new.author_name := who;
    new.created_at  := now();
    new.updated_at  := now();
    new.written_at  := least(greatest(coalesce(new.written_at, now()), now() - interval '72 hours'), now());
    new.done_at     := null;
    new.done_by     := null;
    new.edited_at   := null;
    new.edited_by   := null;
    new.removed_at  := null;
    new.removed_by  := null;
  else
    -- Fixed for good once written.
    if new.id <> old.id or new.rec_key <> old.rec_key or new.rec_type <> old.rec_type
       or new.rec_label is distinct from old.rec_label or new.action_needed <> old.action_needed
       or new.written_at <> old.written_at or new.created_at <> old.created_at
       or new.author_id is distinct from old.author_id or new.author_name <> old.author_name then
      raise exception 'That part of a note cannot be changed';
    end if;

    -- The done flag: anyone signed in may tick or reopen it. The time it was
    -- ticked comes from the phone (it may have been ticked with no signal) but
    -- is held to the same 72-hour window as everything else.
    if new.done_at is distinct from old.done_at then
      if new.done_at is not null then
        new.done_at := least(greatest(new.done_at, now() - interval '72 hours'), now());
      end if;
      new.done_by := case when new.done_at is not null then who else null end;
    else
      new.done_by := old.done_by;
    end if;

    -- Text edits: author only, never once removed, and the old text is kept.
    new.edited_at := old.edited_at;
    new.edited_by := old.edited_by;
    if new.body <> old.body then
      if old.removed_at is not null then
        raise exception 'A removed note cannot be edited';
      end if;
      if old.author_id is distinct from auth.uid() then
        raise exception 'Only the author can edit a note';
      end if;
      new.edited_at := now();
      new.edited_by := who;
      insert into public.note_events (note_id, kind, old_body, new_body, by_name, author_id)
        values (old.id, 'edit', old.body, new.body, who, auth.uid());
    end if;

    -- Removal: author only, one way from the app, the text stays on file.
    if new.removed_at is distinct from old.removed_at then
      if old.removed_at is not null or new.removed_at is null then
        raise exception 'A removal cannot be undone from the app';
      end if;
      if old.author_id is distinct from auth.uid() then
        raise exception 'Only the author can remove a note';
      end if;
      new.removed_at := now();
      new.removed_by := who;
      insert into public.note_events (note_id, kind, old_body, by_name, author_id)
        values (old.id, 'remove', old.body, who, auth.uid());
    else
      new.removed_by := old.removed_by;
    end if;

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
  new.written_at  := least(greatest(coalesce(new.written_at, now()), now() - interval '72 hours'), now());
  return new;
end $$;
drop trigger if exists checks_guard_trg on public.checks;
create trigger checks_guard_trg before insert on public.checks
  for each row execute function public.checks_guard();

-- ------------------------------------------------------------------ queries
-- Who opened a record, and when ("Queried"). See 003_queries.sql.
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
  new.written_at  := least(greatest(coalesce(new.written_at, now()), now() - interval '72 hours'), now());
  return new;
end $$;
drop trigger if exists queries_guard_trg on public.queries;
create trigger queries_guard_trg before insert on public.queries
  for each row execute function public.queries_guard();


-- ------------------------------------------------------------------ grants
revoke all on public.profiles, public.notes, public.checks from anon;
grant select on public.profiles to authenticated;
grant select, insert, update on public.notes to authenticated;
grant select, insert on public.checks to authenticated;
revoke all on public.queries from anon;
grant select, insert on public.queries to authenticated;
revoke all on public.note_events from anon;
grant select on public.note_events to authenticated;


-- ------------------------------------------------------------------- links
-- Operator-vehicle links, see 005_links.sql.
create table if not exists public.links (
  id           text primary key check (length(id) between 3 and 420),
  op_key       text not null check (length(op_key) between 1 and 200),
  veh_key      text not null check (length(veh_key) between 1 and 200),
  op_label     text check (length(op_label) <= 200),
  veh_label    text check (length(veh_label) <= 200),
  active       boolean not null default true,
  created_at   timestamptz not null default now(),
  created_by   uuid,
  created_name text not null default '',
  updated_at   timestamptz not null default now(),   -- last add or removal
  updated_by   uuid,
  updated_name text not null default ''
);
create index if not exists links_op_idx on public.links (op_key);
create index if not exists links_veh_idx on public.links (veh_key);
create index if not exists links_updated_idx on public.links (updated_at);
alter table public.links enable row level security;

drop policy if exists links_read   on public.links;
drop policy if exists links_insert on public.links;
drop policy if exists links_update on public.links;
create policy links_read   on public.links for select to authenticated using (public.is_pvh_user());
create policy links_insert on public.links for insert to authenticated with check (public.is_pvh_user());
create policy links_update on public.links for update to authenticated
  using (public.is_pvh_user()) with check (public.is_pvh_user());
-- No delete policy: a link is a record. Removing it sets active = false.

create table if not exists public.link_events (
  id         uuid primary key default gen_random_uuid(),
  link_id    text not null,
  kind       text not null check (kind in ('add','remove')),
  op_label   text,
  veh_label  text,
  by_name    text not null default '',
  author_id  uuid,
  created_at timestamptz not null default now()
);
create index if not exists link_events_link_idx on public.link_events (link_id);
alter table public.link_events enable row level security;
drop policy if exists link_events_read on public.link_events;
create policy link_events_read on public.link_events
  for select to authenticated using (public.is_pvh_user());
-- No insert/update/delete policy: events are written only by the trigger below.

create or replace function public.links_guard() returns trigger
language plpgsql security definer set search_path = public as $$
declare who text;
begin
  -- auth.uid() is null for the administrator working in the SQL editor.
  if auth.uid() is null then
    if tg_op = 'UPDATE' then new.updated_at := now(); end if;
    return new;
  end if;
  select display_name into who from public.profiles where id = auth.uid();
  who := coalesce(who, '');
  if tg_op = 'INSERT' then
    new.active       := true;
    new.created_at   := now();
    new.created_by   := auth.uid();
    new.created_name := who;
    new.updated_at   := now();
    new.updated_by   := auth.uid();
    new.updated_name := who;
  else
    -- Fixed for good once written (the labels may follow a renamed record).
    if new.id <> old.id or new.op_key <> old.op_key or new.veh_key <> old.veh_key then
      raise exception 'That part of a link cannot be changed';
    end if;
    new.created_at   := old.created_at;
    new.created_by   := old.created_by;
    new.created_name := old.created_name;
    if new.active is distinct from old.active then
      new.updated_at   := now();
      new.updated_by   := auth.uid();
      new.updated_name := who;
    else
      new.updated_at   := old.updated_at;
      new.updated_by   := old.updated_by;
      new.updated_name := old.updated_name;
    end if;
  end if;
  return new;
end $$;
drop trigger if exists links_guard_trg on public.links;
create trigger links_guard_trg before insert or update on public.links
  for each row execute function public.links_guard();

-- After the row is really written, so a repeat link that changes nothing
-- leaves no entry.
create or replace function public.links_audit() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if tg_op = 'INSERT' then
    insert into public.link_events (link_id, kind, op_label, veh_label, by_name, author_id)
      values (new.id, 'add', new.op_label, new.veh_label, new.updated_name, new.updated_by);
  elsif new.active is distinct from old.active then
    insert into public.link_events (link_id, kind, op_label, veh_label, by_name, author_id)
      values (new.id, case when new.active then 'add' else 'remove' end,
              new.op_label, new.veh_label, new.updated_name, new.updated_by);
  end if;
  return null;
end $$;
drop trigger if exists links_audit_trg on public.links;
create trigger links_audit_trg after insert or update on public.links
  for each row execute function public.links_audit();

revoke all on public.links, public.link_events from anon;
grant select, insert, update on public.links to authenticated;
grant select on public.link_events to authenticated;
