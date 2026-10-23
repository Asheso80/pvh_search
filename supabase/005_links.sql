-- PVH shared operator-vehicle links -- run once in the Supabase SQL editor.
-- Safe to run again. A fresh project does not need it: schema.sql includes it.
--
-- A link says "this operator drives this vehicle". Any officer with a profile
-- can add one and any officer can remove one. Nothing is erased: a removed
-- link stays on the row (active = false) and every add and removal is kept in
-- link_events with who and when. Names and times come from the database.
--
-- One row per operator and vehicle pair. The id is made from the two record
-- keys, so two phones linking the same pair cannot make a duplicate, and
-- linking a pair that was removed brings the same row back.

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
