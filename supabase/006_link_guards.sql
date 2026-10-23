-- Link guards: tightens the operator-vehicle links from 005_links.sql.
-- Run once in the Supabase SQL editor, after 005. Safe to re-run. A fresh
-- project gets all of this from schema.sql. Running 005 again later puts the
-- old rules back, so run this again after it. No app update is needed.
--
-- 1. One row per pair, enforced. The app makes a link's id from its two record
--    keys ("l|o:...|v:..."); the database now insists on it, so a hand-made
--    request cannot file the same operator and vehicle under a second id. Only
--    rows written from now on are checked; existing ones are left as they are.
--
-- 2. No quiet relabelling. A link's labels (the operator's name and the
--    vehicle's description, shown when the record is not in the current data)
--    could be changed by any officer with nothing on file. Now they change only
--    when the link is added or removed, and that is recorded in link_events with
--    who did it. Any other change to them is ignored.

alter table public.links drop constraint if exists links_id_shape;
alter table public.links add constraint links_id_shape check (
  op_key like 'o:%' and veh_key like 'v:%'
  and position('|' in op_key) = 0 and position('|' in veh_key) = 0
  and id = 'l|' || op_key || '|' || veh_key
) not valid;

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
    -- Fixed for good once written.
    if new.id <> old.id or new.op_key <> old.op_key or new.veh_key <> old.veh_key then
      raise exception 'That part of a link cannot be changed';
    end if;
    new.created_at   := old.created_at;
    new.created_by   := old.created_by;
    new.created_name := old.created_name;
    -- Labels may follow a renamed record, but only on an add or removal, which
    -- link_events records with the new labels and who made the change.
    if new.active is distinct from old.active then
      new.updated_at   := now();
      new.updated_by   := auth.uid();
      new.updated_name := who;
    else
      new.op_label     := old.op_label;
      new.veh_label    := old.veh_label;
      new.updated_at   := old.updated_at;
      new.updated_by   := old.updated_by;
      new.updated_name := old.updated_name;
    end if;
  end if;
  return new;
end $$;
