-- PVH tasks -- run once in the Supabase SQL editor, after schema.sql. Safe to run again.
--
-- A task is a follow-up on a vehicle or an operator: a title, a due date, a
-- priority, a person it is assigned to. Unlike a note it is open until someone
-- does it. It is still a record: nothing is deleted, every change is kept in
-- task_events, and names and times are stamped by the database, not the app.
--
-- Who may do what (the administrator, working in the SQL editor, may do anything):
--   create     any officer with a profiles row, for any officer
--   done       the person it is assigned to, or the person who created it
--   reopen     the same two
--   new date   the same two, while it is open
--   priority   the same two, while it is open
--   reassign   the person who created it, while it is open
--   cancel     the person who created it, while it is open
-- The title, description, record and creator are fixed once written.

-- ------------------------------------------------------------------ people
-- The assign-to list. profiles can be read only by its own owner, so this
-- hands every officer the names (and ids) of the others, and nothing else.
create or replace function public.pvh_people() returns table (id uuid, display_name text)
language sql stable security definer set search_path = public as $$
  select p.id, p.display_name from public.profiles p
  where public.is_pvh_user() order by p.display_name;
$$;
revoke all on function public.pvh_people() from public, anon;
grant execute on function public.pvh_people() to authenticated;

-- ------------------------------------------------------------------- tasks
create table if not exists public.tasks (
  id            uuid primary key default gen_random_uuid(),
  rec_key       text not null check (length(rec_key) between 1 and 200),
  rec_type      text not null check (rec_type in ('vehicle','operator')),
  rec_label     text check (length(rec_label) <= 200),
  title         text not null check (length(trim(title)) between 1 and 120),
  description   text check (length(description) <= 2000),
  due_date      date not null,
  priority      text not null default 'normal' check (priority in ('normal','high')),
  status        text not null default 'open' check (status in ('open','done','cancelled')),
  assignee_id   uuid not null,
  assignee_name text not null default '',
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now(),
  author_id     uuid,
  author_name   text not null default '',
  done_at       timestamptz,
  done_by       text,
  outcome       text check (length(outcome) <= 500),
  cancelled_at  timestamptz,
  cancelled_by  text
);
create index if not exists tasks_rec_key_idx  on public.tasks (rec_key);
create index if not exists tasks_updated_idx  on public.tasks (updated_at);
create index if not exists tasks_assignee_idx on public.tasks (assignee_id, status, due_date);
alter table public.tasks enable row level security;

drop policy if exists tasks_read   on public.tasks;
drop policy if exists tasks_insert on public.tasks;
drop policy if exists tasks_update on public.tasks;
create policy tasks_read   on public.tasks for select to authenticated using (public.is_pvh_user());
create policy tasks_insert on public.tasks for insert to authenticated with check (public.is_pvh_user());
create policy tasks_update on public.tasks for update to authenticated
  using (public.is_pvh_user()) with check (public.is_pvh_user());
-- No delete policy: a task is a record. Cancelling sets status = 'cancelled'.

create table if not exists public.task_events (
  id         uuid primary key default gen_random_uuid(),
  task_id    uuid not null references public.tasks(id),
  kind       text not null check (kind in ('reassign','reschedule','priority','done','reopen','cancel')),
  detail     text,
  by_name    text not null default '',
  author_id  uuid,
  created_at timestamptz not null default now()
);
create index if not exists task_events_task_idx on public.task_events (task_id);
alter table public.task_events enable row level security;
drop policy if exists task_events_read on public.task_events;
create policy task_events_read on public.task_events
  for select to authenticated using (public.is_pvh_user());
-- No insert/update/delete policy: events are written only by the trigger below.

create or replace function public.tasks_guard() returns trigger
language plpgsql security definer set search_path = public as $$
declare who text; a_name text; party boolean;
begin
  -- auth.uid() is null for the administrator working in the SQL editor.
  if auth.uid() is null then
    if tg_op = 'UPDATE' then new.updated_at := now(); end if;
    if tg_op = 'INSERT' or new.assignee_id is distinct from old.assignee_id then
      select display_name into a_name from public.profiles where id = new.assignee_id;
      if a_name is not null then new.assignee_name := a_name; end if;
    end if;
    return new;
  end if;

  select display_name into who from public.profiles where id = auth.uid();
  who := coalesce(who, '');

  if tg_op = 'INSERT' then
    select display_name into a_name from public.profiles where id = new.assignee_id;
    if a_name is null then raise exception 'That person cannot be given a task'; end if;
    new.assignee_name := a_name;
    new.author_id     := auth.uid();
    new.author_name   := who;
    new.created_at    := now();
    new.updated_at    := now();
    new.status        := 'open';
    new.done_at := null; new.done_by := null; new.outcome := null;
    new.cancelled_at := null; new.cancelled_by := null;
    return new;
  end if;

  -- Fixed for good once written.
  if new.id <> old.id or new.rec_key <> old.rec_key or new.rec_type <> old.rec_type
     or new.rec_label is distinct from old.rec_label or new.title <> old.title
     or new.description is distinct from old.description
     or new.created_at <> old.created_at
     or new.author_id is distinct from old.author_id or new.author_name <> old.author_name then
    raise exception 'That part of a task cannot be changed';
  end if;
  if old.status = 'cancelled' then
    raise exception 'A cancelled task cannot be changed';
  end if;

  party := (auth.uid() = old.assignee_id or auth.uid() = old.author_id);

  -- Reassign: the creator, while open.
  if new.assignee_id <> old.assignee_id then
    if old.author_id is distinct from auth.uid() then
      raise exception 'Only the person who created a task can reassign it';
    end if;
    if old.status <> 'open' then raise exception 'Only an open task can be reassigned'; end if;
    select display_name into a_name from public.profiles where id = new.assignee_id;
    if a_name is null then raise exception 'That person cannot be given a task'; end if;
    new.assignee_name := a_name;
    insert into public.task_events (task_id, kind, detail, by_name, author_id)
      values (old.id, 'reassign', old.assignee_name || ' to ' || a_name, who, auth.uid());
  else
    new.assignee_name := old.assignee_name;
  end if;

  -- New date and priority: either party, while open.
  if new.due_date <> old.due_date then
    if not party then raise exception 'Only the assignee or the creator can change the date'; end if;
    if old.status <> 'open' then raise exception 'Only an open task can be rescheduled'; end if;
    insert into public.task_events (task_id, kind, detail, by_name, author_id)
      values (old.id, 'reschedule', old.due_date::text || ' to ' || new.due_date::text, who, auth.uid());
  end if;
  if new.priority <> old.priority then
    if not party then raise exception 'Only the assignee or the creator can change the priority'; end if;
    if old.status <> 'open' then raise exception 'Only an open task can be changed'; end if;
    insert into public.task_events (task_id, kind, detail, by_name, author_id)
      values (old.id, 'priority', old.priority || ' to ' || new.priority, who, auth.uid());
  end if;

  -- Status.
  if new.status <> old.status then
    if old.status = 'open' and new.status = 'done' then
      if not party then raise exception 'Only the assignee or the creator can finish a task'; end if;
      -- The time comes from the phone (it may have been ticked with no signal),
      -- held to the 72 hours before the server hears of it, as for notes.
      new.done_at := least(greatest(coalesce(new.done_at, now()), now() - interval '72 hours'), now());
      new.done_by := who;
      new.outcome := nullif(trim(coalesce(new.outcome, '')), '');
      new.cancelled_at := null; new.cancelled_by := null;
      insert into public.task_events (task_id, kind, detail, by_name, author_id)
        values (old.id, 'done', new.outcome, who, auth.uid());
    elsif old.status = 'done' and new.status = 'open' then
      if not party then raise exception 'Only the assignee or the creator can reopen a task'; end if;
      new.done_at := null; new.done_by := null; new.outcome := null;
      new.cancelled_at := null; new.cancelled_by := null;
      insert into public.task_events (task_id, kind, detail, by_name, author_id)
        values (old.id, 'reopen', old.outcome, who, auth.uid());
    elsif old.status = 'open' and new.status = 'cancelled' then
      if old.author_id is distinct from auth.uid() then
        raise exception 'Only the person who created a task can cancel it';
      end if;
      new.done_at := null; new.done_by := null; new.outcome := null;
      new.cancelled_at := now(); new.cancelled_by := who;
      insert into public.task_events (task_id, kind, detail, by_name, author_id)
        values (old.id, 'cancel', null, who, auth.uid());
    else
      raise exception 'That change to a task is not allowed';
    end if;
  else
    new.done_at := old.done_at; new.done_by := old.done_by; new.outcome := old.outcome;
    new.cancelled_at := old.cancelled_at; new.cancelled_by := old.cancelled_by;
  end if;

  new.updated_at := now();
  return new;
end $$;
drop trigger if exists tasks_guard_trg on public.tasks;
create trigger tasks_guard_trg before insert or update on public.tasks
  for each row execute function public.tasks_guard();

revoke all on public.tasks, public.task_events from anon;
grant select, insert, update on public.tasks to authenticated;
grant select on public.task_events to authenticated;
