-- PVH task emails -- run once in the Supabase SQL editor, after 008_tasks.sql.
-- Safe to run again.
--
-- Two emails, both sent by the task-notify edge function (supabase/functions/
-- task-notify/index.ts) through Resend:
--   assigned  when a task is given to someone (new, or reassigned) -- at once.
--   digest    weekdays at 08:00 and 14:30 Halifax time, to each officer who has
--             a task overdue or due today: counts and a link, nothing else.
-- Nothing about a plate, a person or a task title goes in an email.
--
-- The database tells the function what to do (pg_net for the assignment, pg_cron
-- for the schedule). A failed email never blocks a task: the triggers swallow
-- errors, and pg_net sends in the background.
--
-- After running this, two more things are needed -- see SETUP.md, "Task emails":
--   1. insert the function's address and a shared secret into notify_config
--   2. set each officer's notify_email in profiles

create extension if not exists pg_net with schema extensions;
create extension if not exists pg_cron;

-- Where an officer's emails go. Officer logins need not be real mailboxes, so
-- this is separate. Set by the administrator in the SQL editor; officers cannot
-- change it (profiles has no update policy).
alter table public.profiles add column if not exists notify_email text;
alter table public.profiles drop constraint if exists profiles_notify_email_shape;
alter table public.profiles add constraint profiles_notify_email_shape
  check (notify_email is null or notify_email ~* '^[^@[:space:]]+@[^@[:space:]]+\.[^@[:space:]]+$');

-- The function's address and the secret it checks. Readable by no officer and
-- no phone: only the two functions below, which run as their owner.
create table if not exists public.notify_config (
  key   text primary key,
  value text not null
);
alter table public.notify_config enable row level security;
revoke all on public.notify_config from anon, authenticated;

-- The email function reads tasks and each officer's notify_email with the service
-- role. Newer projects do not give that role access to tables made in the SQL
-- editor, so grant read, and only read, on the two it needs.
grant select on public.tasks to service_role;
grant select on public.profiles to service_role;

-- ------------------------------------------------------------- assignment
create or replace function public.tasks_notify() returns trigger
language plpgsql security definer set search_path = public as $$
declare fn text; sec text;
begin
  if new.status <> 'open' then return null; end if;
  -- Giving a task to yourself is not news.
  if new.assignee_id = new.author_id then return null; end if;
  if tg_op = 'UPDATE' and new.assignee_id is not distinct from old.assignee_id then return null; end if;

  select value into fn  from public.notify_config where key = 'function_url';
  select value into sec from public.notify_config where key = 'secret';
  if fn is null or sec is null then return null; end if;

  perform net.http_post(
    url     := fn,
    headers := jsonb_build_object('Content-Type', 'application/json', 'x-notify-secret', sec),
    body    := jsonb_build_object('kind', 'assigned', 'task_id', new.id));
  return null;
exception when others then
  return null;
end $$;
revoke all on function public.tasks_notify() from public, anon, authenticated;
drop trigger if exists tasks_notify_trg on public.tasks;
create trigger tasks_notify_trg after insert or update of assignee_id on public.tasks
  for each row execute function public.tasks_notify();

-- ----------------------------------------------------------------- digest
-- Called by pg_cron, which runs in UTC. Halifax is UTC-3 in summer and UTC-4 in
-- winter, so the job fires at both possible UTC hours and this function lets
-- through only the one that is really 08:00 or 14:30 in Halifax. A weekend or
-- the wrong hour does nothing.
create or replace function public.notify_digest_tick() returns void
language plpgsql security definer set search_path = public as $$
declare fn text; sec text; loc timestamp := (now() at time zone 'America/Halifax'); slot text;
begin
  if extract(isodow from loc) > 5 then return; end if;
  if extract(hour from loc) = 8 and extract(minute from loc) < 5 then
    slot := 'morning';
  elsif extract(hour from loc) = 14 and extract(minute from loc) between 30 and 34 then
    slot := 'afternoon';
  else
    return;
  end if;

  select value into fn  from public.notify_config where key = 'function_url';
  select value into sec from public.notify_config where key = 'secret';
  if fn is null or sec is null then return; end if;

  perform net.http_post(
    url     := fn,
    headers := jsonb_build_object('Content-Type', 'application/json', 'x-notify-secret', sec),
    body    := jsonb_build_object('kind', 'digest', 'slot', slot));
exception when others then
  return;
end $$;
revoke all on function public.notify_digest_tick() from public, anon, authenticated;

-- 08:00 is 11:00 UTC (summer) or 12:00 UTC (winter); 14:30 is 17:30 or 18:30.
select cron.unschedule(jobid) from cron.job where jobname in ('pvh-digest-morning', 'pvh-digest-afternoon');
select cron.schedule('pvh-digest-morning',   '0 11,12 * * 1-5',  'select public.notify_digest_tick()');
select cron.schedule('pvh-digest-afternoon', '30 17,18 * * 1-5', 'select public.notify_digest_tick()');
