-- Edit and remove notes, with a full audit trail.
-- Run once in the Supabase SQL editor on a project that already has schema.sql
-- applied. (A fresh project gets all of this from schema.sql.) Safe to re-run.
--
-- Nothing is ever overwritten or erased:
--   * an edit keeps the previous text in note_events, with who and when;
--   * a removal marks the note removed (who, when) and keeps its text;
--   * only the author can edit or remove their own note;
--   * names and times are stamped by the database, never sent by the app.

alter table public.notes add column if not exists edited_at  timestamptz;
alter table public.notes add column if not exists edited_by  text;
alter table public.notes add column if not exists removed_at timestamptz;
alter table public.notes add column if not exists removed_by text;

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
    new.written_at  := least(coalesce(new.written_at, now()), now());
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

    -- The done flag: anyone signed in may tick or reopen it.
    if new.done_at is distinct from old.done_at then
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

revoke all on public.note_events from anon;
grant select on public.note_events to authenticated;
