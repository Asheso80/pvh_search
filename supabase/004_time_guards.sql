-- Time guards: notes, stops and queries cannot be back-dated.
-- Run once in the Supabase SQL editor on a project that already has schema.sql
-- (and 002, 003) applied. A fresh project gets all of this from schema.sql.
-- Safe to re-run.
--
-- The phone supplies "written at" -- when the officer wrote the note, made
-- the stop or opened the record -- because something written with no signal
-- is sent later and should keep the time it happened. Until now the database
-- only refused a future time. It now also refuses one more than 72 hours
-- before the server receives it (long enough for a weekend without signal),
-- and the time a note is marked done gets the same rule (it had none).
--
-- A time outside the window is moved to the nearest edge, not refused, so
-- nothing an officer wrote is lost. The server's own receive time
-- (created_at / updated_at) is still stamped separately and cannot be set
-- from the app.

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
