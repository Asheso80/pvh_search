# Shared notes and checks -- Supabase setup

Done once by the administrator. Nothing here is a secret except the passwords
you choose for officers; the project URL and public key go in
`pvh_local_config.json`, which is gitignored.

## 1. Create the project

1. Sign in at supabase.com, **New project**.
2. Region: **Canada (Central)**. Pick a strong database password and keep it in
   your password manager -- the app never needs it.

## 2. Create the tables and access rules

**SQL Editor → New query**, paste all of `supabase/schema.sql`, **Run**.
It is safe to run again.

## 3. Lock sign-up

**Authentication → Sign In / Providers** (or **Providers → Email**):

- Turn **off** "Allow new users to sign up".
- Turn **off** "Confirm email" (officers are created by you, not self-registered).

With sign-up off, nobody can create their own account. Even if one existed, the
access rules above give an account nothing until it has a `profiles` row.

## 4. Add an officer

For each person:

1. **Authentication → Users → Add user → Create new user**. Enter their email and
   a password, tick **Auto Confirm User**.
2. In the SQL editor, give them access and the name that appears on their notes:

```sql
insert into public.profiles (id, display_name)
select id, 'Officer Name' from auth.users where email = 'them@example.com'
on conflict (id) do update set display_name = excluded.display_name;
```

To remove someone's access, delete their row from `profiles` (they can still
sign in but see nothing), or delete the user.

## Upgrading an existing project: shared "Queried" log

If the project was set up before the Queried log existed, run
`supabase/003_queries.sql` once in the SQL editor. It is safe to run again. A
fresh project does not need it: `schema.sql` already includes it.

It adds a `queries` table: one row each time an officer opens a record (at most
once per officer per record every 30 minutes), stamped by the database with the
officer's display name and the time. Until it has run, everything else works;
queries are simply not shared, and the app quietly drops the ones it could not
send rather than flagging them as refused.

## Upgrading an existing project: edit and remove notes

If the project was set up before edit/remove existed, run
`supabase/002_edit_remove.sql` once in the SQL editor **before** phones get the
matching app version. It is safe to run again. A fresh project does not need it:
`schema.sql` already includes it.

Until it has run, notes still work, but an edit or removal is refused by the
database and shows "A change to this note was not accepted" with a Dismiss link.

What it guarantees:

- Only the author can edit or remove their own note. Anyone signed in can still
  tick it done.
- Nothing is overwritten or erased. Each edit keeps the previous text in
  `note_events` with who and when; a removal keeps the text and marks who and when.
  Officers cannot add, change or delete audit entries (there is no policy that
  allows it), and names and times come from the database, not the app.
- To review the history: Table Editor → `note_events`, or
  `select * from public.note_events order by created_at;`

## Reset an officer's password

Officer logins need not be real mailboxes, so the dashboard's emailed reset link
will not reach anyone. Set the password directly in the SQL editor:

```sql
update auth.users
set encrypted_password = crypt('NewPasswordHere', gen_salt('bf')),
    updated_at = now()
where email = 'officer@example.com';
```

Keep the `where` line: without it every account gets the same password. It
should report one row. If `crypt` is not found, use `extensions.crypt` and
`extensions.gen_salt`. Their notes are unaffected.

To cut someone off quickly, delete their row from `public.profiles`.

## 5. Point the app at it

**Project Settings → API**: copy the **Project URL** and the **anon / publishable
key**. Add to `pvh_local_config.json`:

```json
"supabase_url": "https://xxxxxxxx.supabase.co",
"supabase_anon_key": "eyJ..."
```

Rebuild. The URL and key travel to devices inside `PVH_data.json` (the private
data file), not in the published app, so the public repo never contains them. The
anon key is designed to be exposed to clients; what protects the data is the
access rules in step 2. **Never put the `service_role` key anywhere in the app.**

Each officer then opens the app, taps **Shared notes** at the bottom of the home
screen, and signs in once. The device stays signed in.

## Operating notes

- The free plan pauses a project after about a week with no activity. Daily
  field use keeps it awake; if it ever pauses, restore it from the dashboard.
- Officers can edit or remove their own notes from the app, never anyone else's.
  Nothing is erased: see the edit and removal history above. A note by someone who
  has left can only be corrected by you, in **Table Editor → notes**.
- Free-plan projects have no point-in-time backups. Export `notes` and `checks`
  from the Table Editor now and then if the history matters.
