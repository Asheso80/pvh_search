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
- Notes cannot be edited or deleted from the app. A mistaken note is removed in
  **Table Editor → notes**.
- Free-plan projects have no point-in-time backups. Export `notes` and `checks`
  from the Table Editor now and then if the history matters.
