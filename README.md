# PVH Field Lookup

A phone-first lookup tool for CBRM Passenger Vehicle for Hire records. Type a deck
light number, plate, name, business or licence number and get the vehicle or
operator back with its expiry status — permits, insurance, licences, MVI — working
offline once the data is loaded.

Two things get built from the same source:

| Output | What it is |
|---|---|
| `PVH_Field_App.html` | One self-contained file with the data baked in. Never goes on GitHub. |
| `docs/` | The installable app (PWA) published to GitHub Pages. Holds **no data** — each device loads that separately. |

---

## Data handling

**The build produces a data file containing personal information. It is never
committed.** `.gitignore` blocks `PVH_data.json`, `PVH_Field_App.html` and `*.xlsx`,
but treat that as a backstop, not the rule.

`docs/` is the only part of this repo that is published, and it contains no records. It does hold `geo.bin`, the public NS Civic Address File for CBRM (roads and civic addresses) used by the location sheet, licensed under the Open Government Licence - Nova Scotia. It is public map data, not PVH data.
Payment, refund, SAP and criminal-check fields are dropped at build time and never
reach the app at all.

Distribution of the data file — where it is kept and how devices get it — is
deliberately not documented here. See the operations note held outside this repo.

---

## Building

Put these next to `Build_PVH_App.bat` and double-click it:

- `All_Active_Vehicles.xlsx` — required. Must be the *All Active Vehicles* report;
  the owner-keyed "Active Owners & Active Vehicles" export lacks the insurance, MVI
  and licence expiry columns, and the build will tell you so if you use it.
- `OperatorList.xlsx` — required.
- `All_Vehicle_LastInspection.xlsx` — optional; adds owner mailing addresses.

It produces `PVH_Field_App.html`, refreshes `docs/`, and writes the data file.

The build also prints a **data quality report** — records missing plates, VINs,
insurance or expiry dates — and a **diff against the previous build**, listing
vehicles and operators added or removed and anything newly expired. Worth reading
each time; it is the cheapest audit of the source exports you will get.

To run it directly instead of via the `.bat`:

```
python build_pvh_field_app.py All_Active_Vehicles.xlsx OperatorList.xlsx [addresses.xlsx] [output.html]
```

---

## Updating the app

Push to `main`. GitHub Pages rebuilds from `docs/` within a minute or two, and each
device picks up the new app the next time it opens with a signal. With no signal it
runs from its cached copy, so the field never depends on this.

To confirm a device is current, look at the bottom of the app's home screen:

```
App version 2026-08-19.2
```

Bump `APP_VERSION` in the template whenever you change the app, so there is
something to check against.

---

## Editing the app

`build_pvh_field_app.py` holds the entire app as a template string — HTML, CSS and
JavaScript. **`docs/index.html` is generated from it**, so an edit made only to
`docs/index.html` gets overwritten by the next build.

Change the template, then either run a build, or make the identical edit in both
files. To verify they still agree:

```python
import ast, io
src = io.open("build_pvh_field_app.py", encoding="utf-8").read()
vals = {t.id: n.value.value for n in ast.parse(src).body
        if isinstance(n, ast.Assign) and isinstance(n.value, ast.Constant)
        and isinstance(n.value.value, str)
        for t in n.targets if isinstance(t, ast.Name)}
shell = (vals["TEMPLATE"].replace("__HEAD_EXTRA__", vals["SHELL_HEAD"])
         .replace("__DATA_SCRIPT__", "").replace("__BOOT__", vals["SHELL_BOOT"]))
print(shell == io.open("docs/index.html", encoding="utf-8").read())
```

Testing locally: serve the repo folder (`py -m http.server 8000`) and open
`http://localhost:8000/docs/`. Opening `docs/index.html` straight off disk will not
work — a `file://` page cannot register a service worker.

---

## Shared notes and checks

Officers can write a note on any vehicle, owner or operator record, and mark an
action as needed or done. Deliberate "Checked" marks are shared too, so the check
history shows who checked a record and when. Both live in a Supabase project;
**nothing else is shared** -- the records, search, and the passing "seen" history
stay on the device.

- A note appears only on the record it was written on, never on a related record.
- A record with notes shows an amber **NOTES n** chip in search results and a strip
  on the record; an unresolved action turns both red (**ACTION NEEDED**) and puts
  the record on a home-screen list.
- Author and time are stamped by the database from the officer's sign-in, not by
  the app.
- An officer can edit or remove their own notes, never anyone else's. Nothing is
  overwritten or erased: edits keep the earlier wording, and a removed note stays on
  file showing who removed it and when, with the original a tap away.
- Offline: the last shared copy is kept on the device, and anything written or
  marked without a signal waits in an outbox and is sent when there is one.
- It is off unless `supabase_url` and `supabase_anon_key` are in
  `pvh_local_config.json`. They travel in `PVH_data.json`, never in `docs/`.

Setup (database, access rules, adding officers): `supabase/SETUP.md`.

Each device holds a copy of the shared notes while signed in, in addition to the
records. Signing out removes it.

---

## What the app shows

Where each status badge comes from in the source export:

| Badge | Column |
|---|---|
| PVH License | `NSVehicle Permit Expiry` |
| Insurance | `Insurance Expiry` |
| Veh licence | `Expiry Date` |
| MVI due | `First MVIDate` + 1 year |
| Op licence | `Renewal Date` |
| NS DL | `NSDLExpired Date` |

Green is valid, amber expires within 30 days, red is expired or missing. The
flagged-records screen collects everything overdue, filterable by category.

A search that finds nothing says so explicitly rather than showing an empty screen —
a plate that is not in the active list may be suspended, transferred or never
licensed, and that is worth verifying rather than clearing.
