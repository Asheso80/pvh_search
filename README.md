# PVH Field Lookup

A phone-first lookup tool for passenger-vehicle-for-hire records. Search by deck
light number, plate, name, business or licence number and get the vehicle or
operator back with its expiry status (permits, insurance, licences, inspection).
Works offline once the data is loaded.

Two things get built from the same source:

| Output | What it is |
|---|---|
| `PVH_Field_App.html` | One self-contained file with the data baked in. Never goes on GitHub. |
| `docs/` | The installable app (PWA) published to GitHub Pages. Holds **no data**. |

---

## Data handling

**The build produces a data file containing personal information. It is never
committed.** `.gitignore` blocks the data file, the self-contained app and
spreadsheets, and commit hooks in `githooks/` refuse a commit that contains
personal information from the current data. Turn them on once per copy of the repo:

```
git config core.hooksPath githooks
```

`python githooks/pii_check.py all` checks everything already committed.

`docs/` is the only published part of this repo and contains no records. It does
hold `geo.bin`, the public NS Civic Address File for CBRM (roads and civic
addresses), licensed under the Open Government Licence - Nova Scotia. It is public
map data, not PVH data.

How the data reaches devices is documented in a private operations note, not here.

---

## Building

Put the source spreadsheet exports next to `Build_PVH_App.bat` and double-click it.
The build tells you if a required file is missing or is the wrong report.

It produces `PVH_Field_App.html` and refreshes `docs/`. It also prints a data
quality report and a diff against the previous build. Worth reading each time.

To run it directly:

```
python build_pvh_field_app.py <vehicles.xlsx> <operators.xlsx> [addresses.xlsx] [output.html]
```

---

## Updating the app

Push to `main`. GitHub Pages rebuilds from `docs/` within a minute or two, and each
device picks up the new app the next time it opens with a signal. With no signal it
runs from its cached copy.

To confirm a device is current, check the version at the bottom of the home screen.
Bump `APP_VERSION` in the template whenever you change the app.

---

## Editing the app

`build_pvh_field_app.py` holds the entire app as a template string (HTML, CSS and
JavaScript). **`docs/index.html` is generated from it**, so an edit made only to
`docs/index.html` is overwritten by the next build. Change the template, then run a
build.

Testing locally: run `Start_Test_Server.bat`, which serves the app with made-up
records and a throwaway shared-notes server that forgets everything when closed.
Test sign-in details are printed by the script. Do not test against the live
shared-notes server.

---

## Features

- **Search and status.** Green is valid, amber expires within 30 days, red is
  expired or missing. A flagged-records screen collects everything overdue.
- **Empty results say so.** A search that finds nothing is reported explicitly,
  because a record not in the active list is worth verifying rather than clearing.
- **Shared notes.** Officers can add notes and actions to records, with author and
  time stamped. Nothing is overwritten or erased. Works offline and syncs later.
- **Tasks.** Follow-ups on vehicles and operators with a due date, priority and an
  assigned officer, plus a Tasks list and optional email reminders.
- **Location sheet.** Look up roads and civic addresses from the public map data.

Setup for the shared features is in the private operations note.

---

## Intel export (PC only)

A separate PC-only tool builds relationship workbooks and maps from the shared
activity. Its output goes to a gitignored folder and is never published. See the
private operations note.
