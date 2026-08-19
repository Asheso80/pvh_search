# PVH Field Lookup

A phone-first lookup tool for CBRM Passenger Vehicle for Hire records. Type a deck
light number, plate, name, business or licence number and get the vehicle or
operator back with its expiry status — permits, insurance, licences, MVI — without
needing a signal once the data is loaded.

Two things get built from the same source:

| Output | What it is |
|---|---|
| `PVH_Field_App.html` | One self-contained file with the data baked in. Email it, open it, done. Never goes on GitHub. |
| `docs/` | The installable app (PWA) published to GitHub Pages. Holds **no data** — each device loads that separately. |

---

## The rule that matters

**The data contains personal information — owner names, addresses, phone numbers.
It must never be committed.**

`.gitignore` blocks the three offenders (`PVH_data.json`, `PVH_Field_App.html`,
`*.xlsx`), but the rule is worth knowing rather than trusting to a file. `docs/` is
the only thing meant to be public, and it is deliberately empty of records.

---

## Building

Put these next to `Build_PVH_App.bat` and double-click it:

- `All_Active_Vehicles.xlsx` — required. Must be the *All Active Vehicles* report;
  the owner-keyed "Active Owners & Active Vehicles" export lacks the insurance, MVI
  and licence expiry columns, and the build will tell you so if you use it.
- `OperatorList.xlsx` — required.
- `All_Vehicle_LastInspection.xlsx` — optional; adds owner mailing addresses.

It produces `PVH_Field_App.html`, refreshes `docs/`, and writes `PVH_data.json`.

The build also prints a **data quality report** — records missing plates, VINs,
insurance or expiry dates — and a **diff against the previous build**, listing
vehicles and operators added or removed and anything newly expired. Worth reading
each time; it is the cheapest audit of the source exports you will get.

To run it directly instead of via the `.bat`:

```
python build_pvh_field_app.py All_Active_Vehicles.xlsx OperatorList.xlsx [addresses.xlsx] [output.html]
```

---

## Getting new data onto the phones

The published app holds no data, so each device fetches it from a link you control.

**Setup, once per device:** open the app → bottom of the home screen → **Data
source…** → paste the Dropbox share link to `PVH_data.json` → **Save and check
now**. Paste the link exactly as Dropbox gives it; the app converts it to a direct
file link itself.

**Every build after that:** overwrite the same file in Dropbox. Each device checks
the link when the app opens (and when you switch back to it, at most every ten
minutes), compares the build stamp against what it already has, and offers *"New
data available — Update now"*. It never swaps data out mid-lookup.

> **Overwrite that file in place. Do not delete it and upload a new one.** Dropbox
> gives a re-uploaded file a brand new link, which silently breaks every device set
> up with the old one.

**The link is the credential.** Anyone holding that URL can download the file,
names and addresses included. Don't paste it into shared threads. If it leaks,
regenerate the link in Dropbox — and re-paste the new one on every device.

A direct link from any other host works too; Dropbox just gets the automatic
conversion. Loading `PVH_data.json` by hand still works as a fallback and needs no
link at all.

---

## Updating the app itself

Push to `main`. GitHub Pages rebuilds from `docs/` within a minute or two, and each
device picks up the new app the next time it opens with a signal. If there is no
signal it runs from its cached copy, so the field never depends on this.

To confirm a device is current, look at the bottom of the home screen:

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
work — a `file://` page cannot register a service worker or fetch data.

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

Payment, refund, SAP and criminal-check fields are dropped at build time and never
reach the app.

---

## When the data source will not load

The **Data source…** screen names the actual failure:

| Message | What it means |
|---|---|
| *the host would not let this app read the file (CORS)* | The host blocks other sites from reading it. Work/school OneDrive and SharePoint do this by policy and cannot be used. |
| *that link gave back a web page, not PVH_data.json* | The link points at a preview page rather than the file. |
| *HTTP 404 — nothing at that link* | Wrong link, or the file was replaced and got a new one. |
| *the link needs a sign-in this app cannot do* | Shared with specific people rather than anyone with the link. |
