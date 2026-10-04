#!/usr/bin/env python3
"""
refresh_geo.py -- refresh the CBRM map data behind the app's location sheet.

WHAT IT DOES
    Downloads the current CBRM civic points and addressable roads from the Nova
    Scotia Open Data portal, builds the lookup bundle the app uses, and writes
    it to docs/geo.bin (gzipped JSON). Nothing else is touched: the app's code is
    not rebuilt. After you commit and push docs/geo.bin, each phone picks up the
    new map the next time it opens with a signal (the service worker's map cache
    is named from the file's contents).

HOW TO RUN
    python refresh_geo.py
    python refresh_geo.py --local points.csv roads.csv     (already downloaded)
    python refresh_geo.py --force                           (skip the size check)

    Needs Python 3 and shapely (pip install shapely). 'requests' is optional.

THEN
    Check the printed summary against the previous file, then:
        git add docs/geo.bin
        git commit -m "Refresh CBRM map data"
        git push

SAFETY
    - The old docs/geo.bin is replaced only after the new one is built in full.
    - If the new data is much smaller than the old (a bad or partial download),
      the script stops and leaves the old file alone. --force overrides that.

Data: Nova Scotia Civic Address File (Civic Points tntn-er5g, Addressable Roads
xtdd-axm7), licensed under the Open Government Licence - Nova Scotia.
"""

import csv, datetime, gzip, io, json, math, os, sys, tempfile

POINTS_URL = "https://data.novascotia.ca/resource/tntn-er5g.csv?$where=mun='CBRM'&$limit=500000"
ROADS_URL = "https://data.novascotia.ca/resource/xtdd-axm7.csv?$where=mun='CBRM'&$limit=500000"

SIMPLIFY_TOL = 0.00002   # ~2 m
GRID_CELL = 0.005        # ~0.5 km index cells; the app reads this back from the file
NODE_SNAP = 6            # endpoint rounding so coincident intersections merge
MIN_RATIO = 0.8          # refuse a new file with under 80% of the old one's records

OUT_PATH = os.path.join("docs", "geo.bin")


def fetch(url):
    try:
        import requests
        r = requests.get(url, timeout=180)
        r.raise_for_status()
        return r.text
    except ImportError:
        from urllib.request import Request, urlopen
        req = Request(url, headers={"User-Agent": "pvh-geo-refresh"})
        with urlopen(req, timeout=180) as resp:
            return resp.read().decode("utf-8")


def get_sources(args):
    if args and args[0] == "--local":
        if len(args) < 3:
            sys.exit("Usage: python refresh_geo.py --local <points.csv> <roads.csv>")
        print(f"Reading {args[1]} and {args[2]}")
        with open(args[1], encoding="utf-8") as f:
            pts = f.read()
        with open(args[2], encoding="utf-8") as f:
            rds = f.read()
        return pts, rds
    print("Downloading civic points from NS Open Data ...")
    pts = fetch(POINTS_URL)
    print("Downloading addressable roads from NS Open Data ...")
    rds = fetch(ROADS_URL)
    return pts, rds


def cell_key(lon, lat):
    return f"{math.floor(lon / GRID_CELL)}:{math.floor(lat / GRID_CELL)}"


def street_name(row):
    parts = [row.get("strprefix", ""), row.get("strname", ""),
             row.get("strsuffix", ""), row.get("strdir", "")]
    return " ".join(p for p in parts if p and p.strip())


def build_bundle(points_csv, roads_csv):
    from shapely import wkt as shapely_wkt
    from shapely.geometry import MultiLineString

    points = []
    for row in csv.DictReader(io.StringIO(points_csv)):
        try:
            lon = float(row["long"])
            lat = float(row["lat"])
        except (ValueError, KeyError):
            continue
        points.append({
            "n": row["civicnum"] + (row.get("civsuffix") or ""),
            "s": street_name(row),
            "c": row.get("comm", ""),
            "lon": round(lon, 6),
            "lat": round(lat, 6),
        })

    roads = []
    node_streets = {}
    unnamed = 0
    for row in csv.DictReader(io.StringIO(roads_csv)):
        street = (row.get("street") or "").strip()
        geom = row.get("the_geom", "")
        if not geom:
            continue
        try:
            g = shapely_wkt.loads(geom)
        except Exception:
            continue
        for line in (g.geoms if isinstance(g, MultiLineString) else [g]):
            simp = line.simplify(SIMPLIFY_TOL, preserve_topology=False)
            coords = [(round(x, 6), round(y, 6)) for x, y in simp.coords]
            if len(coords) < 2:
                continue
            roads.append({
                "id": len(roads), "s": street,
                "fl": row.get("froml", ""), "tl": row.get("tol", ""),
                "fr": row.get("fromr", ""), "tr": row.get("tor", ""),
                "cls": row.get("roadclass", ""), "g": coords,
            })
            if not street:
                unnamed += 1
            else:
                for x, y in (coords[0], coords[-1]):
                    node_streets.setdefault((round(x, NODE_SNAP), round(y, NODE_SNAP)), set()).add(street)

    pt_grid = {}
    for i, p in enumerate(points):
        pt_grid.setdefault(cell_key(p["lon"], p["lat"]), []).append(i)

    rd_grid = {}
    for i, r in enumerate(roads):
        for x, y in r["g"]:
            rd_grid.setdefault(cell_key(x, y), set()).add(i)
    rd_grid = {k: sorted(v) for k, v in rd_grid.items()}

    nodes = [{"lon": x, "lat": y, "st": sorted(s)}
             for (x, y), s in node_streets.items() if len(s) >= 2]

    return {
        "meta": {
            "source": "Nova Scotia Civic Address File (Civic Points + Addressable Roads)",
            "licence": "Contains information licensed under the Open Government Licence - Nova Scotia",
            "municipality": "CBRM",
            "generated": datetime.date.today().isoformat(),
            "counts": {"points": len(points), "roads": len(roads),
                       "intersections": len(nodes), "unnamed_road_segments": unnamed},
            "grid_cell_deg": GRID_CELL, "simplify_tol_deg": SIMPLIFY_TOL,
        },
        "points": points, "roads": roads, "nodes": nodes,
        "pt_grid": pt_grid, "rd_grid": rd_grid,
    }


def old_meta(path):
    try:
        with open(path, "rb") as f:
            return json.loads(gzip.decompress(f.read()))["meta"]
    except Exception:
        return None


def write_atomic(path, data):
    """Temp file next to the target, then rename over it. The project lives in
    OneDrive, where truncating a synced file in place can fail at random."""
    folder = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(dir=folder, prefix="._geo_", suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def main():
    args = [a for a in sys.argv[1:] if a != "--force"]
    force = "--force" in sys.argv[1:]
    if not os.path.isdir("docs"):
        sys.exit("Run this from the project folder (the one that contains docs/).")

    pts_text, rds_text = get_sources(args)
    print("Building the lookup bundle ...")
    bundle = build_bundle(pts_text, rds_text)
    new = bundle["meta"]["counts"]
    old = old_meta(OUT_PATH)

    print()
    print(f"{'':16}{'previous':>12}{'new':>12}")
    for key in ("points", "roads", "intersections"):
        before = old["counts"][key] if old else "-"
        print(f"{key:16}{before:>12}{new[key]:>12}")
    print(f"{'data date':16}{(old or {}).get('generated', '-'):>12}{bundle['meta']['generated']:>12}")

    if new["points"] == 0 or new["roads"] == 0:
        sys.exit("\nSTOPPED: the download produced no points or no roads. The old file is untouched.")
    if old and not force:
        for key in ("points", "roads"):
            if new[key] < old["counts"][key] * MIN_RATIO:
                sys.exit(f"\nSTOPPED: new {key} count is under {int(MIN_RATIO * 100)}% of the old one, "
                         "which looks like a partial download. The old file is untouched. "
                         "Use --force if the drop is real.")

    raw = json.dumps(bundle, separators=(",", ":")).encode("utf-8")
    gz = gzip.compress(raw, 9)
    write_atomic(OUT_PATH, gz)
    print(f"\nWrote {OUT_PATH}: {len(raw) / 1048576:.1f} MB raw, {len(gz) / 1048576:.2f} MB gzipped.")
    print("Next: git add docs/geo.bin, commit, push.")


if __name__ == "__main__":
    main()
