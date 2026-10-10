#!/usr/bin/env python3
"""
Stops personal information from the PVH data reaching GitHub.

Run by the git hooks in this folder (enabled once with
`git config core.hooksPath githooks`):

    pre-commit   python githooks/pii_check.py staged
    commit-msg   python githooks/pii_check.py message <file>

and by hand to audit everything already committed:

    python githooks/pii_check.py all

What it looks for is read from PVH_data.json each run, so this file holds no
names itself and the list follows the data as it changes:

    a person's full name    "First Last", "First Middle Last", "Last, First",
                            plurals and possessives ("the two First Lasts")
    licence numbers         operator and vehicle (TD0000, "TO 0000")
    plates, VINs, master numbers, phone numbers
    street addresses        "12 Example Street" (number + street name)

A surname or first name on its own is not caught -- too many are ordinary
words -- so keep examples in comments made up. Only lines being added are
checked, not what a file already held.

It also refuses files that must never be committed (the data, the exports,
the intel output, the local config) even if someone force-adds one.

Without PVH_data.json there is nothing to compare against: the check says so
and lets the commit through. A false alarm can be overridden with
`git commit --no-verify`.
"""
import os
import re
import sys
import json
import fnmatch
import subprocess

NEVER_COMMIT = [
    "PVH_data.json", "PVH_Field_App.html", "*.xlsx", "*.xls", "*.csv",
    "intel/*", "logs/*", "pvh_local_config.json", "PVH_operations_note.md",
    "HANDOFF.md", "*.bundle", "snapshot.json",
]

STREET_ABBR = {
    "STREET": "ST", "AVENUE": "AVE", "AV": "AVE", "ROAD": "RD", "DRIVE": "DR",
    "LANE": "LN", "CRESCENT": "CRES", "COURT": "CT", "BOULEVARD": "BLVD",
    "PLACE": "PL", "HIGHWAY": "HWY", "TERRACE": "TERR", "CIRCLE": "CIR",
}
STREET_TYPES = set(STREET_ABBR.values())


def git(*args):
    return subprocess.run(["git"] + list(args), capture_output=True).stdout.decode("utf-8", "replace")


def toks(s):
    return re.sub(r"[^A-Z0-9]+", " ", str(s).upper()).split()


def compact(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


# ---------------------------------------------------------------- what to find
def load_terms(path):
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    names, ids, addrs, phones = {}, {}, {}, {}

    def person(first, middle, last):
        f, m, l = toks(first), toks(middle), toks(last)
        if not f or not l or len("".join(f)) < 2 or len("".join(l)) < 2:
            return
        for seq in (f + l, f + m + l, l + f, l + f + m):
            if len(seq) >= 2:
                names[tuple(seq)] = "a person's name"

    def ident(val, label, minlen=5, need_digit=False):
        c = compact(val)
        if len(c) < minlen or not re.search(r"[A-Z]", c) or (need_digit and not re.search(r"\d", c)):
            return
        ids[c] = label

    def address(line):
        t = [STREET_ABBR.get(x, x) for x in toks(str(line or "").split("·")[0])]
        if len(t) < 2 or not t[0].isdigit() or not re.fullmatch(r"[A-Z]{4,}", t[1]):
            return
        addrs[tuple(t[:2])] = "a street address"

    def phone(val):
        dg = re.sub(r"\D", "", str(val or ""))
        if len(dg) == 11 and dg.startswith("1"):
            dg = dg[1:]
        if len(dg) == 10:
            phones[dg] = phones[dg[3:]] = "a phone number"
        elif len(dg) == 7:
            phones[dg] = "a phone number"

    for v in d.get("vehicles", []):
        person(v.get("Owner First Name"), "", v.get("Owner Last Name"))
        c = compact(v.get("Plate No"))
        if len(c) >= 5 or (len(c) >= 4 and re.search(r"\d", c) and re.search(r"[A-Z]", c)):
            ids[c] = "a plate"
        ident(v.get("Licence No"), "a licence number", 4, need_digit=True)
        ident(v.get("VIN"), "a VIN", 11, need_digit=True)
        address(v.get("Owner Address"))
    for w in d.get("owners", []):
        person(w.get("Owner First Name"), "", w.get("Owner Last Name"))
    for o in d.get("operators", []):
        person(o.get("First Name"), o.get("Middle"), o.get("Last Name"))
        ident(o.get("Licence Number"), "a licence number", 4, need_digit=True)
        m = str(o.get("Master Number") or "").strip().upper()
        m = re.sub(r"^0*4\s+", "", re.sub(r"^4A\b", "", m))
        ident(m, "a master number", 8, need_digit=True)
        address(o.get("Address1"))
        phone(o.get("Phone"))
        phone(o.get("Cell Phone"))
    return names, ids, addrs, phones


def find(text, terms):
    """Every term found in one line of text, as (kind, matched text)."""
    names, ids, addrs, phones = terms
    out = []
    t = toks(text)
    ta = [STREET_ABBR.get(x, x) for x in t]
    n = len(t)
    for i in range(n):
        # names: 2-5 tokens, the last may carry a plural / possessive S
        for k in range(2, 6):
            if i + k > n:
                break
            seq = t[i:i + k]
            alt = seq[:-1] + [seq[-1][:-1]] if len(seq[-1]) > 3 and seq[-1].endswith("S") else None
            if tuple(seq) in names or (alt and tuple(alt) in names):
                out.append((names.get(tuple(seq)) or names[tuple(alt)], " ".join(seq)))
        # identifiers, allowing a space or hyphen inside ("TD 0123", "ABC-123")
        for k in (1, 2, 3):
            if i + k > n:
                break
            c = "".join(t[i:i + k])
            if c in ids and (k == 1 or len(c) >= 5):
                out.append((ids[c], " ".join(t[i:i + k])))
        if i + 2 <= n and tuple(ta[i:i + 2]) in addrs:
            out.append((addrs[tuple(ta[i:i + 2])], " ".join(t[i:i + 3]) if i + 3 <= n and ta[i + 2] in STREET_TYPES else " ".join(t[i:i + 2])))
    for m in re.finditer(r"(?<!\d)(?:\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}(?!\d)|(?<!\d)\d{3}[\s.-]\d{4}(?!\d)", text):
        dg = re.sub(r"\D", "", m.group(0))
        if len(dg) == 11 and dg.startswith("1"):
            dg = dg[1:]
        if dg in phones:
            out.append((phones[dg], m.group(0)))
    seen, uniq = set(), []
    for x in out:
        if x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


# ---------------------------------------------------------------- sources
def staged_lines():
    """(path, line number, text) for every line the commit adds."""
    diff = git("diff", "--cached", "-U0", "--no-color", "--no-ext-diff", "--diff-filter=ACMR")
    path, ln = None, 0
    for line in diff.splitlines():
        if line.startswith("+++ "):
            path = line[6:] if line.startswith("+++ b/") else None
        elif line.startswith("@@"):
            m = re.search(r"\+(\d+)", line)
            ln = int(m.group(1)) if m else 0
        elif line.startswith("+") and path:
            yield path, ln, line[1:]
            ln += 1


def all_lines():
    for path in git("ls-files").splitlines():
        try:
            with open(path, "rb") as f:
                raw = f.read()
        except OSError:
            continue
        if b"\0" in raw[:8000]:
            continue
        for ln, line in enumerate(raw.decode("utf-8", "replace").splitlines(), 1):
            yield path, ln, line


def message_lines(path):
    with open(path, encoding="utf-8", errors="replace") as f:
        for ln, line in enumerate(f, 1):
            if not line.startswith("#"):
                yield "commit message", ln, line.rstrip("\n")


def forbidden_files():
    staged = git("diff", "--cached", "--name-only", "--diff-filter=ACMR").splitlines()
    return [p for p in staged
            if any(fnmatch.fnmatch(p, pat) or fnmatch.fnmatch(os.path.basename(p), pat) for pat in NEVER_COMMIT)]


# ---------------------------------------------------------------- main
def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "staged"
    root = git("rev-parse", "--show-toplevel").strip() or "."
    os.chdir(root)

    if mode == "staged":
        bad = forbidden_files()
        if bad:
            print("\nCommit blocked: these files must never go to GitHub:\n")
            for p in bad:
                print("  " + p)
            print("\nUnstage them with:  git restore --staged <file>\n")
            return 1

    data = os.environ.get("PVH_DATA") or os.path.join(root, "PVH_data.json")
    if not os.path.exists(data):
        print("pii check: PVH_data.json not found, so names could not be checked. Commit allowed.")
        return 0
    terms = load_terms(data)

    if mode == "staged":
        lines = staged_lines()
    elif mode == "message":
        lines = message_lines(sys.argv[2])
    elif mode == "all":
        lines = all_lines()
    else:
        print("usage: pii_check.py staged | message <file> | all")
        return 2

    hits = [(p, ln, kind, txt) for p, ln, line in lines for kind, txt in find(line, terms)]
    if not hits:
        if mode == "all":
            print("No personal information from PVH_data.json found in tracked files.")
        return 0

    what = {"staged": "Commit blocked", "message": "Commit blocked (commit message)",
            "all": "Found in tracked files"}[mode]
    print("\n" + what + ": personal information from the PVH data.\n")
    for p, ln, kind, txt in hits:
        print("  %s:%d  %s: %s" % (p, ln, kind, txt))
    if mode != "all":
        print("\nReplace it with a made-up example and commit again.")
        print("If this is a false alarm, commit with:  git commit --no-verify\n")
    return 1


if __name__ == "__main__":
    sys.exit(main())
