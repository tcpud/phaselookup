"""Build public/phases.json from the Tahoe Cedars phase spreadsheet.

Usage:  python build.py TC_Phase_Look_Up.xlsx      (or a .csv export)

Expected columns (header names not case sensitive):
  APN, Property_Street_Number, Property_Street_Name, Relocation, Phase

Relocation: Yes / No / Maybe  (blank, <Null>, N/A -> "not determined")
Phase:      "Phase 1" .. "Phase 4"  (blank, <Null>, 0 -> "no phase assigned")

Rows without a street number or street name are skipped and listed, since
customers can't search for them.
"""
import csv
import json
import re
import sys

SUFFIXES = {
    "avenue": "ave", "av": "ave", "boulevard": "blvd", "bl": "blvd",
    "street": "st", "drive": "dr", "road": "rd", "lane": "ln",
    "court": "ct", "circle": "cir", "place": "pl", "way": "wy",
    "highway": "hwy", "west": "w", "east": "e", "north": "n", "south": "s",
}
NULLS = {"", "<null>", "null", "none", "0", "n/a", "na"}


def normalize(addr: str) -> str:
    """Must match normalize() in public/index.html."""
    words = re.sub(r"[^a-z0-9 ]", " ", addr.lower()).split()
    return " ".join(SUFFIXES.get(w, w) for w in words)


def clean(v) -> str:
    s = re.sub(r"\s+", " ", str(v if v is not None else "")).strip()
    return "" if s.lower() in NULLS else s


def nice_case(s: str) -> str:
    """OLD CEDARS PLACE -> Old Cedars Place, 10TH AVENUE -> 10th Avenue."""
    def word(w):
        if re.match(r"^\d", w):
            return w.lower()
        return w[:1].upper() + w[1:].lower()
    return " ".join(word(w) for w in s.split(" "))


def relocation(v: str) -> str:
    v = v.lower()
    if v in ("yes", "y", "required", "true"):
        return "y"
    if v in ("no", "n", "not required", "false"):
        return "n"
    return "u"  # Maybe, blank, anything else


def phase(v: str) -> str:
    m = re.search(r"\d+", v)
    return m.group(0) if m and m.group(0) != "0" else ""


def read_rows(path: str):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True, read_only=True).active
        it = ws.iter_rows(values_only=True)
        header = [clean(h).lower() for h in next(it)]
        for r in it:
            yield dict(zip(header, r))
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            reader.fieldnames = [h.strip().lower() for h in reader.fieldnames]
            yield from reader


def main(path: str) -> None:
    out, skipped, seen = [], [], {}
    for line, r in enumerate(read_rows(path), start=2):
        if not any(r.values()):
            continue
        num = clean(r.get("property_street_number"))
        street = clean(r.get("property_street_name"))
        apn = clean(r.get("apn"))
        if not num or not street:
            skipped.append(f"row {line}: APN {apn or '?'} '{num} {street}'".rstrip())
            continue
        addr = nice_case(f"{num} {street}")
        rec = {"a": addr, "k": normalize(addr),
               "p": phase(clean(r.get("phase"))),
               "r": relocation(clean(r.get("relocation")))}
        prev = seen.get(rec["k"])
        if prev:
            if (prev["p"], prev["r"]) != (rec["p"], rec["r"]):
                print(f"CONFLICT row {line}: {addr} listed twice with different phase/relocation")
            continue  # same address, same answer (e.g. two APNs) -> keep one
        seen[rec["k"]] = rec
        out.append(rec)

    out.sort(key=lambda r: r["k"])
    with open("public/phases.json", "w", encoding="utf-8") as f:
        json.dump(out, f, separators=(",", ":"))

    print(f"Wrote {len(out)} addresses to public/phases.json")
    print(f"  no phase assigned: {sum(1 for r in out if not r['p'])}")
    print(f"  relocation not determined: {sum(1 for r in out if r['r'] == 'u')}")
    if skipped:
        print(f"Skipped {len(skipped)} rows with no street number or name:")
        for s in skipped:
            print("  " + s)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "TC_Phase_Look_Up.xlsx")
