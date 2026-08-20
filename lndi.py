#!/usr/bin/env python3
"""
LNDI pricing tool.

Reads molecule names from a chosen sheet of any workbook, looks each one up in
Lebanon's National Drug Index by active ingredient, pulls price / stratum / B-G
off each product page, and computes:

    After Division = Price / FOB to PP LL[stratum]
    After -30%     = After Division * 0.70   (brand only; generics unchanged)
    After -40%     = After Division * 0.60   (brand only; generics unchanged)
    Final -30%     = After -30% / 3
    Final -40%     = After -40% / 3

Both reduction paths are carried through, so no reference-country choice is
needed: read whichever column applies. Each step is its own column in the
output, written as a live Excel formula.

Usage:  python3 lndi.py [--refresh]
"""

import re, sys, html, difflib, hashlib, argparse, threading, subprocess
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
from pathlib import Path
from urllib.parse import quote

import pandas as pd

# Inside a PyInstaller bundle __file__ points at a read-only temp folder, so
# anything we write has to live in the user's home instead. BUNDLE is where
# read-only files we shipped (the stratum table) can be found.
FROZEN   = getattr(sys, "frozen", False)
BUNDLE   = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
ROOT     = (Path.home() / ".drugprices") if FROZEN else Path(__file__).resolve().parent
DATA     = ROOT / "Data"
CACHE    = ROOT / "cache"
OUT      = ROOT / "output.xlsx"
INBOX    = ROOT / "input"          # drop any workbook here
SHEETEXT = (".xlsx", ".xlsm", ".xls", ".xltx", ".csv")

BASE     = "https://www.moph.gov.lb"
LIST_URL = f"{BASE}/en/Drugs/index/3/4848"
VIEW_URL = f"{BASE}/en/Drugs/view/{{}}"
AUTO_URL = f"{BASE}/en/Drugs/autocomplete_ingredientSearch?term={{}}"
UA       = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"

REFRESH  = False                   # set from argv in main()


class InputError(Exception):
    """A problem with the user's files, reportable rather than fatal."""

# A cell that could plausibly be a molecule label.
NAME_RE  = re.compile(r"^[A-Za-z][A-Za-z0-9\s\-\+/(),.'%]{2,}$")
MIN_STEM = 5                       # never truncate a name shorter than this
PROBE_JOBS = 8                     # parallel autocomplete lookups
MIN_SIM  = 0.75                    # min similarity to accept a fuzzy match
                                   # (0.62 let section headers like
                                   #  "LYOPHILIZED INJECTIONS" match "Lyophilized")

# Stratum is assigned by FOB USD bracket (per the MoPH pricing guideline).
# Used as a self-check: a correct FOB_USD must fall inside its own stratum's band.
BRACKETS = {"A1": (0, 5), "A2": (5, 10), "B": (10, 50), "C": (50, 100),
            "D": (100, 300), "E1": (300, 700), "E2": (700, float("inf"))}

# Columns written to the Results sheet, in order: (internal field, Excel heading).
# Everything else stays internal (used for the bracket self-check and reporting).
OUTPUT_COLS = [
    ("B/G", "B/G"), ("Ingredient", "Ingredient"), ("Brand", "Name"),
    ("Dosage", "Dosage"), ("Presentation", "Presentation"), ("Form", "Form"),
    ("Route", "Route"), ("Country", "Country"), ("Price_LL", "Price"),
    ("Stratum", "Stratum"), ("FOB_to_PP_LL", "FOB to PP LL"),
    ("FOB_USD", "After Division"), ("After_30", "After -30%"),
    ("After_40", "After -40%"), ("Final_30", "Final -30%"),
    ("Final_40", "Final -40%"),
]

# Site label -> our field name. Everything else passes through unchanged.
LABELS = {"Registration Nb": "Registration_Nb", "Pharmacist Margin": "Ph_Margin",
          "Responsible Party Name": "Responsible_Party_Name",
          "Responsible Party Country": "Responsible_Party_Country"}

# Strength units, normalised to milligrams so '1g' == '1000mg'.
UNIT_MG  = {"mcg": 0.001, "ug": 0.001, "mg": 1, "gm": 1000, "gr": 1000, "g": 1000}
SKIP_UNIT = {"ml", "l", "vial", "vials"}   # volumes, not strengths

# Words that are dosage-form noise, not part of a molecule name.
FORM_WORDS = {"TAB", "TABS", "TABLET", "TABLETS", "CAP", "CAPS", "CAPSULE", "CAPSULES",
              "VIAL", "VIALS", "INJ", "INJECTION", "AMP", "AMPOULE", "SYRUP", "SUSP",
              "SOL", "SOLUTION", "POWDER", "IV", "ORAL", "MG", "ML", "G"}


# ---------------------------------------------------------------- http + cache

def fetch(url, post=None, tag=None):
    """GET/POST through curl (the system Python has no usable CA bundle), with an
    on-disk cache so re-parsing never re-hits the ministry's server."""
    CACHE.mkdir(parents=True, exist_ok=True)   # ~/.drugprices itself may not exist yet
    key = tag or hashlib.sha1(f"{url}|{post}".encode()).hexdigest()[:16]
    path = CACHE / f"{key}.html"
    if path.exists() and not REFRESH:
        return path.read_text(encoding="utf-8", errors="replace")

    cmd = ["curl", "-sS", "--compressed", "-A", UA, "--max-time", "60",
           "--retry", "3", "--retry-delay", "2", url]
    for k, v in (post or []):
        cmd += ["--data-urlencode", f"{k}={v}"]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"fetch failed for {url}: {r.stderr.strip()}")
    tmp = path.with_name(f"{path.name}.{threading.get_ident()}.tmp")
    tmp.write_text(r.stdout, encoding="utf-8")
    tmp.replace(path)                     # atomic: threads may share a key
    return r.stdout


# ---------------------------------------------------------------- html helpers

def slug(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", str(s))[:40]


def strip_tags(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", s))).strip()

def flatten(page):
    """Detail pages are label-block-then-value-block; flattening to a token list
    lets us index fields positionally without depending on the table markup."""
    body = re.sub(r"<script.*?</script>", "", page, flags=re.S)
    toks = (re.sub(r"\s+", " ", html.unescape(t)).strip()
            for t in re.sub(r"<[^>]+>", "|", body).split("|"))
    return [t for t in toks if t]

def dose_keys(text):
    """Every strength in a label, normalised to (milligrams, unit-class).

    '1g/Vial' and '1,000mg' both become (1000.0, 'mg'); '100mg/5ml' -> (100.0,
    'mg') because the 5ml is volume, not strength; '5%' keeps its own class so it
    can never match a milligram dose. A bare number is assumed to be mg."""
    if text is None:
        return []
    body = str(text).replace(",", "").lower()
    found, trailing = [], None
    for m in re.finditer(r"(\d+(?:\.\d+)?)", body):
        u = re.match(r"\s*(mcg|ug|mg|gm|gr|g|ml|l|vials?|%)", body[m.end():])
        unit = u.group(1) if u else None
        if unit in SKIP_UNIT:
            continue
        found.append((float(m.group(1)), unit))
        if unit:
            trailing = unit
    keys = []
    for val, unit in found:
        unit = unit or trailing or "mg"          # '100/200/500MG': MG covers all
        keys.append((val, "%") if unit == "%"
                    else (round(val * UNIT_MG.get(unit, 1), 6), "mg"))
    return keys


def dose_key(text):
    """The single strength of a product listing, or None if it has none."""
    keys = dose_keys(text)
    return keys[0] if keys else None


def fmt_dose(key):
    val, unit = key
    if unit == "%":
        return f"{val:g}%"
    return f"{val / 1000:g}g" if val >= 1000 else f"{val:g}mg"


def price_value(text):
    digits = re.sub(r"[^\d]", "", text or "")
    return int(digits) if digits else None


# ---------------------------------------------------------------- input sheets

def parse_label(label):
    """'NILOTINIB 150' -> ('NILOTINIB', [150.0]).
    'DACARBAZINE 100/200/500MG' -> ('DACARBAZINE', [3 mg-normalised keys])."""
    label = str(label).strip()
    m = re.search(r"\d", label)
    name, rest = (label[:m.start()], label[m.start():]) if m else (label, "")
    words = [w for w in re.split(r"[\s\-]+", name.upper().strip())
             if w and w not in FORM_WORDS]
    return " ".join(words), dose_keys(rest)


@lru_cache(maxsize=4096)
def ingredient_hits(term):
    """Ingredient names the site's autocomplete returns for a search term."""
    if not term or len(term) < 3:
        return []
    body = fetch(AUTO_URL.format(quote(term)), tag=f"auto_{slug(term)}")
    return tuple(re.findall(r'"([^"]+)"\s*:', body))


@lru_cache(maxsize=4096)
def resolve_ingredient(name):
    """Map our spelling onto the site's. The autocomplete is a prefix search, so
    truncating the stem recovers variants (ANASTRAZOLE -> ANASTR -> Anastrozole).
    Returns None when the molecule genuinely is not in the index."""
    hits = ingredient_hits(name)
    if hits:
        return hits[0]
    stem = name.split()[0] if name else ""
    for n in range(len(stem) - 1, MIN_STEM - 1, -1):
        hits = ingredient_hits(stem[:n])
        if not hits:
            continue
        best = max(hits, key=lambda h: difflib.SequenceMatcher(
            None, h.upper(), name.upper()).ratio())
        if difflib.SequenceMatcher(None, best.upper(), name.upper()).ratio() >= MIN_SIM:
            return best
        return None
    return None


def find_input(explicit=None):
    """Whichever workbook the user dropped in: an explicit path, else the newest
    file in input/, else the newest non-stratum workbook in Data/."""
    if explicit:
        f = Path(explicit)
        if not f.exists():
            raise InputError(f"no such file: {f}")
        return f
    pools = [sorted((q for q in INBOX.glob("*") if q.suffix.lower() in SHEETEXT
                     and not q.name.startswith("~$")),
                    key=lambda q: q.stat().st_mtime, reverse=True),
             sorted((q for q in DATA.glob("*") if q.suffix.lower() in SHEETEXT
                     and not q.name.startswith("~$")
                     and not re.search(r"stratum|rates", q.name, re.I)),
                    key=lambda q: q.stat().st_mtime, reverse=True)]
    for pool in pools:
        if pool:
            return pool[0]
    raise InputError(f"No input workbook found. Drop one into {INBOX}/ and re-run.")


def sheets_of(path):
    if path.suffix.lower() == ".csv":
        return {"csv": pd.read_csv(path, header=None, dtype=object)}
    xl = pd.ExcelFile(path)
    return {n: xl.parse(n, header=None, dtype=object) for n in xl.sheet_names}


def list_sheets(path):
    """Sheet names of a real spreadsheet, or [] when there is nothing to choose."""
    try:
        if Path(path).suffix.lower() == ".csv":
            return []
        return pd.ExcelFile(path).sheet_names
    except Exception:
        return []


def best_sheet(path):
    """The sheet that most looks like a molecule list — used as the default."""
    try:
        options = sheet_options(Path(path))
    except Exception:
        return None
    return max(options, key=lambda o: o[1])[0] if options else None


def sheet_options(path):
    """(sheet, molecule-ish row count, sample values) for every sheet, so the user
    has something to choose between. Structure only — no network lookups."""
    out = []
    for name, df in sheets_of(path).items():
        best = (0, 0, [])
        for col in df.columns:
            cells = [str(v).strip() for v in df[col]
                     if pd.notna(v) and str(v).strip()]
            ok = [c for c in cells
                  if (lambda n: n and len(n) >= 3 and NAME_RE.match(n))(parse_label(c)[0])]
            if not ok:
                continue
            # Repeated values ("Potential") are labels, not molecule names.
            score = len(ok) * (len(set(ok)) / len(ok))
            if score > best[0]:
                # Prefer entries carrying a strength for the preview: those are
                # certainly molecules, whereas "OSD (TAB / CAP) - ONCOLOGY" is a
                # section heading that would otherwise lead every sample.
                dosed = [c for c in ok if parse_label(c)[1]]
                best = (score, len(ok), (dosed or ok)[:3])
        out.append((name, best[1], ", ".join(best[2])))
    return out


def detect_column(frames, want_sheet=None, want_col=None):
    """Find the sheet+column holding molecule names, with no assumption about
    layout. Structural scoring narrows the field; the site's own ingredient
    index then confirms which column actually contains drug molecules."""
    cands = []
    for sheet, df in frames.items():
        if want_sheet and sheet != want_sheet:
            continue
        for col in df.columns:
            if want_col is not None and col != want_col:
                continue
            cells = [(i, str(v).strip()) for i, v in df[col].items()
                     if pd.notna(v) and str(v).strip()]
            named = [(i, parse_label(v)[0]) for i, v in cells]
            ok = [(i, n) for i, n in named if n and len(n) >= 3 and NAME_RE.match(n)]
            if not ok:
                continue
            # Repeated values ("Potential", "Found in") are labels, not molecules.
            distinct = len({n for _, n in ok}) / len(ok)
            cands.append({"sheet": sheet, "col": col, "cells": cells,
                          "ok": ok, "struct": len(ok) * distinct})

    if not cands:
        raise InputError("Could not find any column that looks like molecule names.")
    if want_col is not None:
        return cands[0]

    cands.sort(key=lambda c: -c["struct"])
    top = cands[:4]
    for c in top:
        c["probe"] = list(dict.fromkeys(n for _, n in c["ok"]))[:5]
    with ThreadPoolExecutor(max_workers=PROBE_JOBS) as pool:
        futures = {n: pool.submit(ingredient_hits, n)
                   for c in top for n in c["probe"]}
    for c in top:
        hits = sum(1 for n in c["probe"] if futures[n].result())
        c["rate"] = hits / len(c["probe"]) if c["probe"] else 0
    best = max(top, key=lambda c: (c["rate"], c["struct"]))
    return best if best.get("rate") else cands[0]


def read_molecules(path, want_sheet=None, want_col=None):
    """(raw_label, molecule_name, [doses]) for every molecule row in the file."""
    frames = sheets_of(path)
    pick = detect_column(frames, want_sheet, want_col)
    # The first row that the drug index recognises marks where the data begins.
    # Stop there rather than resolving every cell, then step back over any rows
    # above it that only a fuzzy match reaches (a misspelling in row one).
    names = dict(pick["ok"])
    first = next((i for i, n in pick["ok"] if ingredient_hits(n)), None)
    if first is None:
        start = min(i for i, _ in pick["ok"])
    else:
        start = first
        for i in sorted((i for i in names if i < first), reverse=True):
            if not resolve_ingredient(names[i]):
                break
            start = i

    rows, skipped = [], []
    for i, raw in pick["cells"]:
        name, doses = parse_label(raw)
        if not name:
            continue
        (rows if i >= start else skipped).append((raw, name, doses))

    print(f"  input:     {path.name}  [sheet '{pick['sheet']}', column {pick['col']}]")
    if skipped:
        print(f"  skipped above data: {', '.join(repr(r) for r, _, _ in skipped)}")
    return rows


def read_stratum(src=None):
    """stratum -> FOB to PP LL, plus a consistency warning on the implied FX rate."""
    if src:
        return _stratum_from(Path(src))
    # A table the user dropped in beats the copy shipped inside the app, so the
    # rates can be updated without rebuilding.
    def rates_in(folder):
        return [p for p in folder.glob("*")
                if re.search(r"stratum|rates", p.name, re.I)
                and p.suffix.lower() in (".xlsx", ".xltx", ".xlsm", ".csv")]

    files = (rates_in(DATA) or rates_in(BUNDLE / "Data")
             or sorted(Path.home().glob("Downloads/rates_table*.xlsx")))
    if not files:
        raise InputError("No stratum table found. Use the Change… button to pick "
                         "the FOB-to-PP-LL sheet.")
    return _stratum_from(files[0])


def _stratum_from(src):
    df = (pd.read_csv(src, header=None) if src.suffix.lower() == ".csv"
          else pd.read_excel(src, header=None))

    fob, usd = {}, {}
    for _, r in df.iloc[1:].iterrows():
        try:
            fob[str(r[0]).strip()] = float(r[3])
            usd[str(r[0]).strip()] = float(r[1])
        except (ValueError, TypeError, KeyError):
            continue

    rates = {s: fob[s] / usd[s] for s in fob if usd.get(s)}
    if rates:
        modal = max(set(round(v, -2) for v in rates.values()),
                    key=lambda x: sum(abs(v - x) < 50 for v in rates.values()))
        for s, v in rates.items():
            if abs(v - modal) > 50:
                print(f"  ! stratum {s}: FOB_LL implies FX {v:,.1f} but every other "
                      f"row implies {modal:,.0f} -> expected {usd[s] * modal:,.2f}, "
                      f"sheet says {fob[s]:,.2f}")
    return fob, src


# ---------------------------------------------------------------- site queries

def search(ingredient):
    """All products for an active ingredient: [(view_id, {list-row fields})]."""
    page = fetch(LIST_URL, post=[("_method", "POST"),
                                 ("data[Drug][ingredient]", ingredient)],
                 tag=f"search_{ingredient.replace(' ', '_')}")
    tables = re.findall(r"<table.*?</table>", page, re.S)
    if not tables:
        return []
    rows = re.findall(r"<tr.*?</tr>", tables[0], re.S)[1:]
    found = []
    for r in rows:
        cells = [strip_tags(c) for c in re.findall(r"<t[hd].*?</t[hd]>", r, re.S)]
        vid = re.search(r"/Drugs/view/(\d+)", r)
        if len(cells) < 7 or not vid:
            continue
        found.append((vid.group(1), {"Brand": cells[1], "B/G": cells[2],
                                     "Dosage": cells[4], "Form": cells[5]}))
    return found


def detail(view_id):
    """Fields of one drug, read by pairing the detail table's header cells with
    its value cells. Empty cells keep their slot, so a blank Dosage can no longer
    shift every field after it."""
    page = re.sub(r"<script.*?</script>", "",
                  fetch(VIEW_URL.format(view_id), tag=f"view_{view_id}"), flags=re.S)
    cells = lambda row: [strip_tags(c) for c in
                         re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", row, re.S)]
    for tbl in re.findall(r"<table.*?</table>", page, re.S):
        if "Stratum" not in tbl:
            continue
        rows = re.findall(r"<tr.*?</tr>", tbl, re.S)
        if len(rows) < 2:
            continue
        labels, values = cells(rows[0]), cells(rows[1])
        if len(labels) != len(values):
            continue
        rec = {LABELS.get(k.strip(), k.strip()): v
               for k, v in zip(labels, values) if k.strip()}
        if "Price" in rec and "Stratum" in rec:
            rec["id"] = view_id
            return rec
    raise RuntimeError(f"no drug detail table found for drug {view_id}")


# ---------------------------------------------------------------- main

def run(path, want_sheet=None, want_col=None, stratum=None, on_progress=None):
    """Scrape and price every molecule in `path`. Returns (results, exceptions)."""
    print("Reading inputs")
    molecules = read_molecules(path, want_sheet, want_col)
    fob_table, stratum_src = read_stratum(stratum)
    print(f"  molecules: {len(molecules)}")
    print(f"  stratum:   {stratum_src.name} ({len(fob_table)} strata)")

    print("\nSearching")
    results, exceptions, seen = [], [], set()
    for n, (label, name, doses) in enumerate(molecules, 1):
        if on_progress:
            on_progress(n, len(molecules), label)
        ing = resolve_ingredient(name)
        if not ing:
            print(f"  {label:<28} -> not in the LNDI ingredient index")
            exceptions.append({"Molecule": label, "Reason": "molecule not in LNDI index",
                               "Detail": f"no ingredient matching '{name}'"})
            continue
        if ing.upper() != name.upper():
            print(f"  {label:<28} -> matched site spelling '{ing}'")
        products = search(ing)
        wanted = [(vid, row) for vid, row in products
                  if not doses or dose_key(row["Dosage"]) in doses]
        print(f"  {label:<28} -> {ing:<16} {len(wanted)}/{len(products)} product(s)")

        if not products:
            exceptions.append({"Molecule": label, "Reason": "no products for ingredient",
                               "Detail": ing})
            continue
        if not wanted:
            avail = sorted({r["Dosage"] for _, r in products if r["Dosage"]})
            exceptions.append({"Molecule": label, "Reason": "requested strength not registered",
                               "Detail": f"wanted {', '.join(fmt_dose(d) for d in doses)}; "
                                         f"site has {', '.join(avail) or 'no strengths listed'}"})
            continue
        for dose in doses:
            if not any(dose_key(r["Dosage"]) == dose for _, r in wanted):
                exceptions.append({"Molecule": label, "Reason": "strength not registered",
                                   "Detail": f"{fmt_dose(dose)} not found"})

        for vid, _ in wanted:
            if vid in seen:
                continue
            seen.add(vid)
            d = detail(vid)
            price = price_value(d["Price"])
            stratum = (d["Stratum"] or "").strip()
            fob = fob_table.get(stratum)

            if price is None:
                exceptions.append({"Molecule": label, "Reason": "no price on site",
                                   "Detail": f"{d['Name']} {d['Dosage']} (id {vid}, "
                                             f"reg {d['Registration_Nb']})"})
                continue
            if fob is None:
                exceptions.append({"Molecule": label, "Reason": "stratum not in FOB table",
                                   "Detail": f"{d['Name']} stratum '{stratum}' (id {vid})"})
                continue

            usd = price / fob
            lo, hi = BRACKETS.get(stratum, (None, None))
            ok = "OK" if lo is None or lo <= usd <= hi else f"OUT ({lo}-{hi})"
            brand = d["B/G"].strip().upper() == "B"
            # One column per step, so the sheet shows its work: price / rate ->
            # less 30% and less 40% -> each divided by 3. Both paths are carried
            # through; generics take no reduction, so their columns are equal.
            after_30 = usd * 0.70 if brand else usd
            after_40 = usd * 0.60 if brand else usd

            results.append({
                "ID": vid, "Molecule": label, "Brand": d["Name"], "B/G": d["B/G"],
                "Ingredient": d.get("Ingredients", ""),
                "Dosage": d["Dosage"], "Presentation": d["Presentation"],
                "Form": d["Form"], "Route": d["Route"], "Country": d["Country"],
                "Laboratory": d["Laboratory"], "Agent": d["Agent"],
                "Registration_Nb": d["Registration_Nb"], "ATC": d["ATC"],
                "Price_LL": price, "Ph_Margin": d["Ph_Margin"], "Stratum": stratum,
                "FOB_to_PP_LL": fob, "FOB_USD": round(usd, 2), "Bracket_Check": ok,
                "After_30": round(after_30, 2), "After_40": round(after_40, 2),
                "Final_30": round(after_30 / 3, 2), "Final_40": round(after_40 / 3, 2),
            })

    return (pd.DataFrame(results),
            pd.DataFrame(exceptions, columns=["Molecule", "Reason", "Detail"]))


def write_output(res, exc, out=None):
    out = Path(out or OUT)
    sheet = (res.reindex(columns=[f for f, _ in OUTPUT_COLS])
                .rename(columns=dict(OUTPUT_COLS))
             if len(res) else pd.DataFrame(columns=[h for _, h in OUTPUT_COLS]))

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        sheet.to_excel(w, sheet_name="Results", index=False)
        exc.to_excel(w, sheet_name="Exceptions", index=False)

        ws = w.sheets["Results"]
        ws.freeze_panes = "D2"
        for i, col in enumerate(sheet.columns, 1):
            width = (max(len(col), *(len(str(v)) for v in sheet[col]))
                     if len(sheet) else len(col))
            ws.column_dimensions[ws.cell(1, i).column_letter].width = min(width + 2, 34)

        if len(sheet):
            cols = list(sheet.columns)
            at = lambda name: ws.cell(1, cols.index(name) + 1).column_letter
            bg, price, rate = at("B/G"), at("Price"), at("FOB to PP LL")
            div, m30, m40 = at("After Division"), at("After -30%"), at("After -40%")
            f30, f40 = at("Final -30%"), at("Final -40%")

            # Every step is a live formula, so the sheet shows its own working.
            # The reduction columns test B/G rather than hard-coding a factor, so
            # generics correctly take none and a corrected B/G recalculates the row.
            for r in range(2, len(sheet) + 2):
                for cell, formula in (
                    (div, f"={price}{r}/{rate}{r}"),
                    (m30, f'=IF(UPPER({bg}{r})="B",{div}{r}*0.7,{div}{r})'),
                    (m40, f'=IF(UPPER({bg}{r})="B",{div}{r}*0.6,{div}{r})'),
                    (f30, f"={m30}{r}/3"),
                    (f40, f"={m40}{r}/3"),
                ):
                    ws[f"{cell}{r}"] = formula
                    ws[f"{cell}{r}"].number_format = "0.00"
                ws[f"{price}{r}"].number_format = "#,##0"
                ws[f"{rate}{r}"].number_format = "#,##0.00"
        for i, col in enumerate(exc.columns, 1):
            w.sheets["Exceptions"].column_dimensions[
                w.sheets["Exceptions"].cell(1, i).column_letter].width = 46

    return out


def main():
    global REFRESH
    ap = argparse.ArgumentParser(description="Price LNDI molecules from any workbook.")
    ap.add_argument("file", nargs="?", help="input workbook (default: newest in input/)")
    ap.add_argument("--sheet", help="force a sheet name")
    ap.add_argument("--column", type=int, help="force a 0-based column index")
    ap.add_argument("--out", help=f"output workbook (default: {OUT.name})")
    ap.add_argument("--refresh", action="store_true", help="re-fetch instead of using cache")
    a = ap.parse_args()
    REFRESH = a.refresh

    try:
        res, exc = run(find_input(a.file), a.sheet, a.column)
    except InputError as e:
        sys.exit(f"error: {e}")
    out = write_output(res, exc, a.out)
    print(f"\n{len(res)} product(s) priced, {len(exc)} exception(s) -> {out.name}")
    bad = res[res["Bracket_Check"] != "OK"] if len(res) else res
    print("Bracket self-check: " + ("all pass" if not len(bad) else f"{len(bad)} FAILED"))
    return res, exc


if __name__ == "__main__":
    main()
