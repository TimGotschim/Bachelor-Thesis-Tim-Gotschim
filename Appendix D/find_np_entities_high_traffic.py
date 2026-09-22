#!/usr/bin/env python3
"""
find_np_entities_high_traffic.py  v4
======================================
Sammelt bekannte Wikipedia-Entitäten aus thematischen Seed-Kategorien
und extrahiert deren numerische Wikidata-Properties für NP-Fragen
(Höhe, Länge, Fläche, Masse, Kapazität …).

WARUM KEIN PAGEVIEW-CHECK?
Die Wikimedia-Pageview-API erlaubt nur ~1 Request/Minute unter Last.
Bei 1882 Kandidaten würde ein Per-Artikel-Check mehrere Tage dauern.
Stattdessen:
  • Seed-Kategorien garantieren thematische Bekanntheit
  • Wikidata-Sitelinks als Popularitäts-Filter (ohne Extra-API-Call):
    MIN_SITELINKS = 20  →  Artikel existiert in ≥ 20 Wikipedia-Sprachen
    → entspricht grob Artikeln mit > 100.000 Views/Monat

ABLAUF
------
  1. Kandidaten aus Seed-Kategorien sammeln
  2. Wikidata-QID ermitteln (Wikipedia-API)
  3. Wikidata-Entity abrufen: Labels + Claims + Sitelinks-Zahl
  4. Sitelinks-Filter anwenden (MIN_SITELINKS)
  5. Numerische Properties extrahieren
  6. CSV + JSON speichern

Benötigte Pakete:  pip install requests pandas
Resume:  Bei Unterbrechung einfach neu starten —
         np_entities_progress.json wird weiterverwendet.
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

# Mindest-Sitelinks (Anzahl Sprachversionen des Wikipedia-Artikels).
# 20 = Artikel existiert in ≥ 20 Sprachen → gut bekannte Entität.
# Erhöhen für strengere Filterung, senken für mehr Abdeckung.
MIN_SITELINKS = 20

# Wikidata-Properties mit messbaren Zahlen (für NP-Fragen geeignet)
NP_PROPERTIES: dict[str, str] = {
    "P2048": "height",
    "P2044": "elevation above sea level",
    "P2043": "length",
    "P2386": "diameter",
    "P2120": "radius",
    "P4511": "vertical depth",
    "P2410": "depth",
    "P2046": "area",
    "P2234": "volume",
    "P2067": "mass",
    "P1083": "maximum capacity",
    "P1082": "population",
    "P2052": "speed",
    "P2916": "maximum speed",
    "P4100": "nameplate capacity",
    "P2101": "melting point",
    "P2102": "boiling point",
    "P2216": "orbital speed",
    "P2583": "distance from Earth",
    "P2060": "orbital period",
    "P2076": "operating temperature",
    "P2224": "mass of atmosphere",
}

# Seed-Kategorien: bewusst viele und verschiedene Domains
SEED_CATEGORIES: list[str] = [
    # Berge & Vulkane
    "Eight-thousanders",
    "Seven Summits",
    "Stratovolcanoes",
    "Active volcanoes",
    "Highest mountains",
    "Ultra-prominent peaks",
    "Mountains of Africa",
    "Mountains of Asia",
    "Mountains of Europe",
    "Mountains of North America",
    "Mountains of South America",
    # Flüsse, Seen, Meere
    "Rivers of Africa",
    "Rivers of Asia",
    "Rivers of Europe",
    "Rivers of North America",
    "Rivers of South America",
    "Oceans",
    "Marginal seas",
    "Straits",
    # Gebäude & Strukturen
    "Suspension bridges",
    "Cable-stayed bridges",
    "Dams",
    "Tunnels",
    "Skyscrapers in the United States",
    "Skyscrapers in China",
    "Skyscrapers in the United Arab Emirates",
    # Stadien
    "Cricket grounds",
    "Olympic stadiums",
    "Rugby union stadiums",
    "Tennis venues",
    # Astronomie
    "Planets of the Solar System",
    "Moons of Jupiter",
    "Moons of Saturn",
    "Dwarf planets",
    "Near-Earth objects",
    "Asteroids",
    # Länder & Städte
    "Capital cities",
    "Megacities",
    # Chemie & Elemente
    "Chemical elements",
    "Noble gases",
    "Alkali metals",
    "Halogens",
    "Transition metals",
    # Transport & Fahrzeuge
    "Airbus aircraft",
    "Boeing aircraft",
    "Airliners",
    "High-speed trains",
    "Formula One cars",
    "Ocean liners",
    "Aircraft carriers",
    "Submarines",
    # Raumfahrt
    "Space launch vehicles",
    "Space stations",
]

SLEEP        = 0.25   # Pause zwischen Requests (schont die API)
MAX_RETRIES  = 5
TIMEOUT      = 25

OUTPUT_CSV    = "np_entities_high_traffic.csv"
OUTPUT_JSON   = "np_entities_high_traffic.json"
PROGRESS_JSON = "np_entities_progress.json"

WIKIDATA_API  = "https://www.wikidata.org/w/api.php"
WIKIPEDIA_API = "https://en.wikipedia.org/w/api.php"

SKIP_PREFIXES = (
    "List of", "Index of", "Outline of", "Wikipedia:", "Special:",
    "Help:", "Portal:", "Template:", "Category:", "File:", "Draft:",
    "Module:", "User:", "Talk:", "MediaWiki:", "Book:",
)

# ─────────────────────────────────────────────────────────────────────────────
# HTTP
# ─────────────────────────────────────────────────────────────────────────────

_S = requests.Session()
_S.headers.update({"User-Agent": "NP_EntityFinder/4.0 (hallucination-research)"})


def _get(url: str, params: dict | None = None,
         allow_404: bool = False) -> dict | None:
    for attempt in range(MAX_RETRIES):
        try:
            r = _S.get(url, params=params, timeout=TIMEOUT)
        except requests.ConnectionError:
            time.sleep(min(2 ** attempt, 30))
            continue

        if r.status_code == 429:
            wait = min(float(r.headers.get("Retry-After", 2 ** attempt)), 30)
            print(f"  [429] warte {wait:.0f}s …")
            time.sleep(wait)
            continue

        if r.status_code in (500, 502, 503, 504):
            time.sleep(min(2 ** attempt, 20))
            continue

        if r.status_code == 404 and allow_404:
            return None

        r.raise_for_status()
        return r.json()

    return None


# ─────────────────────────────────────────────────────────────────────────────
# STUFE 1: KANDIDATEN AUS SEED-KATEGORIEN
# ─────────────────────────────────────────────────────────────────────────────

def get_category_members(category: str) -> list[str]:
    titles, cont = [], None
    while True:
        params: dict = {
            "action": "query", "format": "json",
            "list": "categorymembers",
            "cmtitle": f"Category:{category}",
            "cmnamespace": 0, "cmlimit": 500, "cmtype": "page",
        }
        if cont:
            params["cmcontinue"] = cont
        data = _get(WIKIPEDIA_API, params)
        if not data:
            break
        for item in data.get("query", {}).get("categorymembers", []):
            titles.append(item["title"])
        cont = data.get("continue", {}).get("cmcontinue")
        if not cont:
            break
        time.sleep(SLEEP)
    return titles


def build_candidate_pool() -> list[str]:
    pool: set[str] = set()
    for cat in SEED_CATEGORIES:
        members = get_category_members(cat)
        before = len(pool)
        for t in members:
            if not any(t.startswith(p) for p in SKIP_PREFIXES):
                pool.add(t)
        added = len(pool) - before
        print(f"  '{cat}': +{added} neu  (Pool: {len(pool)})")
        time.sleep(SLEEP)
    return list(pool)


# ─────────────────────────────────────────────────────────────────────────────
# STUFE 2: WIKIDATA-QID ERMITTELN (batched, 50 Artikel pro Request)
# ─────────────────────────────────────────────────────────────────────────────

def get_qids_batch(titles: list[str]) -> dict[str, str]:
    """Gibt {title: qid} für bis zu 50 Titel zurück."""
    data = _get(WIKIPEDIA_API, {
        "action": "query", "format": "json",
        "titles": "|".join(titles),
        "prop": "pageprops", "ppprop": "wikibase_item",
        "redirects": 1,
    })
    if not data:
        return {}

    # Redirect-Mapping aufbauen
    redirects: dict[str, str] = {}
    for r in data.get("query", {}).get("redirects", []):
        redirects[r["from"]] = r["to"]

    result: dict[str, str] = {}
    for page in data.get("query", {}).get("pages", {}).values():
        title = page.get("title", "")
        qid   = page.get("pageprops", {}).get("wikibase_item", "")
        if qid:
            result[title] = qid
            # Auch ursprünglichen Titel mappen falls Redirect
            for orig, dest in redirects.items():
                if dest == title:
                    result[orig] = qid
    return result


def get_all_qids(candidates: list[str]) -> dict[str, str]:
    """Holt QIDs für alle Kandidaten in 50er-Batches."""
    title_to_qid: dict[str, str] = {}
    batch_size = 50
    total = len(candidates)

    for i in range(0, total, batch_size):
        batch = candidates[i:i + batch_size]
        mapping = get_qids_batch(batch)
        title_to_qid.update(mapping)
        done = min(i + batch_size, total)
        found = len(title_to_qid)
        print(f"  QID-Lookup: {done}/{total} Artikel verarbeitet, "
              f"{found} QIDs gefunden", end="\r")
        time.sleep(SLEEP)

    print(f"\n  → {len(title_to_qid)} QIDs gefunden für {total} Kandidaten")
    return title_to_qid


# ─────────────────────────────────────────────────────────────────────────────
# STUFE 3: WIKIDATA-ENTITY (Labels + Sitelinks + Claims)
# ─────────────────────────────────────────────────────────────────────────────

def get_entity(qid: str) -> dict | None:
    data = _get(WIKIDATA_API, {
        "action": "wbgetentities", "format": "json", "ids": qid,
        "props": "labels|descriptions|claims|sitelinks",
        "languages": "en",
    })
    if not data:
        return None
    return data.get("entities", {}).get(qid)


_label_cache: dict[str, str] = {}


def resolve_labels(qids: list[str]) -> None:
    missing = [q for q in qids if q and q.startswith("Q") and q not in _label_cache]
    if not missing:
        return
    for i in range(0, len(missing), 50):
        batch = missing[i:i + 50]
        data = _get(WIKIDATA_API, {
            "action": "wbgetentities", "format": "json",
            "ids": "|".join(batch), "props": "labels", "languages": "en",
        })
        if not data:
            continue
        for qid, ent in data.get("entities", {}).items():
            lbl = ent.get("labels", {}).get("en", {}).get("value")
            _label_cache[qid] = lbl if lbl else qid
        time.sleep(SLEEP)


def extract_triples(entity: dict) -> tuple[str, int, list[dict]]:
    """
    Gibt (instance_of_label, sitelinks_count, [numerische Tripel]) zurück.
    """
    claims    = entity.get("claims", {})
    sitelinks = len(entity.get("sitelinks", {}))

    # instance_of (P31)
    p31_label = ""
    if "P31" in claims:
        qids = []
        for stmt in claims["P31"][:3]:
            snak = stmt.get("mainsnak", {})
            if snak.get("snaktype") == "value":
                qid = snak.get("datavalue", {}).get("value", {}).get("id", "")
                if qid:
                    qids.append(qid)
        if qids:
            resolve_labels(qids)
            p31_label = " / ".join(_label_cache.get(q, q) for q in qids)

    # Numerische Properties
    triples: list[dict] = []
    for prop_id, prop_label in NP_PROPERTIES.items():
        if prop_id not in claims:
            continue
        for stmt in claims[prop_id]:
            if stmt.get("rank") == "deprecated":
                continue
            snak = stmt.get("mainsnak", {})
            if snak.get("snaktype") != "value" or snak.get("datatype") != "quantity":
                continue
            dv       = snak.get("datavalue", {}).get("value", {})
            raw      = str(dv.get("amount", "")).lstrip("+")
            unit_url = dv.get("unit", "")
            unit_qid = unit_url.split("/")[-1] if unit_url and unit_url != "1" else ""
            if unit_qid:
                resolve_labels([unit_qid])
            unit_lbl = _label_cache.get(unit_qid, unit_qid) if unit_qid else "1"
            try:
                amount = float(raw)
            except ValueError:
                continue
            if amount <= 0:
                continue
            triples.append({
                "property_id":    prop_id,
                "property_label": prop_label,
                "amount":         amount,
                "amount_str":     raw,
                "unit_qid":       unit_qid,
                "unit_label":     unit_lbl,
            })

    return p31_label, sitelinks, triples


# ─────────────────────────────────────────────────────────────────────────────
# PERSISTENZ & AUSGABEN
# ─────────────────────────────────────────────────────────────────────────────

def load_progress() -> dict:
    p = Path(PROGRESS_JSON)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def save_progress(data: dict) -> None:
    Path(PROGRESS_JSON).write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def write_outputs(results: dict) -> None:
    Path(OUTPUT_JSON).write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    rows = []
    for article, info in results.items():
        base = {
            "article":        article,
            "wikidata_label": info.get("label", article),
            "qid":            info.get("qid", ""),
            "sitelinks":      info.get("sitelinks", ""),
            "instance_of":    info.get("instance_of", ""),
            "wikipedia_url":  f"https://en.wikipedia.org/wiki/{article.replace(' ', '_')}",
            "wikidata_url":   (f"https://www.wikidata.org/wiki/{info['qid']}"
                               if info.get("qid") else ""),
        }
        triples = info.get("numeric_triples") or []
        if not triples:
            rows.append({**base, "property_id": "", "property_label": "",
                         "amount": "", "unit_label": "", "unit_qid": ""})
        else:
            for t in triples:
                rows.append({**base,
                    "property_id":    t["property_id"],
                    "property_label": t["property_label"],
                    "amount":         t["amount"],
                    "unit_label":     t["unit_label"],
                    "unit_qid":       t["unit_qid"],
                })
    pd.DataFrame(rows).to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    n_art  = len(results)
    n_prop = sum(1 for v in results.values() if v.get("numeric_triples"))
    n_trip = sum(len(v.get("numeric_triples", [])) for v in results.values())
    print(f"  Gespeichert: {n_art} Artikel | "
          f"{n_prop} mit num. Properties | {n_trip} Tripel")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 60)
    print(f"NP Entity Finder  v4  (Filter: ≥ {MIN_SITELINKS} Sitelinks)")
    print("=" * 60)

    # ── 1. Kandidaten-Pool ────────────────────────────────────────────────
    print("\n[1/3] Kandidaten aus Seed-Kategorien …")
    candidates = build_candidate_pool()
    print(f"→ {len(candidates)} Kandidaten\n")

    # ── 2. QIDs in Batches holen (viel schneller als einzeln) ─────────────
    print("[2/3] Wikidata-QIDs ermitteln (50 Artikel pro Request) …")
    title_to_qid = get_all_qids(candidates)
    print()

    # ── 3. Entity-Daten + Sitelinks-Filter + Tripel-Extraktion ───────────
    print(f"[3/3] Entity-Daten abrufen & Sitelinks-Filter "
          f"(≥ {MIN_SITELINKS}) anwenden …\n")

    progress = load_progress()
    results  = {k: v for k, v in progress.items() if v.get("status") == "qualified"}
    checked  = set(progress.keys())

    for idx, (title, qid) in enumerate(title_to_qid.items(), 1):
        if title in checked:
            continue

        entity = get_entity(qid)
        time.sleep(SLEEP)

        if not entity:
            progress[title] = {"status": "no_entity", "qid": qid}
            if idx % 50 == 0:
                save_progress(progress)
            continue

        label     = entity.get("labels", {}).get("en", {}).get("value", title)
        desc      = entity.get("descriptions", {}).get("en", {}).get("value", "")
        instance_of, sitelinks, triples = extract_triples(entity)

        # Sitelinks-Filter
        if sitelinks < MIN_SITELINKS:
            progress[title] = {"status": "few_sitelinks",
                                "qid": qid, "sitelinks": sitelinks}
            if idx % 50 == 0:
                save_progress(progress)
            continue

        entry = {
            "status":          "qualified",
            "qid":             qid,
            "label":           label,
            "description":     desc,
            "sitelinks":       sitelinks,
            "instance_of":     instance_of,
            "numeric_triples": triples,
        }
        results[title]  = entry
        progress[title] = entry

        # Ausgabe nur für Artikel mit numerischen Properties
        if triples:
            prop_str = ", ".join(
                f"{t['property_label']}={t['amount_str']} {t['unit_label']}"
                for t in triples[:3]
            )
            print(f"  [{len(results):>4}] {title[:42]:<42} "
                  f"SL={sitelinks:>4}  |  {prop_str}")

        save_progress(progress)
        if len(results) % 30 == 0:
            write_outputs(results)

        time.sleep(SLEEP)

    # ── Finale Ausgabe ────────────────────────────────────────────────────
    print("\n→ Finale Ausgabe …")
    write_outputs(results)

    prop_counts: dict[str, int] = {}
    for v in results.values():
        for t in v.get("numeric_triples", []):
            lbl = t["property_label"]
            prop_counts[lbl] = prop_counts.get(lbl, 0) + 1

    print("\n" + "=" * 60)
    print(f"FERTIG")
    print(f"  Qualifizierte Artikel (≥ {MIN_SITELINKS} Sitelinks): "
          f"{len(results)}")
    print(f"  Davon mit num. Properties:  "
          f"{sum(1 for v in results.values() if v.get('numeric_triples'))}")
    print(f"  Numerische Tripel gesamt:   "
          f"{sum(len(v.get('numeric_triples', [])) for v in results.values())}")
    print()
    print("  Häufigste Properties:")
    for lbl, cnt in sorted(prop_counts.items(), key=lambda x: -x[1])[:12]:
        print(f"    {lbl:<35} {cnt:>4}×")
    print(f"\n  CSV:  {OUTPUT_CSV}")
    print(f"  JSON: {OUTPUT_JSON}")
    print("=" * 60)


if __name__ == "__main__":
    main()
