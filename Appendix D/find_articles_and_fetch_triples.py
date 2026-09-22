#!/usr/bin/env python3
"""
find_articles_and_fetch_triples.py
====================================
Findet 200+ englische Wikipedia-Artikel mit maximal 400 durchschnittlichen
monatlichen Aufrufen (2021–2023) und holt für jeden Artikel automatisch
die verfügbaren Wikidata-Tripel.

STRATEGIE
----------
Phase 1 – Artikel finden
  • Zufallsartikel (Wikipedia Random-API) als Basis
  • Sobald ein Treffer gefunden wird: Links, Backlinks und Kategoriemitglieder
    dieses Artikels werden in die Queue aufgenommen (Enrichment).
    Nischenartikel verlinken typischerweise auf ähnlich nischige Artikel.
  • Pageview-Prüfung: Jahresdurchschnitt über 2021–2023 (ein einziger
    API-Call deckt alle 36 Monate ab).
  • Akzeptiert: 50 ≤ Ø Views/Monat ≤ 400 (mind. 12 Datenpunkte)

Phase 2 – Wikidata-Tripel abrufen
  • QID via Wikipedia-API
  • Alle verfügbaren Claims via Wikidata-API
  • Property-Labels und Item-Labels auf Englisch auflösen (gebatcht)
  • Unbrauchbare Properties (Bilder, externe IDs etc.) herausfiltern

AUSGABEN
----------
  longtail_articles_with_triples.json   vollständige Daten je Artikel
  longtail_articles_with_triples.csv    flache Tabelle (1 Zeile pro Tripel)
  longtail_finder_progress.json         Resume-Datei

AUFRUF
------
  pip install requests pandas
  python3 find_articles_and_fetch_triples.py

  # Ziel anpassen (Default: 200):
  python3 find_articles_and_fetch_triples.py --target 250

  # Pageview-Obergrenze ändern:
  python3 find_articles_and_fetch_triples.py --max-views 300
"""

from __future__ import annotations

import argparse
import json
import random
import time
from collections import deque
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

TARGET_COUNT       = 200     # Mindestanzahl gesuchter Artikel
MIN_MONTHLY_VIEWS  = 50      # Untergrenze (Stubs ausschließen)
MAX_MONTHLY_VIEWS  = 400     # Obergrenze

PAGEVIEW_START     = "20210101"   # Prüfzeitraum: 2021–2023
PAGEVIEW_END       = "20231231"
REQUIRE_N_MONTHS   = 12           # Mindestanzahl Monatswerte

RANDOM_BATCH_SIZE  = 25           # Zufallsartikel pro Batch
MAX_CANDIDATES     = 50_000       # Sicherheitsgrenze

# Enrichment: wie viele Artikel pro Quelle holen
# (kleiner = weniger API-Calls = weniger 429-Fehler)
LINKS_PER_FOUND      = 20
BACKLINKS_PER_FOUND  = 15
CATMEMBERS_PER_FOUND = 25

# Enrichment nur jeden N-ten Treffer ausführen (nicht jeden)
# Spart API-Calls erheblich; die Queue bleibt trotzdem lang genug
ENRICH_EVERY_N      = 3

SLEEP_DEFAULT      = 0.8    # Pause zwischen Wikipedia-Requests
SLEEP_WIKIDATA     = 1.2    # Pause zwischen Wikidata-Requests
MAX_RETRIES        = 5
TIMEOUT            = 45     # Länger – schützt gegen langsame Verbindungen

OUTPUT_JSON        = "longtail_articles_with_triples.json"
OUTPUT_CSV         = "longtail_articles_with_triples.csv"
PROGRESS_JSON      = "longtail_finder_progress.json"

WIKIPEDIA_API      = "https://en.wikipedia.org/w/api.php"
WIKIDATA_API       = "https://www.wikidata.org/w/api.php"
PAGEVIEWS_API      = "https://wikimedia.org/api/rest_v1"

# Properties, die für Fragen ungeeignet sind (externe IDs, Bilder usw.)
SKIP_PROPERTIES: set[str] = {
    "P18", "P373", "P910", "P301", "P935", "P856",
    "P2013", "P2002", "P4223", "P2397", "P3267",
    "P214", "P213", "P244", "P268", "P269", "P227", "P349",
    "P906", "P950", "P648", "P1006", "P2163", "P3430",
    "P691", "P7293", "P4619", "P1890", "P2799", "P1368",
    "P2188", "P5034", "P8094", "P109", "P2860", "P6886",
    "P1038", "P1440", "P1871", "P7859", "P4265", "P7902",
    "P5587", "P3987", "P9984", "P4228", "P4293",
}

# Titelpatterns, die ausgeschlossen werden
SKIP_PREFIXES = (
    "List of", "Index of", "Outline of",
    "Wikipedia:", "Special:", "Help:", "Portal:", "Template:",
    "Category:", "File:", "Draft:", "Module:", "User:",
    "Talk:", "MediaWiki:", "Book:",
)
SKIP_CONTAINS = (
    "Deaths in ", "Births in ", "Events in ",
    "in film", "in music", "in television",
)
SKIP_EXACT = {"Main Page", "-"}

# ─────────────────────────────────────────────────────────────────────────────
# HTTP
# ─────────────────────────────────────────────────────────────────────────────

_S = requests.Session()
_S.headers.update({"User-Agent": "LTF_ArticleFinder/1.0 (hallucination-research)"})


def _get(url: str, params: dict | None = None,
         allow_404: bool = False, sleep: float = SLEEP_DEFAULT) -> dict | None:
    for attempt in range(MAX_RETRIES):
        try:
            r = _S.get(url, params=params, timeout=TIMEOUT)
        except (requests.ConnectionError,
                requests.exceptions.ReadTimeout,
                requests.exceptions.Timeout):
            # Netzwerkunterbrechung oder Timeout → kurz warten und erneut versuchen
            wait = min(2 ** attempt, 30)
            print(f"    [Netzwerkfehler] warte {wait}s und versuche erneut …")
            time.sleep(wait)
            continue

        if r.status_code == 429:
            wait = min(float(r.headers.get("Retry-After", 2 ** attempt)), 60)
            print(f"    [429] warte {wait:.0f}s …")
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
# HILFSFUNKTION: Titelfilter
# ─────────────────────────────────────────────────────────────────────────────

def is_valid_title(title: str) -> bool:
    if not title or title in SKIP_EXACT:
        return False
    if any(title.startswith(p) for p in SKIP_PREFIXES):
        return False
    if any(p in title for p in SKIP_CONTAINS):
        return False
    return True


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 1a: ZUFALLSARTIKEL
# ─────────────────────────────────────────────────────────────────────────────

def get_random_articles(n: int = 25) -> list[str]:
    data = _get(WIKIPEDIA_API, {
        "action": "query", "format": "json",
        "list": "random", "rnnamespace": 0, "rnlimit": n,
    })
    if not data:
        return []
    return [item["title"] for item in data.get("query", {}).get("random", [])]


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 1b: PAGEVIEW-PRÜFUNG (2021–2023 in einem Call)
# ─────────────────────────────────────────────────────────────────────────────

def get_avg_monthly_views(title: str) -> float | None:
    """
    Gibt den Durchschnitt der monatlichen Aufrufe 2021–2023 zurück.
    None wenn weniger als REQUIRE_N_MONTHS Datenpunkte vorhanden.
    """
    encoded = quote(title.replace(" ", "_"), safe="")
    url = (
        f"{PAGEVIEWS_API}/metrics/pageviews/per-article/"
        f"en.wikipedia/all-access/user/{encoded}/monthly/"
        f"{PAGEVIEW_START}/{PAGEVIEW_END}"
    )
    data = _get(url, allow_404=True)
    if not data:
        return None
    items = data.get("items", [])
    if len(items) < REQUIRE_N_MONTHS:
        return None
    return sum(i.get("views", 0) for i in items) / len(items)


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 1c: ENRICHMENT (Links, Backlinks, Kategorien)
# ─────────────────────────────────────────────────────────────────────────────

def get_links(title: str, limit: int = LINKS_PER_FOUND) -> list[str]:
    """Interne Links aus einem Artikel (vorwärts)."""
    collected, cont = [], None
    while len(collected) < limit:
        params: dict = {
            "action": "query", "format": "json",
            "titles": title, "prop": "links",
            "plnamespace": 0, "pllimit": 500,
        }
        if cont:
            params["plcontinue"] = cont
        data = _get(WIKIPEDIA_API, params)
        if not data:
            break
        for page in data.get("query", {}).get("pages", {}).values():
            for lnk in page.get("links", []):
                t = lnk.get("title", "")
                if t and is_valid_title(t):
                    collected.append(t)
        cont = data.get("continue", {}).get("plcontinue")
        if not cont or len(collected) >= limit:
            break
        time.sleep(SLEEP_DEFAULT)
    random.shuffle(collected)
    return collected[:limit]


def get_backlinks(title: str, limit: int = BACKLINKS_PER_FOUND) -> list[str]:
    """Artikel, die auf diesen Artikel verlinken."""
    collected, cont = [], None
    while len(collected) < limit:
        params: dict = {
            "action": "query", "format": "json",
            "list": "backlinks", "bltitle": title,
            "blnamespace": 0, "bllimit": 500,
        }
        if cont:
            params["blcontinue"] = cont
        data = _get(WIKIPEDIA_API, params)
        if not data:
            break
        for item in data.get("query", {}).get("backlinks", []):
            t = item.get("title", "")
            if t and is_valid_title(t):
                collected.append(t)
        cont = data.get("continue", {}).get("blcontinue")
        if not cont or len(collected) >= limit:
            break
        time.sleep(SLEEP_DEFAULT)
    random.shuffle(collected)
    return collected[:limit]


def get_categories_of(title: str, limit: int = 5) -> list[str]:
    """Kategorien eines Artikels (gefiltert auf inhaltliche Kategorien)."""
    data = _get(WIKIPEDIA_API, {
        "action": "query", "format": "json",
        "titles": title, "prop": "categories", "cllimit": 20,
    })
    cats = []
    if not data:
        return cats
    for page in data.get("query", {}).get("pages", {}).values():
        for cat in page.get("categories", []):
            name = cat.get("title", "").replace("Category:", "").strip()
            # Wartungskategorien etc. ausschließen
            if any(kw in name.lower() for kw in (
                "stub", "birth", "death", "living people", "articles",
                "pages ", "wikipedia ", "cs1 ", "use ", "all ", "accuracy"
            )):
                continue
            if name:
                cats.append(name)
            if len(cats) >= limit:
                break
    return cats


def get_category_members(category: str, limit: int = CATMEMBERS_PER_FOUND) -> list[str]:
    """Alle Artikel in einer Wikipedia-Kategorie."""
    collected, cont = [], None
    while len(collected) < limit:
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
            t = item.get("title", "")
            if t and is_valid_title(t):
                collected.append(t)
        cont = data.get("continue", {}).get("cmcontinue")
        if not cont or len(collected) >= limit:
            break
        time.sleep(SLEEP_DEFAULT)
    random.shuffle(collected)
    return collected[:limit]


def enrich_queue(found_title: str,
                 queue: deque[str],
                 already_seen: set[str]) -> None:
    """
    Fügt Links, Backlinks und Kategorien-Mitglieder eines gefundenen
    Nischenartikels in die Queue ein.
    """
    # Vorwärtslinks
    try:
        links = get_links(found_title)
        for t in links:
            if t not in already_seen:
                queue.append(t)
        time.sleep(SLEEP_DEFAULT)
    except Exception:
        pass

    # Backlinks
    try:
        bls = get_backlinks(found_title)
        for t in bls:
            if t not in already_seen:
                queue.append(t)
        time.sleep(SLEEP_DEFAULT)
    except Exception:
        pass

    # Kategorien → Mitglieder
    try:
        cats = get_categories_of(found_title)
        for cat in cats:
            members = get_category_members(cat)
            for t in members:
                if t not in already_seen:
                    queue.append(t)
            time.sleep(SLEEP_DEFAULT)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# PHASE 2: WIKIDATA-TRIPEL
# ─────────────────────────────────────────────────────────────────────────────

def get_qid(title: str) -> str | None:
    data = _get(WIKIPEDIA_API, {
        "action": "query", "format": "json",
        "titles": title, "prop": "pageprops",
        "ppprop": "wikibase_item", "redirects": 1,
    }, sleep=SLEEP_WIKIDATA)
    if not data:
        return None
    for page in data.get("query", {}).get("pages", {}).values():
        return page.get("pageprops", {}).get("wikibase_item")
    return None


def get_entity(qid: str) -> dict | None:
    data = _get(WIKIDATA_API, {
        "action": "wbgetentities", "format": "json", "ids": qid,
        "props": "labels|descriptions|claims", "languages": "en",
    }, sleep=SLEEP_WIKIDATA)
    if not data:
        return None
    return data.get("entities", {}).get(qid)


_label_cache: dict[str, str] = {}


def resolve_labels_batch(qids: list[str]) -> None:
    missing = [q for q in qids if q and q.startswith("Q") and q not in _label_cache]
    if not missing:
        return
    for i in range(0, len(missing), 50):
        batch = missing[i:i + 50]
        data = _get(WIKIDATA_API, {
            "action": "wbgetentities", "format": "json",
            "ids": "|".join(batch), "props": "labels", "languages": "en",
        }, sleep=SLEEP_WIKIDATA)
        if not data:
            continue
        for qid, ent in data.get("entities", {}).items():
            lbl = ent.get("labels", {}).get("en", {}).get("value")
            _label_cache[qid] = lbl if lbl else qid
        time.sleep(SLEEP_WIKIDATA)


def resolve_prop_labels_batch(pids: list[str]) -> None:
    missing = [p for p in pids if p and p.startswith("P") and p not in _label_cache]
    if not missing:
        return
    for i in range(0, len(missing), 50):
        batch = missing[i:i + 50]
        data = _get(WIKIDATA_API, {
            "action": "wbgetentities", "format": "json",
            "ids": "|".join(batch), "props": "labels", "languages": "en",
        }, sleep=SLEEP_WIKIDATA)
        if not data:
            continue
        for pid, ent in data.get("entities", {}).items():
            lbl = ent.get("labels", {}).get("en", {}).get("value")
            _label_cache[pid] = lbl if lbl else pid
        time.sleep(SLEEP_WIKIDATA)


def snak_to_value(snak: dict) -> tuple[str, str]:
    """
    Gibt (wert_als_string, wert_typ) zurück.
    wert_typ: "item" | "time" | "string" | "quantity" | "other"
    """
    datatype  = snak.get("datatype", "")
    datavalue = snak.get("datavalue", {})
    vtype     = datavalue.get("type", "")
    value     = datavalue.get("value", "")

    if vtype == "wikibase-entityid":
        return value.get("id", ""), "item"

    if vtype == "time":
        ts        = str(value.get("time", "")).lstrip("+").rstrip("Z")
        precision = value.get("precision", 11)
        try:
            parts = ts.split("T")[0].split("-")
            if precision == 9:
                return parts[0], "time"
            elif precision == 10:
                return f"{parts[0]}-{parts[1]}", "time"
            else:
                return f"{parts[0]}-{parts[1]}-{parts[2]}", "time"
        except Exception:
            return ts, "time"

    if vtype == "string":
        return str(value), "string"

    if vtype == "quantity":
        amount = str(value.get("amount", "")).lstrip("+")
        return amount, "quantity"

    if vtype == "monolingualtext":
        return value.get("text", ""), "string"

    return str(value), "other"


def extract_triples(entity: dict) -> list[dict]:
    """
    Extrahiert alle nutzbaren Tripel aus einer Wikidata-Entität.
    Löst Property-Labels und Item-Labels auf.
    """
    claims = entity.get("claims", {})
    raw: list[dict] = []

    for prop_id, stmts in claims.items():
        if prop_id in SKIP_PROPERTIES:
            continue
        for stmt in stmts:
            if stmt.get("rank") == "deprecated":
                continue
            snak = stmt.get("mainsnak", {})
            if snak.get("snaktype") != "value":
                continue
            val_raw, val_type = snak_to_value(snak)
            if not val_raw:
                continue
            raw.append({
                "property_id": prop_id,
                "value_raw":   val_raw,
                "value_type":  val_type,
                "rank":        stmt.get("rank", "normal"),
            })

    if not raw:
        return []

    # Labels auflösen (gebatcht)
    pids  = list({t["property_id"] for t in raw})
    items = [t["value_raw"] for t in raw if t["value_type"] == "item"]
    resolve_prop_labels_batch(pids)
    resolve_labels_batch(items)

    result: list[dict] = []
    seen_props: set[str] = set()
    for t in raw:
        pid = t["property_id"]
        prop_label  = _label_cache.get(pid, pid)
        if t["value_type"] == "item":
            value_label = _label_cache.get(t["value_raw"], t["value_raw"])
            value_id    = t["value_raw"]
        else:
            value_label = t["value_raw"]
            value_id    = ""

        # Pro Property nur einen Wert (bevorzuge "preferred"-Rank)
        rank = t["rank"]
        if pid in seen_props:
            if rank == "preferred":
                # Ersetzen
                result = [x for x in result if x["property_id"] != pid]
                seen_props.discard(pid)
            else:
                continue

        result.append({
            "property_id":    pid,
            "property_label": prop_label,
            "value_label":    value_label,
            "value_id":       value_id,
            "value_type":     t["value_type"],
            "rank":           rank,
        })
        seen_props.add(pid)

    return result


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
    # JSON
    Path(OUTPUT_JSON).write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # CSV (flach: eine Zeile pro Tripel)
    rows: list[dict] = []
    for article, info in results.items():
        base = {
            "article":           article,
            "wikidata_label":    info.get("label", article),
            "qid":               info.get("qid", ""),
            "avg_monthly_views": info.get("avg_monthly_views", ""),
            "months_with_data":  info.get("months_with_data", ""),
            "wikipedia_url":     f"https://en.wikipedia.org/wiki/{article.replace(' ', '_')}",
            "wikidata_url":      (f"https://www.wikidata.org/wiki/{info['qid']}"
                                  if info.get("qid") else ""),
        }
        triples = info.get("triples", [])
        if not triples:
            rows.append({**base, "property_id": "", "property_label": "",
                         "value_label": "", "value_id": "", "value_type": ""})
        else:
            for t in triples:
                rows.append({**base,
                    "property_id":    t["property_id"],
                    "property_label": t["property_label"],
                    "value_label":    t["value_label"],
                    "value_id":       t["value_id"],
                    "value_type":     t["value_type"],
                })

    pd.DataFrame(rows).to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    n_art  = len(results)
    n_prop = sum(1 for v in results.values() if v.get("triples"))
    n_trip = sum(len(v.get("triples", [])) for v in results.values())
    print(f"  → {n_art} Artikel  |  {n_prop} mit Tripeln  |  {n_trip} Tripel gesamt")


# ─────────────────────────────────────────────────────────────────────────────
# HAUPTPROGRAMM
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target",    type=int, default=TARGET_COUNT,
                        help=f"Zielanzahl Artikel (Default: {TARGET_COUNT})")
    parser.add_argument("--max-views", type=int, default=MAX_MONTHLY_VIEWS,
                        help=f"Max. Ø Views/Monat (Default: {MAX_MONTHLY_VIEWS})")
    parser.add_argument("--min-views", type=int, default=MIN_MONTHLY_VIEWS,
                        help=f"Min. Ø Views/Monat (Default: {MIN_MONTHLY_VIEWS})")
    args = parser.parse_args()

    max_views = args.max_views
    min_views = args.min_views
    target    = args.target

    print("=" * 65)
    print("Long-Tail Facts – Artikel- und Tripel-Finder")
    print(f"Ziel:         {target} Artikel")
    print(f"Views/Monat:  {min_views} – {max_views}  (Ø 2021–2023)")
    print("=" * 65)

    # Fortschritt laden
    progress: dict = load_progress()
    results:  dict = {k: v for k, v in progress.items()
                      if v.get("status") == "found"}
    checked:  set[str] = set(progress.keys())
    enriched: set[str] = set()

    print(f"\nBereits gefunden:    {len(results)}")
    print(f"Bereits geprüft:     {len(checked)}")

    # Queue mit bereits gefundenen Artikeln vorwärmen (Enrichment)
    queue: deque[str] = deque()
    for art in list(results.keys())[:10]:
        if art not in enriched:
            print(f"  Enrichment-Seed: {art}")
            enrich_queue(art, queue, checked)
            enriched.add(art)

    candidates_checked = 0

    # ── Hauptschleife ─────────────────────────────────────────────────────────
    while len(results) < target and candidates_checked < MAX_CANDIDATES:

        # Queue auffüllen wenn leer
        if len(queue) < RANDOM_BATCH_SIZE:
            randoms = get_random_articles(RANDOM_BATCH_SIZE)
            for t in randoms:
                if is_valid_title(t) and t not in checked:
                    queue.append(t)
            time.sleep(SLEEP_DEFAULT)

        if not queue:
            time.sleep(1)
            continue

        title = queue.popleft()
        if title in checked:
            continue
        if not is_valid_title(title):
            checked.add(title)
            continue

        candidates_checked += 1

        # Pageview-Prüfung
        avg = get_avg_monthly_views(title)
        time.sleep(SLEEP_DEFAULT)

        if avg is None:
            progress[title] = {"status": "no_data"}
            checked.add(title)
            if candidates_checked % 100 == 0:
                save_progress(progress)
            continue

        if avg < min_views or avg > max_views:
            progress[title] = {"status": "out_of_range", "avg": round(avg, 1)}
            checked.add(title)
            if candidates_checked % 100 == 0:
                save_progress(progress)
            continue

        # ── Artikel qualifiziert → Wikidata-Tripel holen ─────────────────
        print(f"\n  ✓ [{len(results)+1:>3}] {title}  (Ø {avg:,.1f} Views/Monat)")

        # Enrichment für Folge-Kandidaten
        # (nur jeden ENRICH_EVERY_N-ten Treffer, um API-Calls zu sparen)
        if title not in enriched and len(results) % ENRICH_EVERY_N == 0:
            enrich_queue(title, queue, checked)
            enriched.add(title)

        # QID
        qid = get_qid(title)
        time.sleep(SLEEP_WIKIDATA)

        if not qid:
            print(f"         → keine Wikidata-QID")
            entry = {
                "status":            "found",
                "qid":               None,
                "label":             title,
                "avg_monthly_views": round(avg, 1),
                "triples":           [],
                "error":             "no_qid",
            }
            results[title]  = entry
            progress[title] = entry
            checked.add(title)
            save_progress(progress)
            continue

        # Entity-Daten
        entity = get_entity(qid)
        time.sleep(SLEEP_WIKIDATA)

        if not entity:
            print(f"         → Entity-Daten nicht abrufbar")
            entry = {
                "status":            "found",
                "qid":               qid,
                "label":             title,
                "avg_monthly_views": round(avg, 1),
                "triples":           [],
                "error":             "entity_failed",
            }
            results[title]  = entry
            progress[title] = entry
            checked.add(title)
            save_progress(progress)
            continue

        label = entity.get("labels", {}).get("en", {}).get("value", title)
        desc  = entity.get("descriptions", {}).get("en", {}).get("value", "")

        triples = extract_triples(entity)

        entry = {
            "status":            "found",
            "qid":               qid,
            "label":             label,
            "description":       desc,
            "avg_monthly_views": round(avg, 1),
            "triples":           triples,
        }
        results[title]  = entry
        progress[title] = entry
        checked.add(title)

        if triples:
            sample = ", ".join(
                f"{t['property_label']}={t['value_label']}"
                for t in triples[:3]
            )
            print(f"         QID={qid}  {len(triples)} Tripel: {sample}")
        else:
            print(f"         QID={qid}  keine nutzbaren Tripel")

        save_progress(progress)

        # Zwischenausgabe alle 20 gefundenen Artikel
        if len(results) % 20 == 0:
            write_outputs(results)
            print(f"\n  [Zwischenspeicherung: {len(results)} Artikel]")

        time.sleep(SLEEP_DEFAULT)

    # ── Finale Ausgabe ────────────────────────────────────────────────────────
    print(f"\n[Finale Ausgabe …]")
    write_outputs(results)
    save_progress(progress)

    # Statistik
    print("\n" + "=" * 65)
    print("FERTIG")
    print(f"  Gefundene Artikel:          {len(results)}")
    print(f"  Davon mit Tripeln:          "
          f"{sum(1 for v in results.values() if v.get('triples'))}")
    print(f"  Gesamte Tripel:             "
          f"{sum(len(v.get('triples', [])) for v in results.values())}")
    print(f"  Kandidaten geprüft:         {candidates_checked}")
    print()

    # Häufigste Properties
    prop_counts: dict[str, int] = {}
    for v in results.values():
        for t in v.get("triples", []):
            lbl = t["property_label"]
            prop_counts[lbl] = prop_counts.get(lbl, 0) + 1

    print("  Häufigste Properties:")
    for lbl, cnt in sorted(prop_counts.items(), key=lambda x: -x[1])[:15]:
        print(f"    {lbl:<35} {cnt:>4}×")

    print(f"\n  CSV:  {OUTPUT_CSV}")
    print(f"  JSON: {OUTPUT_JSON}")
    print("=" * 65)

    if len(results) < target:
        print(
            f"\n  Hinweis: Nur {len(results)} von {target} Artikeln gefunden.\n"
            f"  Skript neu starten – Fortschritt wird aus {PROGRESS_JSON} geladen."
        )


if __name__ == "__main__":
    main()
