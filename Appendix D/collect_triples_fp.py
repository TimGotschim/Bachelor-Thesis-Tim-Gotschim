"""
collect_triples_fp.py
=====================
Bachelorarbeit - Tim Gotschim | WU Wien
Supervisor: Dr. Svitlana Vakulenko

Findet 300 einzigartige Wikipedia-Artikel mit 80.000-180.000 monatlichen
Aufrufen (Mittelwert 2022+2023) und extrahiert ALLE verfuegbaren
Wikidata-Tripel pro Artikel.

Das Ergebnis enthaelt mehrere Tripel pro Artikel, um spaeter manuell das
geeignetste fuer den Fragenkatalog auswaehlen zu koennen.

WORKFLOW:
  Phase 1: Vorhandene articles.csv einlesen (kein erneuter Pageview-Check)
  Phase 2: Fehlende Artikel via SPARQL + Pageview-API suchen

STOP-BEDINGUNG: 300 einzigartige Artikel (nicht 300 Tripel)

INSTALLATION:  pip install requests
AUSFUEHRUNG:
  python3 collect_triples_fp.py --articles articles.csv
  python3 collect_triples_fp.py --articles articles.csv --target 300

OPTIONEN:
  --articles FILE   Vorhandene articles.csv (Phase 1)
  --target N        Ziel: einzigartige Artikel (Default: 300)
  --out DIR         Ausgabeverzeichnis (Default: fp_output)
"""

import argparse
import csv
import json
import os
import random
import time
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Set, Tuple
from urllib.parse import quote

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

YEAR          = 2023
MIN_AVG_VIEWS = 80_000
MAX_AVG_VIEWS = 180_000

TARGET_PROPERTIES: Dict[str, str] = {
    "P19":  "place of birth",
    "P20":  "place of death",
    "P27":  "country of citizenship",
    "P106": "occupation",
    "P108": "employer",
    "P69":  "educated at",
    "P57":  "director",
    "P50":  "author",
    "P495": "country of origin",
    "P17":  "located in country",
    "P136": "genre",
    "P37":  "official language",
    "P102": "member of political party",
    "P101": "field of work",
    "P407": "language of work",
}

MIN_FALSE_OBJ_SITELINKS = 10

SPARQL_ENDPOINT   = "https://query.wikidata.org/sparql"
WIKIDATA_API      = "https://www.wikidata.org/w/api.php"
WIKIMEDIA_PV_BASE = (
    "https://wikimedia.org/api/rest_v1/metrics/pageviews/per-article"
    "/en.wikipedia/all-access/all-agents/{article}/monthly/{start}/{end}"
)
PAGEVIEW_YEARS = [
    ("20220101", "20221231"),
    ("20230101", "20231231"),
]

SLEEP_API    = 0.3
SLEEP_SPARQL = 1.5

ENTITY_QUERIES = [
    ("athletes", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q5 .
  ?subject wdt:P106 wd:Q2066131 .
  ?subject wdt:P569 ?birth .
  FILTER(YEAR(?birth) > 1975 && YEAR(?birth) < 2000)
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 500"""),
    ("musicians", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q5 .
  ?subject wdt:P106 wd:Q639669 .
  ?subject wdt:P569 ?birth .
  FILTER(YEAR(?birth) > 1970 && YEAR(?birth) < 2000)
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 500"""),
    ("actors", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q5 .
  ?subject wdt:P106 wd:Q33999 .
  ?subject wdt:P569 ?birth .
  FILTER(YEAR(?birth) > 1970 && YEAR(?birth) < 1995)
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 500"""),
    ("politicians", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q5 .
  ?subject wdt:P106 wd:Q82955 .
  ?subject wdt:P569 ?birth .
  FILTER(YEAR(?birth) > 1960 && YEAR(?birth) < 1985)
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 400"""),
    ("films", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q11424 .
  ?subject wdt:P577 ?release .
  FILTER(YEAR(?release) > 2000 && YEAR(?release) < 2020)
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 500"""),
    ("scientists", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q5 .
  ?subject wdt:P106 wd:Q901 .
  ?subject wdt:P569 ?birth .
  FILTER(YEAR(?birth) > 1940 && YEAR(?birth) < 1980)
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 300"""),
    ("tv_series", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q5398426 .
  ?subject wdt:P495 ?country .
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 300"""),
    ("books", """
SELECT DISTINCT ?subject WHERE {
  ?subject wdt:P31 wd:Q7725634 .
  ?subject wdt:P50 ?author .
  [] schema:about ?subject ; schema:isPartOf <https://en.wikipedia.org/> .
} LIMIT 300"""),
]

# ─────────────────────────────────────────────────────────────────────────────
# DATENKLASSEN
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ArticleStats:
    title:          str
    avg_views_2023: int
    wikidata_qid:   str
    wikipedia_url:  str

@dataclass
class TripleRecord:
    article_title:      str
    avg_views_2023:     int
    wikipedia_url:      str
    subject_qid:        str
    subject_label:      str
    property_id:        str
    property_label:     str
    object_qid_true:    str
    object_label_true:  str
    object_qid_false:   str
    object_label_false: str
    wikidata_url:       str

# ─────────────────────────────────────────────────────────────────────────────
# HTTP SESSION
# ─────────────────────────────────────────────────────────────────────────────

def build_session() -> requests.Session:
    s = requests.Session()
    retries = Retry(total=5, backoff_factor=1.5,
                    status_forcelist=[429, 500, 502, 503, 504],
                    allowed_methods=["GET"])
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://",  HTTPAdapter(max_retries=retries))
    s.headers.update({"User-Agent": "BachelorThesisFP/4.0 (tim.gotschim@wu.ac.at)"})
    return s

SESSION = build_session()

# ─────────────────────────────────────────────────────────────────────────────
# API-HILFSFUNKTIONEN
# ─────────────────────────────────────────────────────────────────────────────

def run_sparql(query: str, timeout: int = 40) -> list:
    for attempt in range(1, 4):
        try:
            resp = SESSION.get(
                SPARQL_ENDPOINT,
                params={"query": query, "format": "json"},
                headers={"Accept": "application/sparql-results+json"},
                timeout=timeout,
            )
            if resp.status_code == 429:
                wait = 20 * attempt
                print(f"    [RATE LIMIT] Warte {wait}s ...")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()["results"]["bindings"]
        except requests.exceptions.Timeout:
            print(f"    [TIMEOUT] Versuch {attempt}/3")
            time.sleep(10 * attempt)
        except Exception as e:
            print(f"    [SPARQL ERROR] {e}")
            if attempt < 3:
                time.sleep(5)
    return []


def get_wikipedia_title(qid: str) -> str:
    try:
        resp = SESSION.get(WIKIDATA_API, params={
            "action": "wbgetentities", "ids": qid,
            "props": "sitelinks", "sitefilter": "enwiki", "format": "json",
        }, timeout=15)
        resp.raise_for_status()
        entity = resp.json().get("entities", {}).get(qid, {})
        return entity.get("sitelinks", {}).get("enwiki", {}).get("title", "")
    except Exception:
        return ""


def get_qid_from_title(title: str) -> str:
    try:
        resp = SESSION.get("https://en.wikipedia.org/w/api.php", params={
            "action": "query", "titles": title, "prop": "pageprops",
            "ppprop": "wikibase_item", "redirects": "1",
            "format": "json", "formatversion": "2",
        }, timeout=15)
        resp.raise_for_status()
        pages = resp.json().get("query", {}).get("pages", [])
        if pages:
            return pages[0].get("pageprops", {}).get("wikibase_item", "")
    except Exception:
        pass
    return ""


def get_avg_monthly_pageviews(title: str) -> int:
    """Mittelwert ueber 2022 und 2023."""
    if not title:
        return 0
    article   = quote(title.replace(" ", "_"), safe="")
    all_views = []
    for start, end in PAGEVIEW_YEARS:
        try:
            url  = WIKIMEDIA_PV_BASE.format(article=article, start=start, end=end)
            resp = SESSION.get(url, timeout=15)
            if resp.status_code == 404:
                continue
            resp.raise_for_status()
            items = resp.json().get("items", [])
            all_views.extend(i.get("views", 0) for i in items)
        except Exception:
            continue
    return sum(all_views) // len(all_views) if all_views else 0


def get_entity_data(qid: str) -> Tuple[Dict[str, List[str]], str]:
    """Gibt ({property_id: [obj_qid, ...]}, subject_label) zurueck."""
    try:
        resp = SESSION.get(WIKIDATA_API, params={
            "action": "wbgetentities", "ids": qid,
            "props": "claims|labels", "languages": "en", "format": "json",
        }, timeout=20)
        resp.raise_for_status()
        entity = resp.json().get("entities", {}).get(qid, {})
        label  = entity.get("labels", {}).get("en", {}).get("value", "")
        claims: Dict[str, List[str]] = {}
        for prop_id, claim_list in entity.get("claims", {}).items():
            if prop_id not in TARGET_PROPERTIES:
                continue
            obj_qids = []
            for claim in claim_list:
                snak = claim.get("mainsnak", {})
                if snak.get("snaktype") != "value":
                    continue
                dv = snak.get("datavalue", {})
                if dv.get("type") != "wikibase-entityid":
                    continue
                num = dv.get("value", {}).get("numeric-id")
                if num:
                    obj_qids.append(f"Q{num}")
            if obj_qids:
                claims[prop_id] = obj_qids
        return claims, label
    except Exception:
        return {}, ""


def batch_get_labels(qids: List[str]) -> Dict[str, str]:
    labels: Dict[str, str] = {}
    for i in range(0, len(qids), 50):
        batch = qids[i:i+50]
        try:
            resp = SESSION.get(WIKIDATA_API, params={
                "action": "wbgetentities", "ids": "|".join(batch),
                "props": "labels", "languages": "en", "format": "json",
            }, timeout=20)
            resp.raise_for_status()
            for qid, entity in resp.json().get("entities", {}).items():
                lbl = entity.get("labels", {}).get("en", {}).get("value", "")
                if lbl:
                    labels[qid] = lbl
        except Exception as e:
            print(f"    [WARN] Batch-Labels: {e}")
        time.sleep(SLEEP_API)
    return labels


def labels_too_similar(a: str, b: str) -> bool:
    a_l, b_l = a.lower().strip(), b.lower().strip()
    if a_l == b_l:
        return True
    if a_l in b_l or b_l in a_l:
        return True
    a_w, b_w = a_l.split(), b_l.split()
    if a_w and b_w and a_w[-1] == b_w[-1] and len(a_w) <= 3:
        return True
    return False


def find_plausible_false_object(
        true_qid: str, true_label: str,
        min_sitelinks: int = MIN_FALSE_OBJ_SITELINKS,
        n: int = 20) -> Optional[Tuple[str, str]]:
    query = f"""
SELECT DISTINCT ?candidate ?candidateLabel WHERE {{
  wd:{true_qid} wdt:P31 ?class .
  ?candidate wdt:P31 ?class .
  FILTER(?candidate != wd:{true_qid})
  ?candidate rdfs:label ?candidateLabel .
  FILTER(LANG(?candidateLabel) = "en")
}}
LIMIT {n * 5}
"""
    try:
        bindings = run_sparql(query, timeout=20)
        raw = []
        for b in bindings:
            lbl = b.get("candidateLabel", {}).get("value", "")
            qid = b.get("candidate",      {}).get("value", "").split("/")[-1]
            if (lbl and qid
                    and not (lbl.startswith("Q") and lbl[1:].isdigit())
                    and 2 < len(lbl) < 60
                    and not labels_too_similar(lbl, true_label)):
                raw.append((qid, lbl))
        if not raw:
            return None
        # Sitelinks pruefen
        try:
            resp = SESSION.get(WIKIDATA_API, params={
                "action": "wbgetentities",
                "ids": "|".join(c[0] for c in raw[:50]),
                "props": "sitelinks", "format": "json",
            }, timeout=15)
            resp.raise_for_status()
            entities = resp.json().get("entities", {})
            qualified = [(qid, lbl) for qid, lbl in raw[:50]
                         if len(entities.get(qid, {}).get("sitelinks", {})) >= min_sitelinks]
        except Exception:
            qualified = raw[:50]
        return random.choice(qualified[:n]) if qualified else None
    except Exception as e:
        print(f"    [WARN] False-Object: {e}")
        return None

# ─────────────────────────────────────────────────────────────────────────────
# KERN: ALLE TRIPEL EINES ARTIKELS EXTRAHIEREN
# ─────────────────────────────────────────────────────────────────────────────

def extract_all_triples(article: ArticleStats) -> List[TripleRecord]:
    """
    Extrahiert ALLE verfuegbaren Tripel fuer einen Artikel.
    Mehrere Tripel pro Artikel sind erwuenscht fuer spaeteren
    manuellen Auswahlspielraum.
    """
    claims, subject_label = get_entity_data(article.wikidata_qid)
    time.sleep(SLEEP_API)

    if not claims or not subject_label:
        return []

    all_obj_qids = [q for qids in claims.values() for q in qids]
    labels = batch_get_labels(all_obj_qids)

    triples = []
    for prop_id, obj_qids in claims.items():
        # Erstes Objekt mit lesbarem Label
        true_qid   = None
        true_label = None
        for oqid in obj_qids:
            lbl = labels.get(oqid, "")
            if lbl and not (lbl.startswith("Q") and lbl[1:].isdigit()):
                true_qid   = oqid
                true_label = lbl
                break

        if not true_qid:
            continue

        false_obj = find_plausible_false_object(true_qid, true_label)
        time.sleep(SLEEP_SPARQL)

        if not false_obj:
            continue

        false_qid, false_label = false_obj
        triples.append(TripleRecord(
            article_title=article.title,
            avg_views_2023=article.avg_views_2023,
            wikipedia_url=article.wikipedia_url,
            subject_qid=article.wikidata_qid,
            subject_label=subject_label,
            property_id=prop_id,
            property_label=TARGET_PROPERTIES[prop_id],
            object_qid_true=true_qid,
            object_label_true=true_label,
            object_qid_false=false_qid,
            object_label_false=false_label,
            wikidata_url=f"https://www.wikidata.org/wiki/{article.wikidata_qid}",
        ))

    return triples

# ─────────────────────────────────────────────────────────────────────────────
# AUSGABE
# ─────────────────────────────────────────────────────────────────────────────

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def save_articles_csv(articles: List[ArticleStats], filepath: str):
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["title", "avg_views_2023", "wikidata_qid", "wikipedia_url"])
        writer.writeheader()
        for a in articles:
            writer.writerow(asdict(a))

def save_triples_csv(records: List[TripleRecord], filepath: str):
    fields = ["article_title", "avg_views_2023", "wikipedia_url",
              "subject_qid", "subject_label", "property_id", "property_label",
              "object_qid_true", "object_label_true",
              "object_qid_false", "object_label_false", "wikidata_url"]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for rec in records:
            writer.writerow(asdict(rec))

def save_triples_jsonl(records: List[TripleRecord], filepath: str):
    with open(filepath, "w", encoding="utf-8") as f:
        for rec in records:
            f.write(json.dumps(asdict(rec), ensure_ascii=False) + "\n")

# ─────────────────────────────────────────────────────────────────────────────
# HAUPTPROGRAMM
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Findet 300 einzigartige Wikipedia-Artikel und extrahiert alle Tripel"
    )
    parser.add_argument("--articles", default=None,
        help="Vorhandene articles.csv (uebergeht Pageview-Check, empfohlen)")
    parser.add_argument("--target", type=int, default=300,
        help="Ziel: einzigartige Artikel (Default: 300)")
    parser.add_argument("--out", type=str, default="fp_output")
    args = parser.parse_args()

    ensure_dir(args.out)
    triples_csv  = os.path.join(args.out, "triples.csv")
    triples_jl   = os.path.join(args.out, "triples.jsonl")
    articles_out = os.path.join(args.out, "articles_processed.csv")

    # Verarbeitete Artikel und Tripel
    processed_articles: Set[str] = set()   # Titel der bereits verarbeiteten Artikel
    triple_records: List[TripleRecord] = []
    all_articles:   List[ArticleStats] = []

    print(f"=== Fictitious Premises Triple Collector ===")
    print(f"Ziel: {args.target} einzigartige Artikel")
    print(f"STOP-Bedingung: Anzahl einzigartiger Artikel >= {args.target}\n")

    def autosave():
        save_triples_csv(triple_records, triples_csv)
        save_triples_jsonl(triple_records, triples_jl)
        save_articles_csv(all_articles, articles_out)

    def process_article(article: ArticleStats) -> int:
        """Verarbeitet einen Artikel, gibt Anzahl neuer Tripel zurueck."""
        if article.title in processed_articles:
            return 0

        print(f"  {article.title}  [{article.avg_views_2023:,}/Monat]")

        new_triples = extract_all_triples(article)
        if new_triples:
            triple_records.extend(new_triples)
            processed_articles.add(article.title)
            all_articles.append(article)
            for t in new_triples:
                print(f"    [{t.property_label}] '{t.object_label_true}' -> falsch: '{t.object_label_false}'")
            print(f"    {len(new_triples)} Tripel  |  "
                  f"Artikel gesamt: {len(processed_articles)}/{args.target}  "
                  f"Tripel gesamt: {len(triple_records)}")
            return len(new_triples)
        else:
            print(f"    Keine Tripel gefunden -> uebersprungen")
            return 0

    # ── Phase 1: Vorhandene articles.csv nutzen ───────────────────────────
    if args.articles and os.path.exists(args.articles):
        print(f"[Phase 1] Lese {args.articles} ...")
        with open(args.articles, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
        print(f"  {len(rows)} Artikel gefunden\n")

        for i, row in enumerate(rows, start=1):
            if len(processed_articles) >= args.target:
                break

            article = ArticleStats(
                title=row["title"],
                avg_views_2023=int(row.get("avg_views_2023", 0)),
                wikidata_qid=row.get("wikidata_qid", ""),
                wikipedia_url=row.get("wikipedia_url", ""),
            )

            # QID holen falls fehlend
            if not article.wikidata_qid:
                article.wikidata_qid = get_qid_from_title(article.title)
                time.sleep(SLEEP_API)
                if not article.wikidata_qid:
                    continue

            print(f"  ({i}/{len(rows)})", end=" ")
            process_article(article)

            if len(processed_articles) % 20 == 0 and processed_articles:
                autosave()
                print(f"  [Auto-Save]")

    # ── Phase 2: Neue Artikel via SPARQL falls Ziel nicht erreicht ────────
    if len(processed_articles) < args.target:
        still_needed = args.target - len(processed_articles)
        print(f"\n[Phase 2] Noch {still_needed} Artikel benoetigt. Suche via SPARQL ...\n")

        # QIDs sammeln
        all_qids = []
        seen_qids: Set[str] = set()
        for label, query in ENTITY_QUERIES:
            if len(all_qids) >= still_needed * 10:
                break
            print(f"  SPARQL: {label} ...")
            for b in run_sparql(query):
                qid = b.get("subject", {}).get("value", "").split("/")[-1]
                if qid and qid not in seen_qids:
                    seen_qids.add(qid)
                    all_qids.append(qid)
            time.sleep(SLEEP_SPARQL)

        random.shuffle(all_qids)
        checked = 0

        for qid in all_qids:
            if len(processed_articles) >= args.target:
                break

            title = get_wikipedia_title(qid)
            time.sleep(SLEEP_API)
            if not title or title in processed_articles:
                continue

            avg = get_avg_monthly_pageviews(title)
            time.sleep(SLEEP_API)
            checked += 1

            if checked % 30 == 0:
                print(f"  Geprueft: {checked} | "
                      f"Artikel: {len(processed_articles)}/{args.target} | "
                      f"Tripel: {len(triple_records)}")

            if not (MIN_AVG_VIEWS <= avg <= MAX_AVG_VIEWS):
                continue

            article = ArticleStats(
                title=title,
                avg_views_2023=avg,
                wikidata_qid=qid,
                wikipedia_url=f"https://en.wikipedia.org/wiki/{title.replace(' ','_')}",
            )
            print(f"  [NEU] ", end="")
            process_article(article)

            if len(processed_articles) % 20 == 0 and processed_articles:
                autosave()

    # ── Finale Ausgabe ────────────────────────────────────────────────────
    autosave()

    print(f"\n{'='*60}")
    print(f"Einzigartige Artikel : {len(processed_articles)}")
    print(f"Tripel gesamt        : {len(triple_records)}")
    print(f"Tripel pro Artikel   : {len(triple_records)/max(len(processed_articles),1):.1f} (Durchschnitt)")
    print(f"Gespeichert in       : {args.out}/")
    print(f"  - triples.csv")
    print(f"  - triples.jsonl")
    print(f"  - articles_processed.csv")

    if len(processed_articles) < args.target:
        print(f"\nHINWEIS: Nur {len(processed_articles)}/{args.target} Artikel gefunden.")
        print(f"Skript erneut ausfuehren um weitere Artikel zu suchen.")


if __name__ == "__main__":
    main()
