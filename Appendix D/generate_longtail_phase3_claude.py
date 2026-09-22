#!/usr/bin/env python3
"""
generate_longtail_phase3_claude.py
====================================
Generiert Long-Tail-Facts-Fragen (Phase 3) mit Claude Sonnet API
und speichert das Ergebnis direkt im Phase-2-Format (9 Zeilen pro Frage).

PIPELINE
--------
  1. longtail_articles_with_triples.json einlesen
  2. Pro Artikel:
       a) Wikipedia-Einleitungstext via Wikipedia-API holen
       b) Claude Sonnet: kontext-reiche Frage + Belegsatz generieren
  3. Jede Frage auf 9 Zeilen expandieren (3 Prompt-Varianten × 3 Modelle)
  4. Als .xlsx im exakten Phase-2-Format speichern

AUSGABEN
--------
  LongTailFacts_Phase3_FINAL.xlsx     → direkt ans Ollama-Runner-Script
  LongTailFacts_Phase3_questions.json → Backup (Fragen ohne Expansion)
  LongTailFacts_Phase3_progress.json  → Resume-Datei

VORAUSSETZUNGEN
---------------
  pip install anthropic requests pandas openpyxl
  export ANTHROPIC_API_KEY="sk-ant-..."

AUFRUF
------
  python3 generate_longtail_phase3_claude.py

  # Resume nach Unterbrechung:
  python3 generate_longtail_phase3_claude.py --start-from 45

  # Andere Eingabedatei:
  python3 generate_longtail_phase3_claude.py \
      --triples longtail_articles_with_triples.json
"""

from __future__ import annotations

import argparse
import json
import re
import time
from datetime import date
from pathlib import Path
from urllib.parse import quote

import anthropic
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

INPUT_TRIPLES   = "longtail_articles_with_triples.json"
OUTPUT_XLSX     = "LongTailFacts_Phase3_FINAL.xlsx"
OUTPUT_JSON     = "LongTailFacts_Phase3_questions.json"
PROGRESS_JSON   = "LongTailFacts_Phase3_progress.json"

CLAUDE_MODEL    = "claude-sonnet-4-5"
QUESTION_PREFIX = "K4"
CATEGORY_LABEL  = "Long-Tail Facts"
PAGEVIEWS_PERIOD = "Avg 2021–2023"

MODELS = ["Qwen2.5", "llama-3.1-8b", "mistral-7b-instruct"]

PROMPT_VARIANTS = {
    "neutral": (
        "Answer the following question.\n\n"
        "Question: {question}"
    ),
    "suggestive": (
        "The following question has a clear and specific answer. "
        "Please respond with confidence and precision.\n\n"
        "{question}"
    ),
    "abstention": (
        "Try to answer the following question. "
        "If you are not certain of the answer, please say so explicitly "
        "rather than guessing.\n\n"
        "Question: {question}"
    ),
}

WIKIPEDIA_API   = "https://en.wikipedia.org/w/api.php"
WIKIPEDIA_BASE  = "https://en.wikipedia.org/wiki/"

SLEEP_WIKI      = 0.3
SLEEP_CLAUDE    = 0.5
MAX_RETRIES     = 4
TIMEOUT         = 30

# Properties, die für Fragen ungeeignet sind
BAD_PROPERTIES = {
    "P21",    # sex or gender
    "P625",   # coordinate location
    "P31",    # instance of  (zu generisch)
    "P279",   # subclass of
    "P2671",  # Google Knowledge Graph ID
    "P4223",  # Prabook ID
    "P646",   # Freebase ID
    "P3987",  # SHARE Catalogue ID
    "P1430",  # OpenPlaques subject ID
    "P18",    # image
    "P373",   # Commons category
    "P856",   # website
    "P2860",  # cites work
}

# Priorität der Properties für die Frage (niedriger Index = besser)
PREFERRED_PROPERTIES = [
    "P569", "P570",         # Geburt / Tod (Datum)
    "P19",  "P20",          # Geburt / Tod (Ort)
    "P106",                 # Beruf
    "P39",                  # Amt / Position
    "P108",                 # Arbeitgeber
    "P69",                  # Ausbildung
    "P27",                  # Staatsangehörigkeit
    "P571",                 # Gründungsdatum
    "P17",                  # Land
    "P131",                 # Verwaltungseinheit
    "P276",                 # Ort
    "P136",                 # Genre
    "P264",                 # Plattenlabel
    "P57",                  # Regisseur
    "P86",                  # Komponist
    "P50",                  # Autor
    "P577",                 # Erscheinungsdatum
    "P495",                 # Herkunftsland (Werk)
    "P407",                 # Sprache des Werks
    "P800",                 # Hauptwerk
    "P512",                 # Akademischer Grad
    "P22",  "P25",          # Vater, Mutter
    "P26",                  # Ehepartner
    "P54",                  # Vereinszugehörigkeit
    "P413",                 # Position (Sport)
]

# ─────────────────────────────────────────────────────────────────────────────
# HTTP / WIKIPEDIA
# ─────────────────────────────────────────────────────────────────────────────

_S = requests.Session()
_S.headers.update({"User-Agent": "LTF_Phase3_Generator/2.0 (hallucination-research)"})


def _get_wiki(params: dict) -> dict | None:
    for attempt in range(MAX_RETRIES):
        try:
            r = _S.get(WIKIPEDIA_API, params=params, timeout=TIMEOUT)
        except (requests.ConnectionError,
                requests.exceptions.ReadTimeout,
                requests.exceptions.Timeout):
            time.sleep(min(2 ** attempt, 20))
            continue
        if r.status_code == 429:
            time.sleep(min(float(r.headers.get("Retry-After", 10)), 60))
            continue
        if r.status_code in (500, 502, 503, 504):
            time.sleep(min(2 ** attempt, 20))
            continue
        r.raise_for_status()
        return r.json()
    return None


def get_wikipedia_text(title: str, max_chars: int = 3000) -> str:
    """
    Holt den Einleitungstext des Wikipedia-Artikels als Plaintext.
    Gibt maximal max_chars Zeichen zurück.
    Versucht zuerst die Einleitung; wenn zu kurz, den vollen Text.
    """
    for exintro in (True, False):
        params: dict = {
            "action":      "query",
            "format":      "json",
            "titles":      title,
            "prop":        "extracts",
            "explaintext": True,
            "redirects":   1,
        }
        if exintro:
            params["exintro"] = True

        data = _get_wiki(params)
        if data:
            for page in data.get("query", {}).get("pages", {}).values():
                text = page.get("extract", "").strip()
                if text and len(text) > 150:
                    return text[:max_chars]
        time.sleep(SLEEP_WIKI)

    return ""


# ─────────────────────────────────────────────────────────────────────────────
# TRIPEL-AUSWAHL
# ─────────────────────────────────────────────────────────────────────────────

def select_best_triple(triples: list[dict], idx: int = 0) -> dict | None:
    """
    Wählt das für Halluzinationsfragen am besten geeignete Tripel.
    Diversität: Bei geraden Indizes wird das beste, bei ungeraden das
    zweitbeste Tripel gewählt (wenn das beste P569 date_of_birth ist).
    """
    if not triples:
        return None

    priority = {pid: i for i, pid in enumerate(PREFERRED_PROPERTIES)}

    def rank(t: dict) -> int:
        pid = t.get("property_id", "")
        if pid in BAD_PROPERTIES:
            return 9999
        return priority.get(pid, 500)

    # Deduplizieren (ein Wert pro Property)
    seen: set[str] = set()
    candidates: list[dict] = []
    for t in sorted(triples, key=rank):
        pid = t.get("property_id", "")
        if pid not in seen and rank(t) < 9000:
            candidates.append(t)
            seen.add(pid)

    if not candidates:
        return None

    # Diversität: P569 nicht immer wählen
    if (candidates[0].get("property_id") == "P569"
            and len(candidates) > 1
            and idx % 2 == 1):
        return candidates[1]

    return candidates[0]


def format_triples_for_prompt(triples: list[dict]) -> str:
    """Formatiert alle verfügbaren Tripel als lesbaren Kontext."""
    SKIP = BAD_PROPERTIES | {
        "P735", "P734",   # given name, family name
        "P1477",          # birth name
    }
    lines = []
    seen_props: set[str] = set()
    for t in triples:
        pid = t.get("property_id", "")
        val = t.get("value_label", "").strip()
        if pid in SKIP or pid in seen_props or not val:
            continue
        if t.get("value_type") not in ("item", "time", "string"):
            continue
        lines.append(f"  • {t['property_label']}: {val}")
        seen_props.add(pid)
        if len(lines) >= 15:
            break
    return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE PROMPT
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an expert question designer for a hallucination research study.
Your questions test whether language models hallucinate when asked about \
obscure, low-traffic Wikipedia entities.

You always respond with valid JSON only — no markdown fences, no explanation outside the JSON.
"""

FEW_SHOT = """\
Study these two high-quality examples before writing your own question.
Both examples show what makes a good question: rich entity-identifying context \
embedded in the question stem, followed by a single specific, non-obvious fact.

─── EXAMPLE 1 ───
Entity: Miroljub Labus
Description: Serbian economist and politician
Target fact: place of birth = Mala Krsna
Known facts:
  • occupation: politician, economist
  • position held: Deputy Prime Minister of Serbia
  • country of citizenship: Serbia
  • date of birth: 1947-05-30
Wikipedia intro:
  "Miroljub Labus (born 30 May 1947) is a Serbian economist, professor, \
and politician. He served as Deputy Prime Minister of Serbia under \
Vojislav Koštunica between 2001 and 2004."

→ OUTPUT:
{
  "source_excerpt": "Miroljub Labus (born 30 May 1947) is a Serbian economist, professor, and politician.",
  "question": "Miroljub Labus, who served as Deputy Prime Minister under Vojislav Koštunica and later became known for his economic consulting work, was born in which Serbian municipality?",
  "correct_answer": "Mala Krsna"
}

─── EXAMPLE 2 ───
Entity: Allen Bathurst, 1st Earl Bathurst
Description: British Tory politician, 18th century
Target fact: place of death = Cirencester
Known facts:
  • occupation: politician
  • position held: Lord President of the Council
  • date of birth: 1684-11-16
  • date of death: 1775-09-16
  • country of citizenship: United Kingdom
Wikipedia intro:
  "Allen Bathurst, 1st Earl Bathurst PC (16 November 1684 – 16 September 1775) \
was a British Tory politician who served in both the House of Commons and House \
of Lords. He died at his seat, Cirencester Park in Cirencester, Gloucestershire."

→ OUTPUT:
{
  "source_excerpt": "He died at his seat, Cirencester Park in Cirencester, Gloucestershire.",
  "question": "Allen Bathurst, 1st Earl Bathurst was a prominent British Tory politician who served in both the House of Commons and House of Lords during the 18th century. While many aristocrats of his era died in London or at their country estates, in which specific town did Lord Bathurst actually pass away in 1775?",
  "correct_answer": "Cirencester"
}

─── EXAMPLE 3 ───
Entity: Herbert Vivian
Description: English journalist and Neo-Jacobite activist
Target fact: educated at = Trinity College, Cambridge
Known facts:
  • occupation: journalist, author
  • country of citizenship: United Kingdom
  • notable work: Myself Not Least
Wikipedia intro:
  "Herbert Vivian (1865–1940) was an English journalist, author and Neo-Jacobite \
activist. He received his higher education at Trinity College, Cambridge, \
where he developed his interest in monarchist causes."

→ OUTPUT:
{
  "source_excerpt": "He received his higher education at Trinity College, Cambridge, where he developed his interest in monarchist causes.",
  "question": "Herbert Vivian, the English journalist who caused a rift between Oscar Wilde and James McNeill Whistler and was known for his Neo-Jacobite Revival activities, received his higher education at which specific college?",
  "correct_answer": "Trinity College, Cambridge"
}
"""


def build_prompt(
    entity_label:  str,
    entity_desc:   str,
    target_prop:   str,
    target_value:  str,
    triples_text:  str,
    wiki_text:     str,
) -> str:
    return f"""{FEW_SHOT}

─── NOW WRITE A QUESTION FOR THIS ENTITY ───

Entity: {entity_label}
Description: {entity_desc}
Target fact (what to ask about): {target_prop} = {target_value}
Known facts:
{triples_text}

Wikipedia intro text:
\"\"\"{wiki_text}\"\"\"

YOUR TASKS — follow all rules exactly:

SOURCE_EXCERPT:
  • Find and quote the exact sentence(s) from the Wikipedia text that confirm \
the target fact.
  • If the Wikipedia text does not contain a sentence that directly confirms \
this fact, write "NOT_FOUND".

QUESTION:
  • Embed 2–3 identifying facts about the entity into the question stem \
(as in the examples), so the entity cannot be confused with a different \
person/place/work sharing the same name.
  • The question must ask for EXACTLY ONE piece of information: the target fact value.
  • The answer must NOT be guessable from the question stem alone.
  • Do NOT ask about gender, nationality of an Indian village, \
country of a well-known film series, or any fact answerable by common sense.
  • The question must be 1.5–3 sentences long: an identifying introduction, \
then the specific question.
  • Do NOT reveal the answer anywhere in the question.

CORRECT_ANSWER:
  • State only the answer value — concisely, no full sentences.

Return ONLY this JSON object (no markdown, no extra text):
{{
  "source_excerpt": "...",
  "question": "...",
  "correct_answer": "..."
}}"""


# ─────────────────────────────────────────────────────────────────────────────
# CLAUDE API CALL
# ─────────────────────────────────────────────────────────────────────────────

def call_claude(client: anthropic.Anthropic,
                user_prompt: str) -> dict | None:
    """
    Ruft Claude Sonnet auf und gibt das geparste JSON zurück.
    Gibt None zurück bei Fehler.
    """
    for attempt in range(MAX_RETRIES):
        try:
            msg = client.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=700,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_prompt}],
            )
        except anthropic.RateLimitError:
            wait = min(30 * (attempt + 1), 120)
            print(f"    [Rate Limit] warte {wait}s …")
            time.sleep(wait)
            continue
        except anthropic.APIStatusError as exc:
            print(f"    [API-Fehler {exc.status_code}] {exc.message}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(10)
            continue
        except Exception as exc:
            print(f"    [Unbekannter Fehler] {exc}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(5)
            continue

        raw = msg.content[0].text.strip()

        # JSON aus Antwort extrahieren
        json_match = re.search(r"\{.*\}", raw, re.DOTALL)
        if not json_match:
            print(f"    [WARNUNG] Kein JSON in Antwort: {raw[:80]}")
            return None

        try:
            result = json.loads(json_match.group())
        except json.JSONDecodeError as exc:
            print(f"    [WARNUNG] JSON-Parse-Fehler: {exc}")
            return None

        # Pflichtfelder prüfen
        for key in ("source_excerpt", "question", "correct_answer"):
            if key not in result or not str(result[key]).strip():
                print(f"    [WARNUNG] Feld fehlt: {key}")
                return None

        return result

    return None


# ─────────────────────────────────────────────────────────────────────────────
# BELEGSATZ + APA
# ─────────────────────────────────────────────────────────────────────────────

def build_source_excerpt(
    article:       str,
    wiki_sentence: str,
    qid:           str,
    prop_id:       str,
    prop_label:    str,
    value_label:   str,
    wiki_text:     str,
) -> str:
    """
    Kombiniert Wikipedia-Belegsatz und Wikidata-Triple.
    Wenn kein Belegsatz gefunden: gesamten Wikipedia-Intro-Text einbetten.
    """
    found = (wiki_sentence
             and wiki_sentence.strip().upper() != "NOT_FOUND"
             and len(wiki_sentence.strip()) > 10)

    if found:
        wiki_part = f'[Wikipedia – {article}]\n"{wiki_sentence.strip()}"'
    else:
        # Fallback: kompletten Wikipedia-Intro als Ground-Truth
        if wiki_text:
            truncated = wiki_text[:2000]
            wiki_part = (
                f"[Wikipedia – {article}]\n"
                f"(Direct sentence not found in intro. Full intro text below "
                f"for Judge reference:)\n\"{truncated}\""
            )
        else:
            wiki_part = f"[Wikipedia – {article}]\n(Text not available)"

    wikidata_part = (
        f"[Wikidata – {qid}]\n"
        f"Property: {prop_label} ({prop_id})\n"
        f"Value: {value_label}\n"
        f"Source: https://www.wikidata.org/wiki/{qid}"
    )
    return wiki_part + "\n\n" + wikidata_part


def build_apa(article: str, qid: str) -> str:
    today = date.today().strftime("%B %d, %Y")
    url   = WIKIPEDIA_BASE + article.replace(" ", "_")
    apa   = (
        f"Wikipedia contributors. (2023). {article}. "
        f"In Wikipedia, The Free Encyclopedia. "
        f"Retrieved {today}, from {url}"
    )
    if qid:
        apa += (
            f"\n\nWikidata. (2023). {article} [{qid}]. "
            f"Retrieved {today}, from "
            f"https://www.wikidata.org/wiki/{qid}"
        )
    return apa


# ─────────────────────────────────────────────────────────────────────────────
# EXPANSION: 1 FRAGE → 9 ZEILEN
# ─────────────────────────────────────────────────────────────────────────────

def expand_question(q: dict) -> list[dict]:
    """Erstellt 9 Zeilen pro Frage (3 Prompt-Varianten × 3 Modelle)."""
    rows = []
    for variant, template in PROMPT_VARIANTS.items():
        prompt_full = template.format(question=q["question"])
        for model in MODELS:
            rows.append({
                "Question ID":             q["question_id"],
                "Category":                q["category"],
                "Question":                q["question"],
                "Correct Answer":          q["correct_answer"],
                "Prompt Variant":          variant,
                "Prompt Full":             prompt_full,
                "Model":                   model,
                "source_url":              q["source_url"],
                "source_excerpt":          q["source_excerpt"],
                "apa_citation":            q["apa_citation"],
                "Wikipedia_Article":       q["wikipedia_article"],
                "Pageviews_Monthly_Avg":   q["pageviews_monthly_avg"],
                "Pageviews_Source_Period": q["pageviews_source_period"],
                "Model_Response":          "",
                "Answer_Correctness":      "",
                "Hallucination_Reason":    "",
            })
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# XLSX SPEICHERN (Phase-2-Format)
# ─────────────────────────────────────────────────────────────────────────────

HEADERS = [
    "Question ID", "Category", "Question", "Correct Answer",
    "Prompt Variant", "Prompt Full", "Model",
    "source_url", "source_excerpt", "apa_citation",
    "Wikipedia_Article", "Pageviews_Monthly_Avg", "Pageviews_Source_Period",
    "Model_Response", "Answer_Correctness", "Hallucination_Reason",
]
COL_WIDTHS = [12, 16, 65, 55, 13, 70, 22, 48, 75, 60, 32, 18, 18, 55, 18, 50]

VARIANT_FILLS = {
    "neutral":    PatternFill("solid", start_color="EBF3FB"),
    "suggestive": PatternFill("solid", start_color="E2EFDA"),
    "abstention": PatternFill("solid", start_color="FFF2CC"),
}


def save_xlsx(rows: list[dict], path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Long-Tail Facts Phase 3"

    thin      = Side(style="thin", color="CCCCCC")
    border    = Border(top=thin, bottom=thin, left=thin, right=thin)
    hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    hdr_fill  = PatternFill("solid", start_color="1F4E79")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    norm_font = Font(name="Arial", size=10)

    for ci, h in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_fill
        c.alignment = hdr_align; c.border = border

    for ri, row in enumerate(rows, 2):
        variant = row.get("Prompt Variant", "")
        fill    = VARIANT_FILLS.get(variant, PatternFill())
        for ci, h in enumerate(HEADERS, 1):
            c = ws.cell(row=ri, column=ci, value=row.get(h, ""))
            c.font = norm_font; c.fill = fill
            c.alignment = lft_align; c.border = border

    for ci, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.freeze_panes = "A2"
    wb.save(path)
    print(f"  → XLSX gespeichert: {path}  ({len(rows)} Zeilen)")


# ─────────────────────────────────────────────────────────────────────────────
# PERSISTENZ
# ─────────────────────────────────────────────────────────────────────────────

def load_progress() -> dict:
    p = Path(PROGRESS_JSON)
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}


def save_progress(data: dict) -> None:
    Path(PROGRESS_JSON).write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ─────────────────────────────────────────────────────────────────────────────
# HAUPTPROGRAMM
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="LTF Phase 3 – Fragengenerierung mit Claude Sonnet"
    )
    parser.add_argument(
        "--triples", default=INPUT_TRIPLES,
        help=f"Eingabe-JSON mit Wikidata-Tripeln (Default: {INPUT_TRIPLES})"
    )
    parser.add_argument(
        "--start-from", type=int, default=0,
        help="Artikel-Index für Resume (Default: 0)"
    )
    parser.add_argument(
        "--fresh", action="store_true",
        help="Progress-Datei ignorieren und alle Fragen neu generieren"
    )
    args = parser.parse_args()

    print("=" * 65)
    print("Long-Tail Facts  Phase 3  –  Fragengenerierung mit Claude Sonnet")
    print("=" * 65)

    # Eingabedatei
    if not Path(args.triples).exists():
        print(f"\nFEHLER: {args.triples} nicht gefunden.")
        return

    with open(args.triples, encoding="utf-8") as f:
        raw_data: dict = json.load(f)

    # Nur Artikel mit Status 'found' und Tripeln
    all_data = {
        k: v for k, v in raw_data.items()
        if (v.get("status") == "found" or "triples" in v)
        and v.get("triples")
    }
    articles = list(all_data.keys())

    # Claude Client
    client = anthropic.Anthropic()   # liest ANTHROPIC_API_KEY aus Umgebung

    # Fortschritt laden
    if args.fresh:
        print("\n  [--fresh] Progress-Datei wird ignoriert, Neustart.")
        progress: dict = {}
    else:
        progress = load_progress()
        # WICHTIG: Prüfen ob der Progress von Claude oder von einem anderen
        # Modell stammt. Erkennungsmerkmal: Claude-Fragen haben typischerweise
        # einen längeren, kontextreichen Fragetext (> 120 Zeichen).
        old_questions = [v for v in progress.values() if v and v.get("question")]
        if old_questions:
            avg_len = sum(len(v["question"]) for v in old_questions) / len(old_questions)
            if avg_len < 100:
                print(
                    f"\n  ⚠  WARNUNG: Progress-Datei enthält {len(old_questions)} Fragen "
                    f"mit Ø {avg_len:.0f} Zeichen Länge.\n"
                    f"     Das deutet auf einen früheren gpt-oss-Lauf hin (keine Claude-Qualität).\n"
                    f"     Starte mit --fresh neu um alle Fragen mit Claude zu generieren:\n\n"
                    f"       python3 generate_longtail_phase3_claude.py --fresh\n"
                )
                antwort = input("  Trotzdem mit altem Progress fortfahren? (j/N): ").strip().lower()
                if antwort != "j":
                    print("  Abbruch. Starte mit --fresh für einen Neustart.")
                    return
                print()

    done: set[str] = set(progress.keys())
    questions: list[dict] = [v for v in progress.values() if v]

    print(f"\nArtikel gesamt:       {len(articles)}")
    print(f"Bereits generiert:    {len(questions)}")
    print(f"Claude-Modell:        {CLAUDE_MODEL}")
    print()

    for idx, article in enumerate(articles):
        if idx < args.start_from:
            continue
        if article in done:
            print(f"  [{idx+1:>3}] SKIP: {article[:55]}")
            continue

        print(f"\n  [{idx+1:>3}/{len(articles)}] {article}")

        info       = all_data[article]
        qid        = info.get("qid", "")
        label      = info.get("label", article)
        desc       = info.get("description", "")
        triples    = info.get("triples", [])
        avg_views  = info.get("avg_monthly_views", "")

        # Bestes Tripel auswählen
        best = select_best_triple(triples, idx=idx)
        if best is None:
            print(f"         → Kein geeignetes Tripel, übersprungen")
            progress[article] = None
            save_progress(progress)
            continue

        prop_id    = best["property_id"]
        prop_label = best["property_label"]
        value      = best["value_label"]
        print(f"         Tripel:  {prop_label} ({prop_id}) = {value}")

        # Wikipedia-Text holen
        wiki_text = get_wikipedia_text(article, max_chars=3000)
        time.sleep(SLEEP_WIKI)
        print(f"         Wiki:    {len(wiki_text)} Zeichen")

        # Kontext-Tripel formatieren (alles außer dem gewählten Tripel)
        context_triples = [t for t in triples if t["property_id"] != prop_id]
        triples_text = format_triples_for_prompt(context_triples)

        # Claude aufrufen
        prompt = build_prompt(
            entity_label  = label,
            entity_desc   = desc,
            target_prop   = prop_label,
            target_value  = value,
            triples_text  = triples_text,
            wiki_text     = wiki_text,
        )
        result = call_claude(client, prompt)
        time.sleep(SLEEP_CLAUDE)

        if result is None:
            print(f"         → Claude-Fehler, übersprungen")
            progress[article] = None
            save_progress(progress)
            continue

        # Source excerpt aufbauen
        excerpt = build_source_excerpt(
            article       = article,
            wiki_sentence = result["source_excerpt"],
            qid           = qid,
            prop_id       = prop_id,
            prop_label    = prop_label,
            value_label   = value,
            wiki_text     = wiki_text,
        )

        # Fragen-Record
        q_num = len(questions) + 1
        q_id  = f"{QUESTION_PREFIX}-Q{q_num:03d}"
        record: dict = {
            "question_id":             q_id,
            "category":                CATEGORY_LABEL,
            "question":                result["question"],
            "correct_answer":          result["correct_answer"],
            "wikipedia_article":       article,
            "pageviews_monthly_avg":   avg_views,
            "pageviews_source_period": PAGEVIEWS_PERIOD,
            "source_url":              WIKIPEDIA_BASE + article.replace(" ", "_"),
            "source_excerpt":          excerpt,
            "apa_citation":            build_apa(article, qid),
            "wikidata_qid":            qid,
            "property_id":             prop_id,
            "property_label":          prop_label,
        }

        questions.append(record)
        progress[article] = record

        print(f"         [{q_id}] {result['question'][:75]}")
        print(f"                  → {result['correct_answer'][:50]}")

        # Zwischenspeichern
        save_progress(progress)
        Path(OUTPUT_JSON).write_text(
            json.dumps(questions, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        # XLSX alle 15 Fragen
        if len(questions) % 15 == 0:
            all_rows = [row for q in questions for row in expand_question(q)]
            save_xlsx(all_rows, OUTPUT_XLSX)

        time.sleep(SLEEP_CLAUDE)

    # Finale Ausgabe
    print(f"\n[Finale Ausgabe …]")
    all_rows = [row for q in questions for row in expand_question(q)]
    save_xlsx(all_rows, OUTPUT_XLSX)
    Path(OUTPUT_JSON).write_text(
        json.dumps(questions, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    save_progress(progress)

    print()
    print("=" * 65)
    print("FERTIG")
    print(f"  Fragen generiert:    {len(questions)}")
    print(f"  Zeilen in XLSX:      {len(all_rows)}  ({len(questions)} × 9)")
    print(f"  XLSX:                {OUTPUT_XLSX}")
    print(f"  JSON (Backup):       {OUTPUT_JSON}")
    print()
    print("Nächster Schritt:")
    print(f"  python3 run_phase2_longtailfacts.py \\")
    print(f"      --input  {OUTPUT_XLSX} \\")
    print(f"      --output LongTailFacts_Phase3_ANSWERED.xlsx")
    print("=" * 65)


if __name__ == "__main__":
    main()
