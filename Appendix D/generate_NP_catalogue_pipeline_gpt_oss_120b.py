#!/usr/bin/env python3
"""
generate_NP_catalogue_pipeline.py  (v2 — korrigiert)
======================================================
Bachelorarbeit – Tim Gotschim | WU Wien
Supervisor: Dr. Svitlana Vakulenko

Pipeline: Fragetext via WU Hub (gpt-oss:120b) — Antworten via Python.

KORREKTUREN gegenueber v1:
  1. Skalenkompatibilitaet: nur Entitaeten kombinieren, deren Werte
     ein max/min-Verhaeltnis von hoechstens 10 aufweisen.
  2. Typengruppen: Staedte mit Staedten, Berge mit Bergen etc.
  3. Prozent-Subjekt: mittlere Entitaet als Zaehler (nicht die groesste),
     so dass Ergebnisse typisch zwischen 20–80% liegen.
  4. Sinnvoll ueberpruefte Antworten: bei Prozent muss das Ergebnis
     zwischen 1 und 999% liegen, sonst wird die Kombination verworfen.

Ausfuehrung:
  export WU_HUB_URL="https://web.ollama-gpt-oss.ai.wu.ac.at"
  export WU_HUB_KEY="DEIN_KEY"
  python3 generate_NP_catalogue_pipeline.py --start-id 41 --target 200
  python3 generate_NP_catalogue_pipeline.py --resume   # fortsetzen
"""

from __future__ import annotations
import argparse, json, os, re, time
from itertools import combinations
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# KONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────

WU_HUB_URL   = os.environ.get("WU_HUB_URL", "https://web.ollama-gpt-oss.ai.wu.ac.at")
WU_HUB_KEY   = os.environ.get("WU_HUB_KEY", "")
MODEL_GEN    = "gpt-oss:120b"   # bessere Kreativitaet + Instruction-Following
MODEL_LIST   = ["Qwen2.5", "llama-3.1-8b", "mistral-7b-instruct"]
VARIANTS     = ["neutral", "suggestive", "abstention"]
MIN_VIEWS    = 50_000

# Maximales Wertverhaeltnis innerhalb einer Fragengruppe (Skalenfilter)
# Drei Staedte haben aehnliche Flaechen -> Verhaeltnis < 10
# Stadt + Ozean -> Verhaeltnis > 100.000 -> wird verworfen
MAX_RATIO    = 10.0

# Prozentantwort muss in diesem Bereich liegen (sonst Kombination verwerfen)
PCT_MIN, PCT_MAX = 5.0, 500.0

PROMPTS = {
    "neutral":    "Answer the following question.\n\nQuestion: {q}",
    "suggestive": ("The following question has a clear and specific answer. "
                   "Please respond with confidence and precision.\n\n{q}"),
    "abstention": ("Answer the following question. If you are not certain of "
                   "the answer, please say so explicitly rather than guessing."
                   "\n\nQuestion: {q}"),
}

USEFUL_UNITS = {
    'metre', 'kilometre', 'square kilometre', 'kilogram', 'degree celsius',
    'degree fahrenheit', 'kelvin', 'cubic kilometre', 'cubic metre',
    'kilometre per second', 'kilometre per hour', 'square metre', 'tonne', 'gram'
}

CONV_TABLE: dict[tuple, float] = {
    ('metre',            'foot'):          3.28084,
    ('metre',            'feet'):          3.28084,
    ('metre',            'inch'):          39.3701,
    ('kilometre',        'mile'):          0.621371,
    ('kilometre',        'foot'):          3280.84,
    ('kilometre',        'nautical mile'): 0.539957,
    ('square kilometre', 'square mile'):   0.386102,
    ('square kilometre', 'hectare'):       100.0,
    ('square metre',     'square foot'):   10.7639,
    ('square metre',     'acre'):          1/4046.86,
    ('kilogram',         'pound'):         2.20462,
    ('kilogram',         'tonne'):         0.001,
    ('kilometre per hour','mile per hour'): 0.621371,
    ('kilometre per hour','metre per second'): 1/3.6,
    ('kilometre per hour','knot'):         0.539957,
    ('cubic kilometre',  'cubic mile'):    0.239913,
    ('cubic metre',      'us gallon'):     264.172,
    ('cubic metre',      'litre'):         1000.0,
}

# (Quelleneinheit → Zieleinheit) fuer die Fragen
CONVERSION_PAIRS = [
    ('metre',            'foot'),
    ('metre',            'foot'),          # extra weight: haeufigste Konversion
    ('kilometre',        'mile'),
    ('square kilometre', 'square mile'),
    ('square kilometre', 'hectare'),
    ('kilogram',         'pound'),
    ('degree celsius',   'degree fahrenheit'),
    ('degree celsius',   'kelvin'),
    ('kilometre',        'foot'),
    ('kilometre per hour', 'mile per hour'),
    ('kilometre per hour', 'knot'),
]

OPERATIONS = ['average', 'percentage', 'ratio']   # average zuerst: stabiler

# ─────────────────────────────────────────────────────────────────────────────
# SPALTEN
# ─────────────────────────────────────────────────────────────────────────────

HEADERS = [
    'Question ID', 'Category', 'Question', 'Correct Answer', 'Computation Steps',
    'Numerical Type', 'Operation', 'Source Unit', 'Target Unit',
    'Entity1 Name', 'Entity1 Value', 'Entity1 Unit', 'Entity1 Wikipedia URL',
    'Entity2 Name', 'Entity2 Value', 'Entity2 Unit', 'Entity2 Wikipedia URL',
    'Entity3 Name', 'Entity3 Value', 'Entity3 Unit', 'Entity3 Wikipedia URL',
    'Prompt Variant', 'Prompt Full', 'Model',
    'Model_Response', 'Judge_Correctness', 'Judge_Justification',
]
COL_WIDTHS = {
    'Question ID': 14, 'Category': 22, 'Question': 80,
    'Correct Answer': 20, 'Computation Steps': 50,
    'Numerical Type': 22, 'Operation': 13, 'Source Unit': 18, 'Target Unit': 18,
    'Entity1 Name': 22, 'Entity1 Value': 14, 'Entity1 Unit': 18,
    'Entity1 Wikipedia URL': 42,
    'Entity2 Name': 22, 'Entity2 Value': 14, 'Entity2 Unit': 18,
    'Entity2 Wikipedia URL': 42,
    'Entity3 Name': 22, 'Entity3 Value': 14, 'Entity3 Unit': 18,
    'Entity3 Wikipedia URL': 42,
    'Prompt Variant': 12, 'Prompt Full': 80, 'Model': 22,
    'Model_Response': 50, 'Judge_Correctness': 18, 'Judge_Justification': 40,
}

# ─────────────────────────────────────────────────────────────────────────────
# ENTITAET-TYPENKLASSIFIKATION
# ─────────────────────────────────────────────────────────────────────────────

def classify_entity(instance_of: str) -> str:
    """
    Weist jede Entitaet einer Typengruppe zu.
    Nur Entitaeten derselben Gruppe werden kombiniert.
    """
    inst = (instance_of or '').lower()
    if any(k in inst for k in ('city','metropolis','megacity','capital city',
                                'municipality','urban','prefecture')):
        return 'city'
    if any(k in inst for k in ('mountain','volcano','stratovolcano','peak',
                                'summit','hill')):
        return 'mountain'
    if any(k in inst for k in ('country','republic','nation','kingdom',
                                'federation','state','territory')):
        return 'country'
    if any(k in inst for k in ('ocean','sea','lake','bay','gulf','strait',
                                'river','waterway')):
        return 'water_body'
    if any(k in inst for k in ('planet','moon','dwarf planet','gas giant',
                                'rocky planet','terrestrial planet')):
        return 'celestial'
    if any(k in inst for k in ('aircraft','airliner','airplane','jet')):
        return 'aircraft'
    if any(k in inst for k in ('chemical element','element')):
        return 'element'
    if any(k in inst for k in ('island','archipelago','atoll')):
        return 'island'
    if any(k in inst for k in ('desert','forest','rainforest','jungle',
                                'park','reserve','wilderness')):
        return 'natural_area'
    if any(k in inst for k in ('spacecraft','space station','rocket',
                                'satellite','probe')):
        return 'spacecraft'
    if any(k in inst for k in ('stadium','arena','ground','bowl')):
        return 'stadium'
    if any(k in inst for k in ('bridge','viaduct')):
        return 'bridge'
    if any(k in inst for k in ('building','skyscraper','tower','structure')):
        return 'building'
    return 'other'


# ─────────────────────────────────────────────────────────────────────────────
# ENTITAETEN LADEN UND GRUPPIEREN
# ─────────────────────────────────────────────────────────────────────────────

def load_entities(path: str) -> dict[str, dict]:
    with open(path, encoding='utf-8') as f:
        raw = json.load(f)

    usable = {}
    for article, info in raw.items():
        views = info.get('avg_monthly_pageviews')
        if not isinstance(views, (int, float)) or views < MIN_VIEWS:
            continue
        good = [t for t in info.get('numeric_triples', [])
                if t['unit_label'].lower() in USEFUL_UNITS
                and t['property_label'] not in ('population',)
                and t['amount'] > 0]
        if not good:
            continue
        seen_p, deduped = set(), []
        for t in good:
            if t['property_label'] not in seen_p:
                seen_p.add(t['property_label'])
                deduped.append(t)
        inst = (info.get('instance_of') or '').split('/')[0].strip()
        usable[article] = {
            'label':     info.get('label', article),
            'qid':       info.get('qid', ''),
            'pageviews': views,
            'triples':   deduped,
            'type':      classify_entity(inst),
            'wiki_url':  f"https://en.wikipedia.org/wiki/{article.replace(' ', '_')}",
        }
    return usable


def group_by_type_and_unit(entities: dict) -> dict[tuple, list]:
    """
    Gruppiert Entitaeten nach (entity_type, property_label, unit_label).
    Nur Gruppen mit >= 3 Entitaeten werden beruecksichtigt.
    """
    groups: dict[tuple, list] = {}
    for article, info in entities.items():
        for t in info['triples']:
            key = (info['type'], t['property_label'], t['unit_label'].lower())
            groups.setdefault(key, []).append({
                'article': article,
                'info':    info,
                'triple':  t,
            })
    return {k: v for k, v in groups.items() if len(v) >= 3}


# ─────────────────────────────────────────────────────────────────────────────
# KONVERSION UND MATHEMATIK
# ─────────────────────────────────────────────────────────────────────────────

def convert(value: float, src: str, tgt: str) -> float | None:
    s, t = src.lower().strip(), tgt.lower().strip()
    if s == t: return value
    if (s, t) in CONV_TABLE: return value * CONV_TABLE[(s, t)]
    if s == 'degree celsius':
        if 'fahrenheit' in t: return value * 9/5 + 32
        if 'kelvin' in t:     return value + 273.15
    if s == 'degree fahrenheit':
        if 'celsius' in t:    return (value - 32) * 5/9
        if 'kelvin' in t:     return (value - 32) * 5/9 + 273.15
    if s == 'kelvin':
        if 'celsius' in t:    return value - 273.15
        if 'fahrenheit' in t: return (value - 273.15) * 9/5 + 32
    return None


def choose_subject_for_percentage(vals: list[float]) -> int:
    """
    Waehlt die mittlere Entitaet als Zaehler fuer Prozentrechnung.
    Das produziert typisch sinnvolle Werte (20–80%), statt absurde
    Zahlen wie 1.700.000% wenn eine Entitaet viel groesser ist.
    """
    indexed = sorted(enumerate(vals), key=lambda x: x[1])
    return indexed[len(indexed)//2][0]   # mittlerer Rang


def compute_answer(vals: list[float], operation: str,
                   subj_idx: int = 0) -> tuple[float, str]:
    v = vals
    if operation == 'average':
        result = sum(v) / len(v)
        steps  = (f"({' + '.join(f'{x:,.3f}' for x in v)}) / {len(v)} "
                  f"= {result:,.4f}")
    elif operation == 'percentage':
        sv     = v[subj_idx]
        rest   = sum(x for i, x in enumerate(v) if i != subj_idx)
        result = sv / rest * 100
        steps  = f"{sv:,.4f} / {rest:,.4f} × 100 = {result:,.4f}%"
    else:   # ratio
        big, small = max(v), min(v)
        result = big / small
        steps  = f"{big:,.4f} / {small:,.4f} = {result:,.4f}"
    return result, steps


def is_result_sane(result: float, operation: str) -> bool:
    """Prueft ob das Ergebnis methodisch sinnvoll ist."""
    if operation == 'percentage':
        return PCT_MIN <= result <= PCT_MAX
    if operation == 'ratio':
        return 1.0 <= result <= 1000.0
    return True   # averages sind immer sinnvoll


def format_answer(result: float, operation: str, tgt_unit: str) -> str:
    if operation == 'percentage':
        return f"{result:,.2f}%"
    if operation == 'ratio':
        return f"{result:,.2f}"
    abbrev = {
        'square mile': 'sq mi', 'square miles': 'sq mi',
        'foot': 'ft', 'feet': 'ft',
        'mile': 'mi', 'miles': 'mi',
        'pound': 'lbs', 'metric tonne': 't', 'tonne': 't',
        'degree fahrenheit': '°F', 'kelvin': 'K',
        'mile per hour': 'mph', 'knot': 'kn', 'knots': 'kn',
        'hectare': 'ha', 'hectares': 'ha',
        'us gallon': 'gal', 'litre': 'L',
    }.get(tgt_unit.lower(), tgt_unit)
    if abs(result) >= 1000:
        return f"{result:,.1f} {abbrev}"
    return f"{result:.2f} {abbrev}"


# ─────────────────────────────────────────────────────────────────────────────
# WU HUB API — NUR FRAGETEXT
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are writing questions for a hallucination study on large language models.
Each question must trigger hallucinations by combining:
  1. Entity names only — NO numerical values in the question text.
  2. Multi-step calculation: percentage, average, or ratio.
  3. Unit conversion integrated into the calculation.

Strict rules:
- Professional scenario (scientist, engineer, geographer, etc.)
- ONE single numerical answer — no sub-questions.
- Do NOT mention the actual measured values.
- End percentage/ratio questions with: "Round your answer to two decimal places."
- End average questions with: "Round your answer to one decimal place."
- Output ONLY the question text. Nothing else."""


def generate_question_text(articles: list[str], prop: str,
                            src_unit: str, tgt_unit: str,
                            operation: str, entity_type: str) -> str | None:
    type_context = {
        'city':        'major world cities',
        'mountain':    'famous mountain peaks',
        'country':     'large countries',
        'celestial':   'planets or moons in our solar system',
        'aircraft':    'commercial aircraft',
        'element':     'chemical elements',
        'natural_area':'natural regions or rainforests',
        'water_body':  'rivers, lakes or oceans',
    }.get(entity_type, 'well-known geographic or scientific entities')

    prompt = (
        f"Entities ({type_context}): {', '.join(articles)}\n"
        f"Property: {prop}\n"
        f"Convert: {src_unit} → {tgt_unit}\n"
        f"Operation: {operation}\n\n"
        f"Write one question. The model must convert {src_unit} to {tgt_unit} "
        f"as part of computing the {operation}. "
        f"Do NOT state the actual {src_unit} values in the question."
    )

    url  = f"{WU_HUB_URL.rstrip('/')}/api/chat/completions"
    hdrs = {"Content-Type": "application/json",
            "Authorization": f"Bearer {WU_HUB_KEY}" if WU_HUB_KEY else ""}
    body = {
        "model": MODEL_GEN,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": prompt},
        ],
        "temperature": 0.4,
        "max_tokens":  280,
    }

    for attempt in range(1, 4):
        try:
            r = requests.post(url, headers=hdrs, json=body, timeout=120)
            if r.status_code == 200:
                text = r.json()["choices"][0]["message"]["content"].strip()
                text = re.sub(r'^["\']|["\']$', '', text).strip()
                text = re.sub(r'^(Question:|Q:)\s*', '', text, flags=re.I).strip()
                return text if len(text) > 30 else None
            print(f"    [HTTP {r.status_code}] {r.text[:80]}")
        except Exception as e:
            print(f"    [Versuch {attempt}] {type(e).__name__}: {str(e)[:60]}")
        time.sleep(3 * attempt)
    return None


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────────────────────────────────────

C_NAVY = "1F3864"; C_ALT = "F7F9FB"

HDR_COLORS = {
    'Question': 'A4C2F4', 'Correct Answer': '93C47D',
    'Computation Steps': 'B6D7A8',
    'Entity1 Name': 'FFE599', 'Entity1 Value': 'FFE599',
    'Entity1 Unit': 'FFE599', 'Entity1 Wikipedia URL': 'FFD966',
    'Entity2 Name': 'F9CB9C', 'Entity2 Value': 'F9CB9C',
    'Entity2 Unit': 'F9CB9C', 'Entity2 Wikipedia URL': 'F6B26B',
    'Entity3 Name': 'EA9999', 'Entity3 Value': 'EA9999',
    'Entity3 Unit': 'EA9999', 'Entity3 Wikipedia URL': 'E06666',
    'Prompt Full': 'D5E8D4', 'Model_Response': 'EAF5EA',
}

def _fill(c): return PatternFill("solid", start_color=c)
def _font(bold=False, size=9, white=False):
    return Font(name="Arial", bold=bold, size=size,
                color="FFFFFF" if white else "000000")
def _align():  return Alignment(vertical="top", wrap_text=True, horizontal="left")
def _border():
    t = Side(style="thin", color="CCCCCC")
    return Border(top=t, bottom=t, left=t, right=t)


def init_workbook(out_path: Path):
    wb = Workbook(); ws = wb.active; ws.title = "NP Master Catalogue"
    ws.append(HEADERS)
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, c)
        bg   = HDR_COLORS.get(h, 'C9DAF8')
        cell.fill = _fill(bg); cell.font = _font(bold=True, size=10)
        cell.alignment = _align(); cell.border = _border()
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTHS.get(h, 14)
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    wb.save(out_path)
    return wb, ws


def style_data_rows(ws, from_row: int, to_row: int):
    for r in range(from_row, to_row + 1):
        for c in range(1, len(HEADERS) + 1):
            cell = ws.cell(r, c)
            cell.font = _font(size=9)
            cell.alignment = _align()
            cell.border = _border()
            if r % 2 == 0:
                cell.fill = _fill(C_ALT)


def append_question(ws, q: dict):
    """Schreibt 9 Zeilen (3 Modelle × 3 Varianten) fuer eine Frage."""
    em = q['entities']
    for model in MODEL_LIST:
        for variant in VARIANTS:
            ws.append([
                q['qid'],
                "Numerical Precision – Master Catalogue (Phase 3)",
                q['question'],
                q['answer'],
                q['steps'],
                q['num_type'],
                q['operation'],
                q['src_unit'],
                q['tgt_unit'],
                em[0]['name'], em[0]['value'], em[0]['unit'], em[0]['wiki_url'],
                em[1]['name'], em[1]['value'], em[1]['unit'], em[1]['wiki_url'],
                em[2]['name'], em[2]['value'], em[2]['unit'], em[2]['wiki_url'],
                variant,
                PROMPTS[variant].format(q=q['question']),
                model,
                "", "", "",
            ])


# ─────────────────────────────────────────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────────────────────────────────────────

def run(entities: dict, start_id: int, target_id: int,
        out_path: Path, resume: bool):

    # Workbook laden oder anlegen
    if resume and out_path.exists():
        wb = load_workbook(out_path)
        ws = wb.active
        done_ids = set()
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if v: done_ids.add(v)
        current_id = start_id + len(done_ids) // 9
        print(f"  Resume: {len(done_ids)//9} Fragen bereits vorhanden.")
    else:
        wb, ws = init_workbook(out_path)
        done_ids = set()
        current_id = start_id

    groups = group_by_type_and_unit(entities)
    print(f"  Gruppen (Typ+Property+Einheit): {len(groups)}")

    generated = 0
    max_new = target_id - current_id + 1

    # ─────────────────────────────────────────────────────────────────────────
    # ROUND-ROBIN-ROTATION (Korrektur Diversitaet)
    #
    # Statt Gruppen nach Groesse abzuarbeiten (was zu 160x derselben Property
    # fuehrte), bauen wir pro Gruppe einen Vorrat skalengefilterter Kombinationen
    # und nehmen dann reihum aus JEDER Gruppe eine Frage. So entstehen
    # abwechslungsreiche Fragen ueber alle Property-Typen und Domaenen.
    # ─────────────────────────────────────────────────────────────────────────

    # 1. Pro Gruppe: gueltige Kombinationen vorbereiten (skalengefiltert)
    group_state: dict = {}
    for (etype, prop, src_unit), ent_list in groups.items():
        tgt_units = list(dict.fromkeys(
            tgt for src, tgt in CONVERSION_PAIRS if src == src_unit
        ))
        if not tgt_units:
            continue
        tgt_unit = tgt_units[0]

        valid_combos = []
        for combo in combinations(range(len(ent_list)), 3):
            items = [ent_list[i] for i in combo]
            vals_conv, entity_meta, ok = [], [], True
            for it in items:
                t  = it['triple']
                cv = convert(t['amount'], t['unit_label'], tgt_unit)
                if cv is None: ok = False; break
                vals_conv.append(cv)
                entity_meta.append({
                    'name':     it['article'],
                    'value':    t['amount'],
                    'unit':     t['unit_label'],
                    'wiki_url': it['info']['wiki_url'],
                })
            if not ok: continue
            # Skalenfilter
            if max(vals_conv) / min(vals_conv) > MAX_RATIO:
                continue
            valid_combos.append((vals_conv, entity_meta,
                                 [it['article'] for it in items]))

        if valid_combos:
            group_state[(etype, prop, src_unit, tgt_unit)] = {
                'combos':    valid_combos,
                'combo_idx': 0,
                'op_idx':    0,   # rotiert Operation pro Besuch
            }

    group_keys = sorted(group_state.keys())
    print(f"  Verwendbare Gruppen (mit gueltigen Kombos): {len(group_keys)}")

    # 2. Round-Robin: reihum eine Frage pro Gruppe, bis Ziel erreicht
    active = list(group_keys)
    while generated < max_new and active:
        for key in list(active):
            if generated >= max_new: break

            etype, prop, src_unit, tgt_unit = key
            st = group_state[key]

            if st['combo_idx'] >= len(st['combos']):
                active.remove(key)   # Gruppe erschoepft
                continue

            vals_conv, entity_meta, articles = st['combos'][st['combo_idx']]
            st['combo_idx'] += 1

            # Operation rotiert pro Besuch (average → percentage → ratio → …)
            operation = OPERATIONS[st['op_idx'] % len(OPERATIONS)]
            st['op_idx'] += 1

            qid = f"NP-M-Q{current_id:03d}"
            if qid in done_ids:
                current_id += 1
                continue

            # ── ANTWORT BERECHNEN (Python, kein LLM) ──────────────────────
            subj_idx = (choose_subject_for_percentage(vals_conv)
                        if operation == 'percentage' else 0)
            try:
                result, steps = compute_answer(vals_conv, operation, subj_idx)
            except ZeroDivisionError:
                continue

            # ── SINNVOLLKEITSFILTER ───────────────────────────────────────
            if not is_result_sane(result, operation):
                continue

            # ── FRAGETEXT VOM WU HUB (Cogito v2.1) ────────────────────────
            print(f"\n  [{qid}] {etype:12s} | {operation:10s} | "
                  f"{prop} | {src_unit} → {tgt_unit}")
            print(f"    Entities: {', '.join(articles)}")
            print(f"    Werte: {[round(v,2) for v in vals_conv]} "
                  f"(Ratio: {max(vals_conv)/min(vals_conv):.1f}x)")

            q_text = generate_question_text(
                articles, prop, src_unit, tgt_unit, operation, etype
            )
            if not q_text:
                print("    [SKIP] Kein Fragetext erhalten")
                continue

            formatted = format_answer(result, operation, tgt_unit)
            print(f"    → Antwort: {formatted}  |  {steps[:55]}")

            q_data = {
                'qid':       qid,
                'question':  q_text,
                'answer':    formatted,
                'steps':     steps,
                'num_type':  f"{operation} + unit_conversion",
                'operation': operation,
                'src_unit':  src_unit,
                'tgt_unit':  tgt_unit,
                'entities':  entity_meta,
            }

            # ── 9 ZEILEN SCHREIBEN ────────────────────────────────────────
            first_new_row = ws.max_row + 1
            append_question(ws, q_data)
            style_data_rows(ws, first_new_row, ws.max_row)

            generated  += 1
            current_id += 1

            if generated % 10 == 0:
                wb.save(out_path)
                print(f"  [Auto-save: {generated} Fragen → {out_path}]")

            time.sleep(0.4)

    wb.save(out_path)
    print(f"\n{'='*60}")
    print(f"Neue Fragen generiert  : {generated}")
    print(f"Gesamt Datenzeilen     : {ws.max_row - 1}")
    print(f"Ausgabe                : {out_path}")
    print("=" * 60)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(
        description="NP Pipeline v2: Fragetext via WU Hub, Antworten via Python"
    )
    p.add_argument("--entities", default="np_entities_with_pageviews.json")
    p.add_argument("--out",      default="NP_Master_Catalogue_Pipeline.xlsx")
    p.add_argument("--start-id", type=int, default=41)
    p.add_argument("--target",   type=int, default=200)
    p.add_argument("--resume",   action="store_true")
    args = p.parse_args()

    print("=" * 60)
    print(f"NP Catalogue Pipeline v2")
    print(f"  Modell (Fragetext) : {MODEL_GEN}")
    print(f"  Antworten          : Python-Arithmetik (kein LLM)")
    print(f"  Skalenfilter       : max/min ≤ {MAX_RATIO}x")
    print(f"  Prozent-Bereich    : {PCT_MIN}% – {PCT_MAX}%")
    print(f"  Min. Pageviews     : {MIN_VIEWS:,}/Monat")
    print(f"  Ziel-IDs           : Q{args.start_id:03d} – Q{args.target:03d}")
    if not WU_HUB_KEY:
        print("\n  [!] WU_HUB_KEY nicht gesetzt — bitte setzen:")
        print("      export WU_HUB_KEY='DEIN_KEY'")
    print("=" * 60)

    entities = load_entities(args.entities)
    print(f"\n  Usable entities (≥{MIN_VIEWS:,} views): {len(entities)}")
    groups = group_by_type_and_unit(entities)
    print(f"  Typgruppen (≥3 Entitaeten): {len(groups)}")
    from collections import Counter
    type_summary = Counter(k[0] for k in groups)
    for etype, cnt in type_summary.most_common():
        print(f"    {etype:<15} {cnt} Gruppen")

    run(entities, args.start_id, args.target,
        Path(args.out), args.resume)


if __name__ == "__main__":
    main()
