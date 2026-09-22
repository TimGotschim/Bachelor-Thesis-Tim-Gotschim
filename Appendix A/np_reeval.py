#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
np_reeval.py
============
Bachelorarbeit - Tim Gotschim | WU Wien
Supervisor: Dr. Svitlana Vakulenko

Re-evaluation pipeline: LLM-Extraktion (Claude Haiku) + Python 3%-Schwellenwert

AENDERUNGEN in dieser Version:
  - Ueberspringt automatisch Zeilen, die bereits eine Klassifikation in
    New_Correctness haben (z.B. Qwen2.5 / Mistral aus einem frueheren Lauf),
    auch wenn diese Zeilen NICHT zusammenhaengend sind. Das war noetig, weil
    das neue Sonnet 5 Modell in denselben Katalog eingefuegt wurde und dessen
    480 Zeilen zwischen den bereits verarbeiteten Zeilen verstreut liegen.
  - Neuer Parameter --model-filter, um gezielt nur Zeilen eines bestimmten
    Modells zu verarbeiten (z.B. --model-filter "Sonnet 5").
  - Neuer Parameter --force, um bereits ausgefuellte Zeilen zu ueberschreiben.
  - Response-Truncation von 2500 auf 4000 Zeichen erhoeht (Sonnet 5 Antworten
    sind teils laenger als die bisher getesteten Modelle und die finale Zahl
    steht oft erst am Ende der Antwort).
  - Extraktionsprompt erweitert um explizite Regeln fuer LaTeX \boxed{},
    Markdown-Formatierung (fett, Code-Bloecke) und den Fall, dass die Antwort
    nur Programmcode ohne explizit genannte finale Zahl enthaelt (-> NULL).

Ausfuehrung:
  python3 np_reeval.py
  python3 np_reeval.py --model-filter "Sonnet 5"     # nur Sonnet 5 Zeilen
  python3 np_reeval.py --force                       # alles neu berechnen
  python3 np_reeval.py --start-row 500                # zusaetzlicher Filter
"""

import argparse, os, re, sys, time
from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import anthropic

# ═══════════════════════════════════════════════════════════════════════════
# KONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════
TOLERANCE_PCT   = 3.0
SAVE_EVERY      = 20
DELAY           = 0.2
MODEL           = "claude-haiku-4-5-20251001"
RESPONSE_CHARS  = 4000        # war 2500 -- Sonnet 5 Antworten koennen laenger sein

DIRECT_API_KEY = ""           # optional: Key direkt eintragen statt Env-Var

NEW_COLS = [
    "Extracted_Final",
    "Extracted_Final_Norm",
    "Ground_Truth_Norm",
    "New_Deviation_Pct",
    "New_Correctness",
]

# ═══════════════════════════════════════════════════════════════════════════
# PROMPT
# ═══════════════════════════════════════════════════════════════════════════
SYSTEM_PROMPT = """\
You are a precise numerical answer extractor for an academic evaluation pipeline.
Your ONLY task: find the FINAL numerical result the model states as its answer.

Rules:
1. Return ONLY the final answer number -- not intermediate calculation steps.
2. IGNORE unit conversion factors (e.g. 3.28084, 0.3861, 2.20462) that appear
   in the working but are NOT the final answer.
3. Remove units -- return only the bare number (e.g. "147.38" not "147.38 ft").
4. Use a decimal point (.) as separator, not a comma.
5. For percentage answers: return only the number ("38.10" not "38.10%").
6. The response may contain LaTeX or Markdown formatting. Extract the number
   from constructs such as $\\boxed{48.1}$, **48.1**, or `48.1` the same way
   as plain text -- ignore the markup, keep the number.
7. If the response contains programming code that would COMPUTE the answer
   (e.g. a Python function) but does not state the resulting number anywhere
   in plain text, treat this as no final answer given: return NULL. Do not
   attempt to execute or mentally evaluate the code yourself.
8. If the model refuses, says it cannot answer, or gives no final number: NULL

Respond with ONLY a single number (e.g. 82352369.9 or 1.35 or 38.10) OR NULL."""

USER_TEMPLATE = """\
Extract the FINAL numerical answer from this model response.
Ignore intermediate steps, conversion factors, and any LaTeX/Markdown markup.
Return ONLY the number or NULL.

Model response:
{response}"""

# ═══════════════════════════════════════════════════════════════════════════
# NORMALISIERUNG & KLASSIFIKATION
# ═══════════════════════════════════════════════════════════════════════════
def normalise(raw):
    if raw is None:
        return None
    s = str(raw).strip()
    if s.upper() in ("NULL", "", "NAN", "NONE", "N/A"):
        return None
    cleaned = re.sub(r"[^\d.,\-]", "", s)
    if not cleaned:
        return None
    if "." in cleaned and "," in cleaned:
        cleaned = cleaned.replace(",", "")
    elif "," in cleaned and "." not in cleaned:
        parts = cleaned.split(",")
        if len(parts) == 2 and len(parts[1]) == 3:
            cleaned = cleaned.replace(",", "")
        else:
            cleaned = cleaned.replace(",", ".", 1)
    cleaned = cleaned.strip(".,")
    try:
        return float(cleaned)
    except ValueError:
        return None


def deviation(ext, gt):
    return None if gt == 0 else abs(ext - gt) / abs(gt) * 100.0


def classify(ext, dev, tol):
    if ext is None or dev is None:
        return "Not_Attempted"
    return "Correct" if dev <= tol else "Hallucinated"


# ═══════════════════════════════════════════════════════════════════════════
# API-AUFRUF
# ═══════════════════════════════════════════════════════════════════════════
def extract_final(client, response_text, retries=3):
    user_msg = USER_TEMPLATE.format(response=response_text[:RESPONSE_CHARS])
    for attempt in range(1, retries + 1):
        try:
            msg = client.messages.create(
                model=MODEL,
                max_tokens=25,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_msg}],
            )
            return msg.content[0].text.strip()
        except Exception as exc:
            if attempt == retries:
                return f"ERROR: {exc}"
            time.sleep(5 * attempt)
    return "ERROR: max retries"


# ═══════════════════════════════════════════════════════════════════════════
# VERBINDUNGSTEST
# ═══════════════════════════════════════════════════════════════════════════
def test_connection(client):
    print(f"Teste Verbindung (Modell: {MODEL}) ...")
    try:
        r = client.messages.create(
            model=MODEL, max_tokens=10,
            messages=[{"role": "user", "content": "Return only the number 7."}],
        )
        ans = r.content[0].text.strip()
        print(f"  OK -- Antwort: '{ans}'\n")
        return True
    except Exception as e:
        err = str(e)
        print(f"  FEHLER: {err}\n")
        if "401" in err or "authentication" in err.lower():
            print("  Diagnose: ANTHROPIC_API_KEY falsch oder nicht gesetzt.")
        elif "overloaded" in err.lower():
            print("  Diagnose: API gerade ausgelastet -- kurz warten und erneut versuchen.")
        else:
            print(f"  Vollstaendige Fehlermeldung: {err}")
        return False


# ═══════════════════════════════════════════════════════════════════════════
# EXCEL-STYLING
# ═══════════════════════════════════════════════════════════════════════════
_thin  = Side(style="thin", color="CCCCCC")
BORDER = Border(top=_thin, bottom=_thin, left=_thin, right=_thin)

def _c(ws, row, col, value, bg=None, bold=False, fc="000000"):
    c = ws.cell(row=row, column=col, value=value)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    c.border    = BORDER
    if bg:
        c.fill  = PatternFill("solid", start_color=bg)
    c.font      = Font(name="Arial", size=10, bold=bold, color=fc)


# ═══════════════════════════════════════════════════════════════════════════
# HAUPTPROGRAMM
# ═══════════════════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(
        description="NP Re-evaluation: LLM-Extraktion + Python 3%-Schwellenwert")
    parser.add_argument("--input", type=Path,
        default=Path.home() / "Desktop" / "FINAL_Numerical_Precision.xlsx")
    parser.add_argument("--out",          type=Path,  default=None)
    parser.add_argument("--start-row",    type=int,   default=2,
        help="Zusaetzlicher Filter: ignoriert Zeilen vor diesem Index (Standard: 2 = alle)")
    parser.add_argument("--tolerance",    type=float, default=TOLERANCE_PCT)
    parser.add_argument("--delay",        type=float, default=DELAY)
    parser.add_argument("--no-test",      action="store_true")
    parser.add_argument("--model-filter", type=str,   default=None,
        help='Nur Zeilen mit diesem Wert in der Spalte "Model" verarbeiten, '
             'z.B. --model-filter "Sonnet 5"')
    parser.add_argument("--force", action="store_true",
        help="Bereits ausgefuellte Zeilen (New_Correctness gesetzt) trotzdem neu berechnen")
    args = parser.parse_args()

    out_path = args.out or args.input.with_stem(args.input.stem + "_reeval")

    # ── Client ──────────────────────────────────────────────────────────────
    api_key = DIRECT_API_KEY or os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        print("FEHLER: Kein API-Key gefunden.\n")
        print("Loesung A:  export ANTHROPIC_API_KEY='sk-ant-...'")
        print("Loesung B:  DIRECT_API_KEY in Zeile weiter oben eintragen.")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)

    if not args.no_test and not test_connection(client):
        print("Verbindungstest fehlgeschlagen. Abbruch.")
        sys.exit(1)

    # ── Excel laden ──────────────────────────────────────────────────────────
    if not args.input.exists():
        print(f"FEHLER: Datei nicht gefunden: {args.input}")
        sys.exit(1)

    print(f"Lade: {args.input}")
    wb = openpyxl.load_workbook(args.input)
    ws = wb.active
    hdrs = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}

    for req in ["Correct Answer", "Model_Response"]:
        if req not in hdrs:
            print(f"FEHLER: Spalte '{req}' nicht gefunden."); sys.exit(1)

    idx_gt   = hdrs["Correct Answer"]
    idx_resp = hdrs["Model_Response"]
    idx_mdl  = hdrs.get("Model")

    # ── Neue Spalten anlegen (oder vorhandene wiederverwenden) ────────────────
    if all(c in hdrs for c in NEW_COLS):
        nc0 = hdrs[NEW_COLS[0]]
        print(f"Resume-Modus: Spalten ab Spalte {get_column_letter(nc0)} vorhanden.")
    else:
        nc0 = ws.max_column + 1
        for i, name in enumerate(NEW_COLS):
            hdrs[name] = nc0 + i
            _c(ws, 1, nc0 + i, name, bg="1F3864", bold=True, fc="FFFFFF")
            ws.column_dimensions[get_column_letter(nc0 + i)].width = 22

    iEF  = hdrs["Extracted_Final"]
    iEFN = hdrs["Extracted_Final_Norm"]
    iGTN = hdrs["Ground_Truth_Norm"]
    iDEV = hdrs["New_Deviation_Pct"]
    iCLS = hdrs["New_Correctness"]

    BG = {"Correct": "E2EFDA", "Hallucinated": "FDECEA", "Not_Attempted": "FFF2CC"}

    # ── Zielzeilen bestimmen ───────────────────────────────────────────────────
    target_rows = []
    n_skipped_filled = 0
    n_skipped_model  = 0
    for row in range(max(args.start_row, 2), ws.max_row + 1):
        if ws.cell(row, idx_resp).value is None:
            continue
        if args.model_filter and idx_mdl:
            mdl_val = ws.cell(row, idx_mdl).value
            if str(mdl_val).strip() != args.model_filter:
                n_skipped_model += 1
                continue
        already_filled = ws.cell(row, iCLS).value not in (None, "")
        if already_filled and not args.force:
            n_skipped_filled += 1
            continue
        target_rows.append(row)

    total = len(target_rows)
    print(f"\nZeilen im Katalog gesamt : {ws.max_row - 1}")
    if args.model_filter:
        print(f"  Uebersprungen (falsches Modell)      : {n_skipped_model}")
    print(f"  Uebersprungen (bereits klassifiziert) : {n_skipped_filled}")
    print(f"  Zu verarbeiten                        : {total}")
    print(f"Toleranz: +-{args.tolerance}% | Ausgabe: {out_path.name}\n")

    if total == 0:
        print("Nichts zu tun -- alle relevanten Zeilen sind bereits ausgefuellt.")
        print("(Mit --force koennen bestehende Werte ueberschrieben werden.)")
        return

    # ── Hauptschleife ────────────────────────────────────────────────────────
    n_ok = n_h = n_na = n_err = 0

    for row in target_rows:
        resp_raw = ws.cell(row, idx_resp).value
        gt_raw   = ws.cell(row, idx_gt).value

        # 1. Ground Truth normalisieren (Python, kein API)
        gt_norm = normalise(str(gt_raw) if gt_raw is not None else "")

        # 2. Finale Zahl extrahieren (Claude Haiku)
        raw_ext = extract_final(client, str(resp_raw))
        time.sleep(args.delay)

        is_err = raw_ext.startswith("ERROR:")
        if is_err:
            n_err += 1
            _c(ws, row, iEF,  raw_ext, bg="FFCCCC")
            _c(ws, row, iEFN, "NULL",  bg="FFCCCC")
            _c(ws, row, iGTN, str(gt_norm) if gt_norm else "NULL", bg="EBF2F8")
            _c(ws, row, iDEV, "NULL",  bg="FFCCCC")
            _c(ws, row, iCLS, "Not_Attempted", bg="FFF2CC")
            print(f"  Zeile {row:>4}: API-FEHLER -- {raw_ext}")
            continue

        # 3. Extrahierten Wert normalisieren (Python)
        ext_norm = normalise(raw_ext)

        # 4. Abweichung berechnen (Python)
        dev = deviation(ext_norm, gt_norm) \
              if (ext_norm is not None and gt_norm is not None) else None

        # 5. Klassifizieren (Python)
        cls = classify(ext_norm, dev, args.tolerance)

        # 6. Schreiben
        _c(ws, row, iEF,  raw_ext,  bg="EBF2F8")
        _c(ws, row, iEFN, str(ext_norm) if ext_norm is not None else "NULL", bg="EBF2F8")
        _c(ws, row, iGTN, str(gt_norm)  if gt_norm  is not None else "NULL", bg="EBF2F8")
        _c(ws, row, iDEV, round(dev, 4) if dev is not None else "NULL", bg="EBF2F8")
        _c(ws, row, iCLS, cls, bg=BG.get(cls))

        dev_s = f"{dev:.2f}%" if dev is not None else "NULL"
        mdl   = ws.cell(row, hdrs.get("Model", 1)).value or ""
        qid   = ws.cell(row, hdrs.get("Question ID", 1)).value or row
        print(f"  Zeile {row:>4} | {str(qid)[:6]:6s} | {str(mdl)[:22]:22s} | "
              f"'{raw_ext[:12]:12s}' | dev={dev_s:>9s} | {cls}")

        if   cls == "Correct":       n_ok += 1
        elif cls == "Hallucinated":  n_h  += 1
        else:                        n_na += 1

        n_done = n_ok + n_h + n_na + n_err
        if n_done % SAVE_EVERY == 0:
            wb.save(out_path)
            print(f"  [Auto-save nach {n_done} Zeilen]")

    # ── Abschluss ────────────────────────────────────────────────────────────
    wb.save(out_path)
    n = max(n_ok + n_h + n_na, 1)
    print(f"\n{'='*65}")
    print(f"Gespeichert      : {out_path}")
    print(f"Correct  (<=3%)  : {n_ok:>4}  ({n_ok/n*100:.1f}%)")
    print(f"Hallucinated(>3%): {n_h:>4}  ({n_h/n*100:.1f}%)")
    print(f"Not_Attempted    : {n_na:>4}  ({n_na/n*100:.1f}%)")
    if n_err:
        print(f"API-Fehler       : {n_err:>4}  (erneut ausfuehren, bereits erledigte "
              f"Zeilen werden automatisch uebersprungen)")
    filt = f' --model-filter "{args.model_filter}"' if args.model_filter else ""
    print(f"\nErneut ausfuehren (verarbeitet nur noch offene Zeilen automatisch):")
    print(f"  python3 np_reeval.py{filt}")
    print(f"\nKosten ca.:  ${total * 0.001 * 0.0008:.3f} USD "
          f"(Haiku: ~$0.80/1M Tokens, ~1000 Tokens/Anfrage)")


if __name__ == "__main__":
    main()
