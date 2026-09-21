#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
judge_confusable_entities.py
============================
Dual-LLM-as-a-Judge evaluation for the Confusable Entities catalogue.
Sends every model response to gpt-oss:120b (WU Hub, OpenAI-compatible endpoint)
and classifies it as Correct / Hallucinated / Not_Attempted.

CE-specific rule (Section 3.5.3):
  If the model answers about the prominent TOP entity instead of the less-known
  SHADOW entity (Entity Overshadowing), the response is HALLUCINATED — even if
  the information would be factually correct for the Top entity.

Column layout (input file after clean_catalogue.py)
-----------------------------------------------------
  A  Question ID       C  Question           D  Correct Answer
  E  Surface Form      H  Top Entity         I  Top Entity Description
  K  Shadow Entity     L  Shadow Entity Desc  R  source_excerpt
  N  Prompt Variant    O  Prompt Full         P  Model
  T  Model_Response    U  Leakage_Reason (read-only, not modified)

Judge output columns (already present, to be filled)
------------------------------------------------------
  V  Judge1_Correctness    W  Judge1_Confidence   X  Judge1_Justification
  Y  Judge2_Correctness    Z  Judge2_Confidence   AA Judge2_Justification
  AB Final_Correctness     AC Needs_Manual_Review  AD Final_Justification

Usage
-----
  python3 judge_confusable_entities.py
  python3 judge_confusable_entities.py \\
      --input  ConfusableEntities_Master_v2_CLEAN.xlsx \\
      --output ConfusableEntities_Master_v2_EVALUATED.xlsx
  python3 judge_confusable_entities.py --start-row 2      # default: start from row 2
  python3 judge_confusable_entities.py --overwrite
  python3 judge_confusable_entities.py --only-model Qwen2.5
  python3 judge_confusable_entities.py --only-variant neutral

Environment variables
---------------------
  WU_HUB_URL   Base URL of the WU Model Hub (without /v1)
                Default: https://web.ollama-gpt-oss.ai.wu.ac.at/api
  WU_HUB_KEY   API key if required (default: "dummy")
"""

from __future__ import annotations

import argparse
import json
import os
import re
import time
from pathlib import Path

from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
WU_HUB_URL  = os.environ.get("WU_HUB_URL", "https://web.ollama-gpt-oss.ai.wu.ac.at")
WU_HUB_KEY  = os.environ.get("WU_HUB_KEY", "dummy")
JUDGE_MODEL = "gpt-oss:120b"

TEMPERATURE           = 0.0
MAX_TOKENS            = 600
SAVE_EVERY            = 20
PAUSE_BETWEEN_CALLS   = 0.5   # seconds between API calls

# Column indices (1-based for openpyxl) — matches output of clean_catalogue.py
COL = {
    "qid":            1,   # A
    "question":       3,   # C
    "correct_answer": 4,   # D
    "surface_form":   5,   # E
    "top_entity":     8,   # H
    "top_desc":       9,   # I
    "shadow_entity":  11,  # K
    "shadow_desc":    12,  # L
    "prompt_variant": 14,  # N
    "source_excerpt": 18,  # R
    "model_response": 20,  # T
    # Judge output columns
    "j1_class":       22,  # V
    "j1_conf":        23,  # W
    "j1_just":        24,  # X
    "j2_class":       25,  # Y
    "j2_conf":        26,  # Z
    "j2_just":        27,  # AA
    "final_class":    28,  # AB
    "needs_review":   29,  # AC
    "final_just":     30,  # AD
}
MODEL_COL_IDX = 16  # P

# ---------------------------------------------------------------------------
# Judge prompts (English throughout)
# ---------------------------------------------------------------------------
SYSTEM_JUDGE1 = """\
You are a strict, impartial evaluator for an academic LLM hallucination study
(Bachelor's thesis, WU Vienna).

Your task: classify a language model's response using ONLY the verified ground-truth
information provided below. Your own world knowledge is FORBIDDEN.

Category: Confusable Entities (Entity Overshadowing)
Phenomenon: an LLM confuses a lesser-known SHADOW entity with a more famous
TOP entity that shares the same surface form (Provatorova, Vakulenko et al., 2021).

Classification rules:
  Correct       — The response clearly refers to the SHADOW entity and matches
                  the verified Correct Answer.
  Hallucinated  — (a) The model answers about the TOP entity instead of the SHADOW
                  entity [entity confusion], OR (b) the model states incorrect facts
                  about the SHADOW entity.
                  IMPORTANT: Even if the response would be factually accurate for the
                  TOP entity, an entity confusion still counts as HALLUCINATED.
  Not_Attempted — The model explicitly declines to answer, states it is uncertain,
                  or provides no concrete factual claim.

Reply with ONLY a JSON object (no code fences, no preamble):
{"classification": "Correct|Hallucinated|Not_Attempted",
 "confidence": <1-5>, "justification": "<max. 2 sentences>"}"""


def make_prompt_judge1(question, correct_answer, surface_form,
                        top_entity, top_desc, shadow_entity, shadow_desc,
                        source_excerpt, model_response):
    return f"""\
Surface form (shared name): {surface_form}
TOP entity (prominent, NOT the target): {top_entity} — {top_desc}
SHADOW entity (less-known, IS the target): {shadow_entity} — {shadow_desc}

Question: {question}
Verified correct answer: {correct_answer}
Source excerpt (Wikipedia evidence): {source_excerpt}

Model response:
{model_response}

Evaluate the model response. Pay particular attention to whether the model answers
about {top_entity} instead of {shadow_entity} (entity confusion = Hallucinated)."""


SYSTEM_JUDGE2 = """\
You are a strict, impartial evaluator for an academic LLM hallucination study
(Bachelor's thesis, WU Vienna).

Your task: classify a language model's response using your own world knowledge.
You do NOT receive the verified correct answer.

Category: Confusable Entities (Entity Overshadowing)
Phenomenon: an LLM confuses a lesser-known SHADOW entity with a more famous
TOP entity that shares the same surface form (Provatorova, Vakulenko et al., 2021).

Classification rules:
  Correct       — The response clearly refers to the SHADOW entity and is
                  factually correct according to your world knowledge.
  Hallucinated  — (a) The model answers about the TOP entity instead of the SHADOW
                  entity [entity confusion], OR (b) the model states incorrect facts
                  about the SHADOW entity.
                  IMPORTANT: Even if the response would be factually accurate for the
                  TOP entity, an entity confusion still counts as HALLUCINATED.
  Not_Attempted — The model explicitly declines to answer, states it is uncertain,
                  or provides no concrete factual claim.
  Special case: If you yourself are unfamiliar with the SHADOW entity and the model
  response contains a concrete factual claim, classify as Hallucinated (conservative
  default for unknown entities — the model is more likely confabulating than correct).

Reply with ONLY a JSON object (no code fences, no preamble):
{"classification": "Correct|Hallucinated|Not_Attempted",
 "confidence": <1-5>, "justification": "<max. 2 sentences>"}"""


def make_prompt_judge2(question, surface_form, top_entity, shadow_entity, model_response):
    return f"""\
Surface form (shared name): {surface_form}
TOP entity (prominent, NOT the target): {top_entity}
SHADOW entity (less-known, IS the target): {shadow_entity}

Question: {question}

Model response:
{model_response}

Evaluate the model response using your world knowledge. Pay particular attention
to whether the model answers about {top_entity} instead of {shadow_entity}
(entity confusion = Hallucinated)."""


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
def nz(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def clean_latex(text: str) -> str:
    """Strip LaTeX notation that local models sometimes produce."""
    text = re.sub(r"\$\$?[^$]*\$\$?", "", text)
    text = re.sub(r"\\[a-zA-Z]+\{[^}]*\}", "", text)
    return text.strip()


def safe_parse_json(raw: str) -> dict | None:
    raw = re.sub(r"^```(json)?", "", raw.strip()).strip()
    raw = re.sub(r"```$", "", raw).strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except json.JSONDecodeError:
                pass
    return None


VALID_LABELS = {"Correct", "Hallucinated", "Not_Attempted"}


def call_judge(client: OpenAI, system: str, user: str) -> dict:
    """
    Calls the judge with an assistant-turn prefill to force JSON output
    and suppress chain-of-thought preambles.
    """
    try:
        resp = client.chat.completions.create(
            model=JUDGE_MODEL,
            temperature=TEMPERATURE,
            max_tokens=MAX_TOKENS,
            messages=[
                {"role": "system",    "content": system},
                {"role": "user",      "content": user},
                {"role": "assistant", "content": '{"classification":'},  # prefill
            ],
        )
        raw = '{"classification":' + resp.choices[0].message.content.strip()
        parsed = safe_parse_json(raw)
        if parsed and parsed.get("classification") in VALID_LABELS:
            return {
                "classification": parsed["classification"],
                "confidence":     int(parsed.get("confidence", 3)),
                "justification":  str(parsed.get("justification", "")).strip(),
            }
        return {"classification": "ERROR_PARSE", "confidence": 0,
                "justification": f"JSON parse failed: {raw[:200]}"}
    except Exception as exc:
        return {"classification": "ERROR_API", "confidence": 0,
                "justification": str(exc)[:200]}


def resolve(j1: str, j2: str) -> tuple[str, str]:
    """
    Conflict resolution: only identical verdicts produce a final label.
    Any disagreement → Conflicting, manual review required.
    """
    if j1 == j2:
        return j1, "NO"
    return "Conflicting", "YES"


# ---------------------------------------------------------------------------
# Ensure judge column headers exist
# ---------------------------------------------------------------------------
NEW_HEADERS = {
    COL["j1_class"]:    "Judge1_Correctness",
    COL["j1_conf"]:     "Judge1_Confidence",
    COL["j1_just"]:     "Judge1_Justification",
    COL["j2_class"]:    "Judge2_Correctness",
    COL["j2_conf"]:     "Judge2_Confidence",
    COL["j2_just"]:     "Judge2_Justification",
    COL["final_class"]: "Final_Correctness",
    COL["needs_review"]:"Needs_Manual_Review",
    COL["final_just"]:  "Final_Justification",
}

def ensure_headers(ws) -> None:
    for col_idx, name in NEW_HEADERS.items():
        cell = ws.cell(row=1, column=col_idx)
        if not cell.value:
            cell.value = name


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(
        description="Dual-LLM-as-a-Judge for Confusable Entities")
    parser.add_argument("--input",  type=Path,
                        default=Path("ConfusableEntities_Master_v2_CLEAN.xlsx"))
    parser.add_argument("--output", type=Path,
                        default=Path("ConfusableEntities_Master_v2_EVALUATED.xlsx"))
    parser.add_argument("--sheet",  type=str, default="Confusable Entities")
    parser.add_argument("--start-row", type=int, default=2,
                        help="Starting row (1-based; 2 = first data row)")
    parser.add_argument("--overwrite", action="store_true",
                        help="Re-evaluate rows that already have a judge result")
    parser.add_argument("--retry-errors", action="store_true",
                        help="Re-evaluate only rows whose Final_Correctness is ERROR "
                             "(API/parse failures), leaving all other completed rows untouched")
    parser.add_argument("--only-model",   type=str, default=None, metavar="MODEL")
    parser.add_argument("--only-variant", type=str, default=None, metavar="VARIANT")
    args = parser.parse_args()

    if not args.input.exists():
        raise FileNotFoundError(f"File not found: {args.input}")

    print("=" * 65)
    print("Confusable Entities — Dual-LLM-as-a-Judge")
    print(f"  Input:      {args.input}")
    print(f"  Output:     {args.output}")
    print(f"  Judge:      {JUDGE_MODEL}  @  {WU_HUB_URL}")
    print(f"  Start row:  {args.start_row}  (all questions, incl. K3-Q1–Q40)")
    if args.only_model:   print(f"  Filter model:   {args.only_model}")
    if args.only_variant: print(f"  Filter variant: {args.only_variant}")
    print("=" * 65)

    client = OpenAI(base_url=f"{WU_HUB_URL}/v1", api_key=WU_HUB_KEY)

    wb = load_workbook(args.input)
    ws_name = args.sheet if args.sheet in wb.sheetnames else wb.active.title
    ws = wb[ws_name]
    ensure_headers(ws)
    print(f"\n  Sheet: '{ws.title}' | Total data rows: {ws.max_row - 1}\n")

    processed = skipped_existing = skipped_filter = errors = 0

    for row in range(max(2, args.start_row), ws.max_row + 1):
        qid     = nz(ws.cell(row=row, column=COL["qid"]).value)
        model   = nz(ws.cell(row=row, column=MODEL_COL_IDX).value)
        variant = nz(ws.cell(row=row, column=COL["prompt_variant"]).value)

        # User-defined filters
        if args.only_model and model != args.only_model:
            skipped_filter += 1
            continue
        if args.only_variant and variant != args.only_variant:
            skipped_filter += 1
            continue

        # Skip already-evaluated rows unless --overwrite (or --retry-errors on error rows)
        existing_final = nz(ws.cell(row=row, column=COL["final_class"]).value)
        existing_j1    = nz(ws.cell(row=row, column=COL["j1_class"]).value)
        if existing_j1:
            if args.overwrite:
                pass  # re-evaluate everything
            elif args.retry_errors and existing_final == "ERROR":
                pass  # re-evaluate only this error row
            else:
                skipped_existing += 1
                continue

        # Read input data
        question       = nz(ws.cell(row=row, column=COL["question"]).value)
        correct_answer = nz(ws.cell(row=row, column=COL["correct_answer"]).value)
        surface_form   = nz(ws.cell(row=row, column=COL["surface_form"]).value)
        top_entity     = nz(ws.cell(row=row, column=COL["top_entity"]).value)
        top_desc       = nz(ws.cell(row=row, column=COL["top_desc"]).value)
        shadow_entity  = nz(ws.cell(row=row, column=COL["shadow_entity"]).value)
        shadow_desc    = nz(ws.cell(row=row, column=COL["shadow_desc"]).value)
        source_excerpt = nz(ws.cell(row=row, column=COL["source_excerpt"]).value)
        model_response = clean_latex(
            nz(ws.cell(row=row, column=COL["model_response"]).value))

        if not model_response:
            skipped_existing += 1
            continue

        print(f"  [Row {row:>4}] {qid:<12} {variant:<12} {model:<22}"
              f" {model_response[:45]}…")

        # ── Judge 1 (column-restricted, has access to ground truth) ─────────
        j1 = call_judge(
            client, SYSTEM_JUDGE1,
            make_prompt_judge1(question, correct_answer, surface_form,
                               top_entity, top_desc, shadow_entity, shadow_desc,
                               source_excerpt, model_response)
        )
        time.sleep(PAUSE_BETWEEN_CALLS)

        # ── Judge 2 (world knowledge only, no ground truth) ─────────────────
        j2 = call_judge(
            client, SYSTEM_JUDGE2,
            make_prompt_judge2(question, surface_form, top_entity,
                               shadow_entity, model_response)
        )
        time.sleep(PAUSE_BETWEEN_CALLS)

        # ── Conflict resolution ──────────────────────────────────────────────
        j1c = j1["classification"]
        j2c = j2["classification"]

        if j1c.startswith("ERROR") or j2c.startswith("ERROR"):
            final_label  = "ERROR"
            needs_review = "YES"
            final_just   = f"J1={j1c} | J2={j2c}"
            errors += 1
        else:
            final_label, needs_review = resolve(j1c, j2c)
            final_just = (j1["justification"] if final_label != "Conflicting"
                          else f"J1: {j1['justification']} | J2: {j2['justification']}")

        # ── Write to cells ───────────────────────────────────────────────────
        ws.cell(row=row, column=COL["j1_class"]).value  = j1c
        ws.cell(row=row, column=COL["j1_conf"]).value   = j1["confidence"]
        ws.cell(row=row, column=COL["j1_just"]).value   = j1["justification"]
        ws.cell(row=row, column=COL["j2_class"]).value  = j2c
        ws.cell(row=row, column=COL["j2_conf"]).value   = j2["confidence"]
        ws.cell(row=row, column=COL["j2_just"]).value   = j2["justification"]
        ws.cell(row=row, column=COL["final_class"]).value  = final_label
        ws.cell(row=row, column=COL["needs_review"]).value = needs_review
        ws.cell(row=row, column=COL["final_just"]).value   = final_just

        processed += 1

        if processed % SAVE_EVERY == 0:
            wb.save(args.output)
            print(f"  → Auto-save after {processed} evaluations")

    wb.save(args.output)

    print()
    print("=" * 65)
    print(f"  Evaluated:           {processed}")
    print(f"  Skipped (existing):  {skipped_existing}")
    print(f"  Skipped (filter):    {skipped_filter}")
    print(f"  API errors:          {errors}")
    print(f"  Output:              {args.output}")
    print("=" * 65)


if __name__ == "__main__":
    main()
