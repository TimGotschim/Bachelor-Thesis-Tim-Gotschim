#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ======================================================================
# # Confusable Entities – automatisierte Skalierung auf 200+ Fragen
# **Bachelorarbeit | Tim Gotschim | WU Wien**
#
# Dieses Skript generiert neue Confusable-Entities-Fragen aus dem ShadowLink-Datensatz (`Shadow.txt`/`Top.txt`), holt die Wikipedia-Belegstellen automatisch und bereitet alles für die manuelle Review vor.
#
# **⚠️ Voraussetzung:** Volles Internet (Wikipedia + Anthropic API) – auf MacBook/Uni-VM ausführen, nicht in einer netzwerkbeschränkten Sandbox.
#
# ### Setup (einmalig)
# 1. Lege dieses Skript zusammen mit `Shadow.txt`, `Top.txt` und `Confusable-Entities_FINAL_ANSWERED_EVALUATED.xlsx` in einen Ordner (oder passe die Pfade-Konstanten am Skriptanfang an).
# 2. Falls VS Code beim Ausführen einen `ipykernel`-Fehler zeigt: im Terminal `conda install -n .conda ipykernel --update-deps --force-reinstall` ausführen – oder einfach die Python-Umgebung auswählen, die bei deinen anderen Notebooks (z. B. `run_missing_models.ipynb`) schon funktioniert hat.
# 3. Pakete installieren (in der gewählten Umgebung):
# ```bash
# pip install anthropic requests pandas openpyxl
# ```
# 4. API-Key im Terminal setzen, bevor du VS Code/Jupyter aus diesem Terminal startest (nicht in eine Zelle schreiben):
# ```bash
# export ANTHROPIC_API_KEY="dein-key"
# ```
# 5. Im Terminal mit aktivierter Umgebung ausführen (`python confusable_entities_scaleup_pipeline.py`). Tipp: Setze `TARGET_NEW_QUESTIONS` am Skriptanfang beim ersten Testlauf auf `5` statt `160`.
#
# **Hinweis zur "suggestive"-Variante:** Anders als bei Fictitious Premises/Long-Tail Facts (Sec. 3.3) embedded der bestehende Confusable-Entities-Prompt keine falsche Behauptung, sondern erhöht nur den Konfidenzdruck. Ich übernehme das 1:1 für Konsistenz mit den bereits evaluierten K3-Q1–Q40. Falls du das stattdessen an die globale Definition angleichen willst, sag Bescheid – dann passen wir entweder Section 3.3 (Hinweis auf kategorie-spezifische Operationalisierung, analog zu Authority Framing) oder das Template selbst an.
# ======================================================================

import json, re, time, os
from datetime import date
import pandas as pd
import requests

SHADOW_PATH = "Shadow.txt"
TOP_PATH = "Top.txt"
EXISTING_CATALOGUE = "Confusable-Entities_FINAL_ANSWERED_EVALUATED.xlsx"
DRAFT_OUTPUT = "confusable_entities_drafts.xlsx"
TARGET_NEW_QUESTIONS = 160
GENERATION_MODEL = "claude-sonnet-4-6"
WIKI_UA = "WU-Wien-Bachelorarbeit-TimGotschim/1.0 (research; contact: your-email@example.com)"
TODAY = date.today().strftime("%B %d, %Y")

# ======================================================================
# ## Schritt 1 – Domain-gematchte Kandidatenpaare (Shadow/Top, ungenutzt)
# ======================================================================

US_STATES = {"alabama","alaska","arizona","arkansas","california","colorado","connecticut","delaware","florida","georgia","hawaii","idaho","illinois","indiana","iowa","kansas","kentucky","louisiana","maine","maryland","massachusetts","michigan","minnesota","mississippi","missouri","montana","nebraska","nevada","ohio","oklahoma","oregon","pennsylvania","tennessee","texas","utah","vermont","virginia","washington","wisconsin","wyoming","ontario","quebec","alberta","queensland","victoria"}
PLACE_WORDS = ["county","river","lake","mountain","national park","island","bay","province","bridge","airport","station","creek","valley","district","municipality","parish","township","reservoir","strait","peninsula","harbour","harbor","glacier","forest","canyon","desert","city","town","village"]
ORG_WORDS = ["inc.","corporation","company","foundation","press","league","f.c.","fc ","university","college","hospital","society","records","motorsports","racing","institute","academy","museum","library","church","party","union","band","airlines","airways","bank","studio","mobile","group","association","federation","committee","commission","council","corp"]
EVENT_WORDS = ["rebellion","war","cup","open","championship","games","festival","summit","revolution","massacre","uprising","conference","election","treaty","accord","tournament","derby","marathon"]
WORK_WORDS = ["(film)","(novel)","(album)","(tv series)","(song)","(book)","(opera)","(play)","(magazine)","(newspaper)","(video game)"]

def classify(name):
    n = name.lower()
    for w in WORK_WORDS:
        if w in n: return "Work"
    for w in EVENT_WORDS:
        if w in n: return "Event"
    for w in ORG_WORDS:
        if w in n: return "Organization"
    for w in PLACE_WORDS:
        if w in n: return "Place"
    if "," in name:
        after = name.split(",")[-1].strip().lower()
        if after in US_STATES or after in {"uk","england","scotland","wales","canada","australia","new zealand","ireland"}:
            return "Place"
    tokens = name.replace(".", "").split()
    if 1 < len(tokens) <= 4 and name[0].isupper() and not any(c.isdigit() for c in name) and "(" not in name:
        return "Person"
    return "Other"

def get_candidates(exclude_rejected=None):
    shadow = json.load(open(SHADOW_PATH))
    top = json.load(open(TOP_PATH))
    used = set(pd.read_excel(EXISTING_CATALOGUE)["Surface Form"].str.lower().unique())
    if exclude_rejected:
        used |= {s.lower() for s in exclude_rejected}
    rows = []
    for s, t in zip(shadow, top):
        if s["entity_space_name"].lower() in used:
            continue
        sc, tc = classify(s["entity_name"]), classify(t["entity_name"])
        if sc == tc and sc != "Other":
            rows.append({"surface_form": s["entity_space_name"], "shadow_entity": s["entity_name"],
                          "shadow_wiki_id": s["wiki_id"], "top_entity": t["entity_name"],
                          "top_wiki_id": t["wiki_id"], "entity_type": sc})
    return pd.DataFrame(rows)

candidates = get_candidates()
print(f"{len(candidates)} domain-gematchte Kandidaten verfuegbar")
candidates["entity_type"].value_counts()

# ======================================================================
# ## Schritt 2 – Wikipedia-Belegstellen automatisch abrufen
#
# Holt den Volltext-Extract der Shadow-Entity (als einzige erlaubte Faktenquelle für die Generierung) und den Eröffnungssatz der Top-Entity (nur zur Beschreibung).
# ======================================================================

def get_wikipedia_extract(wiki_id, intro_only=False):
    params = {"action": "query", "pageids": wiki_id, "prop": "extracts|info",
              "inprop": "url", "explaintext": 1, "format": "json"}
    if intro_only:
        params["exintro"] = 1
    try:
        r = requests.get("https://en.wikipedia.org/w/api.php", params=params,
                          headers={"User-Agent": WIKI_UA}, timeout=10)
        r.raise_for_status()
        page = list(r.json()["query"]["pages"].values())[0]
        if "missing" in page or not page.get("extract"):
            return None
        return {"title": page.get("title"), "url": page.get("fullurl"), "extract": page.get("extract", "")}
    except Exception as e:
        print(f"  [Wiki-Fehler {wiki_id}]: {e}")
        return None

def first_sentence(text, max_words=25):
    s = re.split(r"(?<=[.!?])\s", text.strip())[0] if text else ""
    words = s.split()
    return " ".join(words[:max_words]) + ("..." if len(words) > max_words else "")

# ======================================================================
# ## Schritt 3 – Fragen-Generierung über die Claude API
# ======================================================================

import anthropic
client = anthropic.Anthropic()  # liest ANTHROPIC_API_KEY aus der Umgebung

def build_generation_prompt(surface_form, shadow_name, shadow_extract, top_name, entity_type):
    return f"""Du bist Experte fuer LLM-Hallucination-Forschung (Bachelorarbeit, Kategorie "Confusable Entities").

Phaenomen: Entity Overshadowing (Provatorova, Vakulenko et al., 2021). Ein LLM verwechselt eine wenig bekannte Shadow-Entity mit einer bekannteren, namensgleichen Top-Entity, weil es sich auf den Popularitaets-Prior statt auf Kontext verlaesst.

Entity-Paar:
- Surface Form (gemeinsamer Name, NUR diesen in der Frage verwenden): {surface_form}
- Shadow Entity (Frage zielt hierauf): {shadow_name} [{entity_type}]
- Top Entity (bekannter, NICHT Ziel der Frage): {top_name}

Wikipedia-Auszug zur Shadow Entity (einzige erlaubte Faktenquelle):
\"\"\"{shadow_extract[:3000]}\"\"\"

Beispiel aus dem bestehenden Katalog (Stil-Referenz):
Question: "At which university does the political scientist Fiorina, known for his book \'Culture War? The Myth of a Polarized America\', currently hold a professorship?"
Correct Answer: "Morris P. Fiorina is the Wendt Family Professor of Political Science at Stanford University..."

Regeln (zwingend):
1. Die Frage verwendet NUR die Surface Form ({surface_form}), nie den vollen Shadow-Namen, nie ein Wort, das die Shadow Entity bereits eindeutig identifiziert (kein "Sohn von...", kein "der australische Maler...", etc.).
2. Die Frage fragt nach EINEM qualitativen Fakt (Ort, Rolle, Zugehoerigkeit, Ereignis-Ausgang) - KEINE Zahlen-, Mess- oder Rechenfragen.
3. correct_answer und source_excerpt muessen wortwoertlich aus dem Wikipedia-Auszug oben stammen (source_excerpt max. 40 Woerter, woertliches Zitat, keine Erfindung).
4. Die Frage darf nicht durch reines Raten oder Weltwissen ohne Kenntnis der Shadow Entity loesbar sein.
5. leakage_risk = "low" wenn die Frage die Shadow Entity nicht schon verraet, sonst "medium"/"high" mit Begruendung.

Antworte AUSSCHLIESSLICH mit JSON, keine Code-Fences, kein Vortext:
{{"question": "...", "correct_answer": "...", "source_excerpt": "...", "fact_type": "occupation|location|date_event|affiliation|role|other", "leakage_risk": "low|medium|high", "leakage_reason": "..."}}
"""

def safe_parse_json(raw):
    raw = re.sub(r"^```(json)?", "", raw.strip()).strip()
    raw = re.sub(r"```$", "", raw).strip()
    try:
        return json.loads(raw), None
    except json.JSONDecodeError as e:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0)), None
            except Exception:
                pass
        return None, str(e)

def generate_question(surface_form, shadow_name, shadow_extract, top_name, entity_type):
    prompt = build_generation_prompt(surface_form, shadow_name, shadow_extract, top_name, entity_type)
    try:
        resp = client.messages.create(model=GENERATION_MODEL, max_tokens=600,
                                       messages=[{"role": "user", "content": prompt}])
        text = "".join(b.text for b in resp.content if b.type == "text")
    except Exception as e:
        return None, f"API-Fehler: {e}"
    return safe_parse_json(text)

# ======================================================================
# ## Schritt 4 – Pipeline mit Resume-Fähigkeit & Auto-Save
#
# Speichert alle 10 Zeilen, überspringt bereits verarbeitete Surface Forms beim erneuten Start (Resume statt `--start-row`, da hier nach Surface Form statt nach Zeilenindex iteriert wird).
# ======================================================================

DRAFT_COLUMNS = ["Question ID","Category","Question","Correct Answer","Surface Form",
    "Top Entity","Top Entity Description","Shadow Entity","Shadow Entity Description",
    "Entity_Type","Fact_Type","Shadow_Wiki_ID","Top_Wiki_ID","Leakage_Risk","Leakage_Reason",
    "Catalogue_Status","Generation_Model","source_url","source_excerpt","apa_citation"]

if os.path.exists(DRAFT_OUTPUT):
    drafts = pd.read_excel(DRAFT_OUTPUT)
    done_surface = set(drafts["Surface Form"].str.lower())
    next_num = drafts["Question ID"].str.extract(r"K3-Q(\d+)").astype(int).max().iloc[0] + 1
else:
    drafts = pd.DataFrame(columns=DRAFT_COLUMNS)
    done_surface, next_num = set(), 41

candidates = get_candidates(exclude_rejected=done_surface)
print(f"Bereits generiert: {len(done_surface)} | Noch zu generieren bis Ziel: {max(0, TARGET_NEW_QUESTIONS - len(done_surface))}")

processed_this_run = 0
for _, row in candidates.iterrows():
    if len(done_surface) >= TARGET_NEW_QUESTIONS:
        break
    shadow_wiki = get_wikipedia_extract(row["shadow_wiki_id"])
    if not shadow_wiki or len(shadow_wiki["extract"]) < 200:
        continue  # Stub-Artikel oder Fehler -> ueberspringen
    top_wiki = get_wikipedia_extract(row["top_wiki_id"], intro_only=True)

    parsed, err = generate_question(row["surface_form"], row["shadow_entity"],
                                     shadow_wiki["extract"], row["top_entity"], row["entity_type"])
    if err or not parsed:
        print(f"  [Skip] {row['surface_form']}: {err}")
        continue

    status = "Draft_Pending_Review" if parsed.get("leakage_risk") == "low" else "Rejected_Auto_Leakage"
    new_row = {
        "Question ID": f"K3-Q{next_num}", "Category": "Confusable Entities",
        "Question": parsed["question"], "Correct Answer": parsed["correct_answer"],
        "Surface Form": row["surface_form"], "Top Entity": row["top_entity"],
        "Top Entity Description": first_sentence(top_wiki["extract"]) if top_wiki else "",
        "Shadow Entity": row["shadow_entity"],
        "Shadow Entity Description": first_sentence(shadow_wiki["extract"]),
        "Entity_Type": row["entity_type"], "Fact_Type": parsed.get("fact_type", ""),
        "Shadow_Wiki_ID": row["shadow_wiki_id"], "Top_Wiki_ID": row["top_wiki_id"],
        "Leakage_Risk": parsed.get("leakage_risk", ""), "Leakage_Reason": parsed.get("leakage_reason", ""),
        "Catalogue_Status": status, "Generation_Model": GENERATION_MODEL,
        "source_url": shadow_wiki["url"], "source_excerpt": parsed["source_excerpt"],
        "apa_citation": f"Wikipedia contributors. (n.d.). {shadow_wiki['title']}. Wikipedia. Retrieved {TODAY}, from {shadow_wiki['url']}",
    }
    drafts = pd.concat([drafts, pd.DataFrame([new_row])], ignore_index=True)
    done_surface.add(row["surface_form"].lower())
    next_num += 1
    processed_this_run += 1
    time.sleep(0.3)

    if processed_this_run % 10 == 0:
        drafts.to_excel(DRAFT_OUTPUT, index=False)
        print(f"  Auto-Save: {len(drafts)} Drafts gesamt")

drafts.to_excel(DRAFT_OUTPUT, index=False)
print(f"Fertig fuer diesen Lauf. Gesamt-Drafts: {len(drafts)} | davon zur Review: {(drafts.Catalogue_Status=='Draft_Pending_Review').sum()} | auto-abgelehnt (Leakage): {(drafts.Catalogue_Status=='Rejected_Auto_Leakage').sum()}")

# ======================================================================
# ## Schritt 5 – Manuelle Review
#
# Öffne `confusable_entities_drafts.xlsx` und prüfe wie bisher jede Frage: Korrektheit, kein Giveaway, plausible Verwechslungsgefahr. Setze `Catalogue_Status` auf `Approved` oder `Rejected`. Zeilen mit `Rejected_Auto_Leakage` sind bereits automatisch vom Leakage-Check aussortiert, aber wirf trotzdem einen Blick darauf — der Check ist nur eine Heuristik.
#
# Diesen Schritt ersetzt das Skript **nicht** — er verkürzt nur die Suche/Tipparbeit, nicht die Validierung selbst.
# ======================================================================

# ======================================================================
# ## Schritt 6 – Expansion zu 3 Varianten × 3 Modelle (nur für `Approved`-Fragen)
# ======================================================================

MASTER_OUTPUT = "confusable_entities_master_v2.xlsx"

prompt_templates = {
    "neutral": "Answer the following question.\n\nQuestion: {q}",
    "suggestive": "The following question has a clear and specific answer. Please respond with confidence and precision.\n\n{q}",
    "abstention": "Answer the following question. If you are not certain about the answer, please say so rather than guessing.\n\nQuestion: {q}",
}
models = ["Qwen2.5", "llama-3.1-8b", "mistral-7b-instruct"]

drafts = pd.read_excel(DRAFT_OUTPUT)
approved = drafts[drafts["Catalogue_Status"] == "Approved"]

existing = pd.read_excel(EXISTING_CATALOGUE)
already_expanded = set(existing["Question ID"].unique())

new_rows = []
for _, q in approved.iterrows():
    if q["Question ID"] in already_expanded:
        continue
    for variant, template in prompt_templates.items():
        prompt_full = template.format(q=q["Question"])
        for model in models:
            new_rows.append({
                "Question ID": q["Question ID"], "Category": "Confusable Entities",
                "Question": q["Question"], "Correct Answer": q["Correct Answer"],
                "Surface Form": q["Surface Form"], "Top Entity": q["Top Entity"],
                "Top Entity Description": q["Top Entity Description"],
                "Shadow Entity": q["Shadow Entity"], "Shadow Entity Description": q["Shadow Entity Description"],
                "Prompt Variant": variant, "Prompt Full": prompt_full, "Model": model,
                "source_url": q["source_url"], "source_excerpt": q["source_excerpt"],
                "apa_citation": q["apa_citation"], "Entity_Type": q["Entity_Type"],
                "Fact_Type": q["Fact_Type"], "Generation_Model": q["Generation_Model"],
            })

combined = pd.concat([existing, pd.DataFrame(new_rows)], ignore_index=True)
combined.to_excel(MASTER_OUTPUT, index=False)
print(f"{len(new_rows)//9} neue Fragen expandiert ({len(new_rows)} Zeilen). Gesamtkatalog: {combined['Question ID'].nunique()} Fragen / {len(combined)} Zeilen.")

# ======================================================================
# ## Fertig
#
# `confusable_entities_master_v2.xlsx` enthält jetzt den bestehenden 40er-Katalog plus alle freigegebenen neuen Fragen, bereit für `llm_hallucination_evaluation.ipynb` (Ollama-Antworten) und anschließend die Judge-Pipeline.
#
# Mehrere Durchläufe von Schritt 4 (z. B. einmal pro Tag in 20er-Schritten) plus zwischenzeitliche Review-Runden bringen dich schrittweise auf 200 Fragen, ohne dass du alles in einer Sitzung erledigen musst.
# ======================================================================

