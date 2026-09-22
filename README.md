# When Language Models Make Things Up

Supplementary code for the bachelor's thesis *When Language Models Make Things
Up: An Empirical Study of Hallucination Triggers in Small, Locally Deployable
LLMs*.

Tim Gotschim — Vienna University of Economics and Business (WU Wien)
Supervisor: Dr. Svitlana Vakulenko

---

## What this repository contains

The study evaluates five hallucination trigger categories across three
open-weight instruction-tuned models under closed-book conditions. This
repository holds the Python scripts referenced in the thesis appendices, so
that both the construction of the question catalogues and the evaluation of
the model responses can be inspected and reproduced.

```
Appendix A/     response evaluation
Appendix D/     question generation
```

The folder names correspond to the appendix sections of the thesis, so that
each script can be read alongside the prompt it executes.

---

## Appendix D — Question generation

One script per trigger category. Each takes verified factual material
(Wikidata triples, Wikipedia extracts, or ShadowLink entity pairs) and prompts
a large language model to turn it into a natural-language question that
instantiates the category's trigger mechanism. The system prompts these
scripts contain are reproduced in Appendix D of the thesis.

| Script | Category | Generation model |
|--------|----------|------------------|
| `collect_triples_fp.py` | Fictitious Premises, step 1 | none (SPARQL) |
| `build_false_triples.py` | Fictitious Premises, step 2 | Claude Sonnet (Anthropic API) |
| `generate_fp_from_triples.py` | Fictitious Premises, step 3 | `gpt-oss:120b` (WU Model Hub) |
| `find_articles_and_fetch_triples.py` | Long-Tail Facts, step 1 | none (Wikimedia + Wikidata) |
| `generate_longtail_phase3_claude.py` | Long-Tail Facts, step 2 | Claude Sonnet 4.5 (Anthropic API) |
| `confusable_entities_scaleup_pipeline.py` | Confusable Entities | Claude Sonnet 4.6 (Anthropic API) |
| `generate_af_catalogue.py` | Authority Framing | `gpt-oss:120b` (WU Model Hub) |
| `generate_af_catalogue.py` | Authority Framing | `gpt-oss:120b` (WU Model Hub) |
| `find_np_entities_high_traffic.py` | Numerical Precision, step 1 | none (Wikidata + pageviews) |
| `generate_NP_catalogue_pipeline_gpt_oss_120b.py` | Numerical Precision, step 2 | `gpt-oss:120b` (WU Model Hub) |

Fictitious Premises runs as a three-stage pipeline rather than a single script.
`collect_triples_fp.py` retrieves verified subject-property-object triples from
Wikidata via SPARQL. `build_false_triples.py` then asks Claude Sonnet to replace
the true object with a plausible but definitively wrong alternative from the
same semantic class, together with a rationale that is later used as
ground-truth evidence by Judge 1. Only then does `generate_fp_from_triples.py`
embed the false object into a natural question. The split is deliberate:
choosing a convincing false object requires world knowledge, whereas phrasing
the question around a given object does not, so the expensive model is used
only where it is needed.

Every generated question was manually reviewed by the author before entering
the final catalogue. The generation models produce question text only; for
Numerical Precision the reference answers are computed deterministically in
Python from the Wikidata source values, never by the model.

---

## Appendix A — Response evaluation

Four of the five categories are evaluated by a dual-judge architecture in
which two instances of the same model class operate under asymmetric
information: Judge 1 sees only the designated ground-truth columns and is
instructed not to use its own knowledge, while Judge 2 sees only the question
and the response and evaluates from world knowledge alone. Disagreements are
not resolved automatically; they are flagged for manual review by the author.

| Script | Category |
|--------|----------|
| `judge_fp_new.py` | Fictitious Premises |
| `judge_phase3_longtailfacts.py` | Long-Tail Facts |
| `judge_confusable_entities.py` | Confusable Entities |
| `judge_authority_framing.py` | Authority Framing |

**Numerical Precision is evaluated differently and uses no judge model for
classification.** A language model is used only to extract the final stated
number from each response; the label is then assigned deterministically in
Python by comparing that number against the ground truth with a ±3% tolerance.
The reasoning behind this departure is given in Appendix A, Section A.5 of the
thesis.

| Script | Role |
|--------|------|
| `np_reeval.py` | Extraction (Claude Haiku 4.5) plus deterministic classification |
| `fill_np_recompute.py` | Refills the labels invalidated by the ground-truth audit |
| `finalize_np_catalogue.py` | Completes the formula documentation column |

---

## Ground-truth audit

A systematic error was found in the Numerical Precision catalogue after the
first evaluation run: the reference answers of the ratio and percentage items
had been computed using an entity that the question text did not name. All
affected items were recomputed from manually verified formulas, five
structurally defective questions were removed, and the catalogue was
re-evaluated. The hallucination rate for this category changed from 87.3% to
68.4% as a result.

The correction is documented row by row in the supplementary catalogue: the
column `GT_Formula_Used` records the formula from which each reference answer
is derived, using `E1`, `E2` and `E3` for the three converted entity values,
`avg()` for the mean of a subset, and the prefix `EXCESS` for items asking by
what percentage one value exceeds another. Every ground truth is therefore
independently verifiable from the raw entity values and the stated unit
conversion, without running any of the scripts in this repository.

---

## Requirements

```bash
pip install openai anthropic openpyxl pandas requests
```

The scripts that call the WU Model Hub expect two environment variables:

```bash
export WU_HUB_URL="https://web.ollama-gpt-oss.ai.wu.ac.at"
export WU_HUB_KEY="<your key>"
```

The scripts that call the Anthropic API expect:

```bash
export ANTHROPIC_API_KEY="<your key>"
```

No API keys are contained in this repository, and none of the scripts write a
key to disk.

---

## Reproducibility notes

All model inference was run locally through Ollama at temperature 0.0 with a
fixed random seed and a 300-token output cap, in a closed-book setting. The
evaluated models are Qwen2.5 7B, LLaMA 3.1 8B Instruct and Mistral 7B
Instruct.

Most scripts support resuming after an interrupted run, either through a
`--start-row` argument or through a progress file, because the full pipeline
involves several thousand API calls and network interruptions occurred
repeatedly during the study.

The full evaluated catalogues are submitted as supplementary Excel files
alongside the thesis rather than stored in this repository.
