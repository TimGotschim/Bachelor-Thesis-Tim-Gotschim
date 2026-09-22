# Script inventory

This file maps every Python script referenced in the thesis to its final
version and records where that version came from. It is a working document
for assembling the repository; it does not need to be part of the final
submission.

Status legend:

- **INCLUDED** — final version is present in this repository
- **ON LOCAL MACHINE** — final version was produced and downloaded during
  development; retrieve it from the local project folder (see note below)

---

## Appendix D — Question generation

| # | Script | Category | API used | Status |
|---|--------|----------|----------|--------|
| D.1 | `collect_triples_fp.py` | Fictitious Premises (SPARQL triple collection) | none | **INCLUDED** |
| D.1 | `build_false_triples.py` | Fictitious Premises (false object generation) | Anthropic API, `claude-sonnet-4-20250514` | **INCLUDED** |
| D.1 | `generate_fp_from_triples.py` | Fictitious Premises (question generation) | WU Model Hub, `gpt-oss:120b` | **INCLUDED** |
| D.2 | `find_articles_and_fetch_triples.py` | Long-Tail Facts (article discovery + triples) | none | **INCLUDED** |
| D.2 | `generate_longtail_phase3_claude.py` | Long-Tail Facts (question generation) | Anthropic API, `claude-sonnet-4-5` | **INCLUDED** |
| D.3 | `confusable_entities_scaleup_pipeline.py` | Confusable Entities | Anthropic API, `claude-sonnet-4-6` | **INCLUDED** |
| D.4 | `generate_af_catalogue.py` | Authority Framing | WU Model Hub, `gpt-oss:120b` | **INCLUDED** |
| D.5 | `find_np_entities_high_traffic.py` | Numerical Precision (entity + triple discovery) | none | **INCLUDED** |
| D.5 | `generate_NP_catalogue_pipeline_gpt_oss_120b.py` | Numerical Precision (question generation) | WU Model Hub, `gpt-oss:120b` | **INCLUDED** |

Notes on version history, so that the correct file is picked from the local
folder:

- **D.2** — supplied and verified. The final version uses Claude Sonnet 4.5 via the
  Anthropic API, not gpt-oss:120b; the earlier gpt-oss variant is superseded.
- **D.4** — supplied and verified. Its two prompt constants were diffed against
  Appendix D.4 of the thesis and now match verbatim.
- **D.5** — supplied and verified. Cogito v2.1 was unavailable on the WU Model Hub, so
  gpt-oss:120b was used for generation.

---

## Appendix A — Response evaluation

| # | Script | Category | Architecture | Status |
|---|--------|----------|--------------|--------|
| A.1 | `judge_fp_new.py` | Fictitious Premises | Dual judge | **INCLUDED** |
| A.2 | `judge_phase3_longtailfacts.py` | Long-Tail Facts | Dual judge | **INCLUDED** |
| A.3 | `judge_confusable_entities.py` | Confusable Entities | Dual judge | **INCLUDED** |
| A.4 | `judge_authority_framing_new.py` | Authority Framing | Dual judge | **INCLUDED** |
| A.5 | `np_reeval.py` | Numerical Precision | Extraction (Claude Haiku 4.5) + Python | **INCLUDED** |
| A.5 | `np_evaluate.py` | Numerical Precision | Pure Python | optional, superseded by np_reeval.py |
| A.5 | `fill_np_recompute.py` | Numerical Precision | Pure Python | **INCLUDED** |
| A.5 | `finalize_np_catalogue.py` | Numerical Precision | Pure Python | **INCLUDED** |

Notes on version history:

- **A.1** — supplied and verified. Its four prompt constants were diffed
  against Appendix A.1 of the thesis and now match verbatim.
- **A.4** — supplied and verified. Its four prompt constants were diffed against
  Appendix A.4 of the thesis and now match verbatim.
- **A.5** — supplied and verified. Extraction runs on Claude Haiku 4.5 via the Anthropic
  API, not on the WU Model Hub; classification is pure Python inside the same script.

---

## Where to find the local files

Every script listed as ON LOCAL MACHINE was generated during development and
downloaded at the time. They are most likely in the Downloads folder or in the
thesis project folder. Sorting by date and searching for the filenames above
should locate them.

Before adding a file to the repository, check it against the version note in
the table above to be certain it is the final revision rather than an
intermediate one.
