# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

This is the working directory for the **IHRM Database Metadata Project** (Center for
International HR Studies, Penn State). It is primarily a *research data* repository — it
holds source documents, an extraction notebook, and its outputs — plus a static-site
builder (`build_site.py`) that publishes the meta-database as a web page. It is a git
repository (remote: `github.com/Rakoon/_IHRM-MetaData`); there is no test suite.

We have identified which empirical datasets are used across ~13,587 international
HRM journal articles, and record each dataset's name, collecting organization, and
description. The meta-database latest version = *IHRM_DB_Metadata_Master_2026-06-05.xlsx*

The research goal: publish the meta-database table online. Make this table visually appealing
and informative to management researchers around the world. Potential publication / hosting 
place would be on Penn State Sites (https://sites.psu.edu/), which runs on Wordpress.

## Layout

- `IHRM_DB_Metadata_Master_<date>.xlsx` — the master metadata workbook (sheets:
  `Metadata` + `Codebook`). Multiple dated snapshots are kept side by side; the most
  recent date is the current version. **Do not overwrite an older snapshot** — they are
  the project's revision history.
- `build_site.py` — builds the published web page from the current master + `article_list.csv`
  (see below).
- `IHRM_DB_Metadata.html` — the generated, self-contained web page (committed output).
- `article_list.csv` — the article corpus index read by `build_site.py`. Columns:
  `Key_13587` (publication key), `Publication Title` (journal), `Author`, `Title`, `Year`.
  **Windows-1252 (cp1252) encoded**, not UTF-8.
- `backups/` — local-only safety copies made before bulk edits; git-ignored, not history.
- `output/dataset_article_list.csv` / `.xlsx` — a separate corpus index under `output/`
  (the extraction pipeline's input); distinct from the root `article_list.csv`.
- `output/txt/<KEY>_a.txt` — one plain-text file per article (~2,000 files), each
  containing `FILE NAME`, `AUTHORS`, `TITLE`, `ABSTRACT`, and a `METHOD (Data
  description)` section. These are the model input. `<KEY>` matches `Key_13587`.
- `output/chat_api_v3.ipynb` — the extraction pipeline (see below).
- `output/_Consolidated_240711.xlsx` — a consolidated results workbook.
- `Meeting_Notes/` — dated `.docx` project notes; reference only, not data.

## The extraction pipeline (`output/chat_api_v3.ipynb`)

Reads each `txt/<KEY>_a.txt`, sends it to the OpenAI Chat API with a fixed few-shot
prompt, and parses the reply into a dataframe of `key_13587, data_name, data_collector,
data_desc`, written to an `_extracted.csv`.

Key things to know before changing it:

- **Model:** `gpt-3.5-turbo` via the `openai` Python SDK (v1.x). Requires `OPENAI_API_KEY`
  in a `.env` file loaded with `python-dotenv`. The `.env` is not in this repo.
- **Few-shot prompt structure:** the conversation is hand-built as alternating
  user/assistant turns (`USER_PROMPT_1`..`8`, `ASSISTANT_PROMPT_1`..`7`). The entity
  definitions and the four worked examples live inside `analyze_text()`. Edit prompt
  wording there.
- **Output contract:** the model must return
  `Name | Source | Description ||| Name | Source | Description ...`. `text_to_dataframe()`
  splits on `|||` then `|` and **silently drops any segment that does not split into
  exactly 3 parts** — a prompt change that alters the delimiter count loses rows.
- **Hard-coded paths:** `file_path` and the final `to_csv()` target point at a specific
  user's OneDrive (`/Users/tomngo/.../faculty/tom/data/`). The notebook also expects a
  `key_13587.csv` list of publication keys in that folder. To run it here, repoint
  `file_path` at `output/` and use `dataset_article_list.csv` (first column = keys).

## The website builder (`build_site.py`)

Reads the `Metadata` sheet of the current master workbook and `article_list.csv`, and
writes a single self-contained `IHRM_DB_Metadata.html` (inline CSS + JS; client-side
search, sort, and a country filter). Run with `python build_site.py`.

- **Bumping to a new snapshot:** edit the `XLSX` and `SNAPSHOT_DATE` constants near the
  top of the file, then rerun.
- **`article_list.csv` is cp1252.** It must be read with `encoding="cp1252"`. Reading it
  as UTF-8 corrupts author names with diacritics (e.g. `Weiß`, `Kampkötter`) into the
  `�` replacement character.
- **`Metadata` columns expected** (2026-06-05): `#`, `Name`, `Owner`, `Link`,
  `Description`, `Category`, `Accessibility`, `How to access data`, `Years`, `Countries`,
  `Freq.`, `Citing Articles`, `Cadence`. Column lookup is by header name, so order is
  flexible, but a missing/renamed header will drop that field.

## Master workbook column conventions

The cleaning conventions below were established 2026-06-05; the `Codebook` sheet is the
authoritative in-file reference and should be kept in sync with any schema change.

- **`Freq.`** is the *citation count* (number of citing articles in the corpus), **not**
  collection frequency. Do not confuse it with `Cadence`.
- **`Years`** holds *temporal coverage only*: continuous span `YYYY–YYYY` (en dash),
  ongoing `YYYY–present`, discrete waves `YYYY, YYYY, …`, or blank if unknown.
- **`Cadence`** holds collection frequency: `Annual`, `Biennial`, `Every 4 years`,
  `Quarterly`, `Monthly`, `Daily`, `Continuous`, `Irregular`, `Varies`, … blank if
  unspecified.
- **`Countries`** is a controlled vocabulary: a canonical single country, a short explicit
  list (≤ ~6 names), or a group label — **`European countries`** (EU + non-EU Europe) or
  **`Global`** (worldwide / large-N / cross-continental), plus `Asia` / `North America`
  where a continent is the natural unit; blank if unknown. **No URLs** in this column
  (they belong in `How to access data`). When a long list/region text was collapsed to a
  group label, the full detail is preserved as a `Country coverage:` line appended to
  `Description`.

## Conventions

- Publication keys are 8-character alphanumeric codes (e.g. `22WHR4FY`). A few articles
  have a split second file suffixed `_a1.txt`.
- When producing new result files, follow the existing dated-snapshot naming
  (`<name>_<YYYY-MM-DD>.xlsx` or `_<YYMMDD>`) rather than overwriting.
- `.gitignore` is **whitelist-style** (everything ignored by default, files opted back in
  with `!`). When adding a new master snapshot that should be tracked, add a matching
  `!IHRM_DB_Metadata_Master_<date>.xlsx` line.
