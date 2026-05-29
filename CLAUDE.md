# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

This is the working directory for the **IHRM Database Metadata Project** (Center for
International HR Studies, Penn State). It is a *research data* repository, not a software
project — it holds source documents, an extraction notebook, and its outputs. There is no
build system, no test suite, and no version control.

We have identified which empirical datasets are used across ~13,587 international
HRM journal articles, and record each dataset's name, collecting organization, and
description. The meta-database latest version = *IHRM_DB_Metadata_Master_2026-05-21.xlsx*

The research goal: publish the meta-database table online. Make this table visually appealing
and informative to management researchers around the world. Potential publication / hosting 
place would be on Penn State Sites (https://sites.psu.edu/), which runs on Wordpress.

## Layout

- `IHRM_DB_Metadata_Master_<date>.xlsx` — the master metadata workbook. Multiple dated
  snapshots are kept side by side; the most recent date is the current version. **Do not
  overwrite an older snapshot** — they are the project's revision history.
- `output/dataset_article_list.csv` / `.xlsx` — the article corpus index. Columns:
  `Key_13587` (publication key), `Publication Title` (journal), `Author`, `Title`.
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

## Conventions

- Publication keys are 8-character alphanumeric codes (e.g. `22WHR4FY`). A few articles
  have a split second file suffixed `_a1.txt`.
- When producing new result files, follow the existing dated-snapshot naming
  (`<name>_<YYYY-MM-DD>.xlsx` or `_<YYMMDD>`) rather than overwriting.
