"""Build a single self-contained HTML page for the IHRM Database Metadata Project.

Reads:
  - IHRM_DB_Metadata_Master_2026-06-05.xlsx (Metadata sheet)
  - article_list.csv (Key_13587, Publication Title, Author, Title, Year)

Writes:
  - IHRM_DB_Metadata.html
"""

from __future__ import annotations

import csv
import html
import json
import re
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
XLSX = ROOT / "IHRM_DB_Metadata_Master_2026-06-19.xlsx"
ARTICLES = ROOT / "article_list.csv"
OUT = ROOT / "IHRM_DB_Metadata.html"
SNAPSHOT_DATE = "2026-06-19"


def last_name(person: str) -> str:
    """'Smith, John A.'->Smith; 'Koyuncu B'->Koyuncu; 'John Smith'->Smith."""
    person = person.strip()
    if "," in person:
        return person.split(",", 1)[0].strip()
    parts = person.split()
    if not parts:
        return person
    # 'Surname Initials' (no comma), e.g. 'Koyuncu B', 'Krishnan SK', 'van Berg J':
    # the trailing token is a short all-caps initial block -> surname is the rest.
    if len(parts) >= 2 and re.fullmatch(r"[A-Z]{1,3}\.?", parts[-1]):
        return " ".join(parts[:-1])
    # 'John Smith' (Western Given Surname) fallback
    return parts[-1]


def cite_label(author_field: str, year: str) -> str:
    """APA-style in-text label: Smith (2019) / Smith & Jones (2019) / Smith et al. (2019)."""
    if not author_field:
        return f"Anonymous ({year})" if year else "Anonymous"
    authors = [a.strip() for a in author_field.split(";") if a.strip()]
    if not authors:
        return f"Anonymous ({year})" if year else "Anonymous"
    if len(authors) == 1:
        head = last_name(authors[0])
    elif len(authors) == 2:
        head = f"{last_name(authors[0])} & {last_name(authors[1])}"
    else:
        head = f"{last_name(authors[0])} et al."
    return f"{head} ({year})" if year else head


def load_articles() -> dict[str, dict]:
    out: dict[str, dict] = {}
    with ARTICLES.open("r", encoding="cp1252", newline="") as f:
        r = csv.DictReader(f)
        for row in r:
            key = (row.get("Key_13587") or "").strip()
            if not key:
                continue
            out[key] = {
                "journal": (row.get("Publication Title") or "").strip(),
                "authors": (row.get("Author") or "").strip(),
                "title": (row.get("Title") or "").strip(),
                "year": (row.get("Year") or "").strip(),
            }
    return out


URL_RE = re.compile(r"https?://[^\s,;<>\"')]+", re.IGNORECASE)


def parse_links(blob: str | None) -> list[str]:
    if not blob:
        return []
    found = URL_RE.findall(str(blob))
    # de-dup, preserve order
    seen = set()
    out = []
    for u in found:
        u = u.rstrip(".,;)")
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def split_citing(blob: str | None) -> list[str]:
    if not blob:
        return []
    return [k.strip() for k in str(blob).split(",") if k.strip()]


def clean(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() in {"#n/a", "n/a", "na", "none"}:
        return ""
    return s


# Group/region labels in the Countries controlled vocabulary that are not
# individual countries (excluded from the unique-country stat).
_GROUP_LABELS = {"Global", "European countries", "Asia", "North America", "Multiple countries"}


def normalize_country(s: str) -> list[str]:
    """Split a cleaned Countries cell into facet tokens.

    The Countries column now uses a controlled vocabulary: a canonical single
    country, a short comma-separated list, or a group label ('Global',
    'European countries', etc.). Split on commas/newlines, de-dup, keep order.
    """
    if not s:
        return []
    out: list[str] = []
    seen = set()
    for t in re.split(r"[,\n]", s):
        t = t.strip()
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def load_datasets(articles: dict[str, dict]) -> list[dict]:
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["Metadata"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    # Expected (2026-06-19): '#','Name','Owner','Link','Description','Category',
    # 'Accessibility','How to access data','Years','Countries','Freq.','Citing Articles',
    # 'Cadence','Level of Analysis'
    idx = {h: i for i, h in enumerate(header) if h}
    out: list[dict] = []
    for raw in rows[1:]:
        if raw is None or all(c is None or str(c).strip() == "" for c in raw):
            continue
        num = raw[idx["#"]]
        name = clean(raw[idx["Name"]])
        if not name:
            continue
        owner = clean(raw[idx["Owner"]])
        link_blob = clean(raw[idx["Link"]])
        description = clean(raw[idx["Description"]])
        category = clean(raw[idx["Category"]])
        accessibility = clean(raw[idx["Accessibility"]])
        years = clean(raw[idx["Years"]])
        cadence = clean(raw[idx["Cadence"]]) if "Cadence" in idx else ""
        level = clean(raw[idx["Level of Analysis"]]) if "Level of Analysis" in idx else ""
        countries_raw = clean(raw[idx["Countries"]])
        freq = raw[idx["Freq."]]
        try:
            freq_n = int(freq) if freq is not None and str(freq).strip() != "" else 0
        except (ValueError, TypeError):
            freq_n = 0
        citing_keys = split_citing(raw[idx["Citing Articles"]])

        citing = []
        for key in citing_keys:
            meta = articles.get(key)
            if meta:
                citing.append(
                    {
                        "key": key,
                        "label": cite_label(meta["authors"], meta["year"]),
                        "title": meta["title"],
                        "journal": meta["journal"],
                        "authors": meta["authors"],
                        "year": meta["year"],
                    }
                )
            else:
                citing.append(
                    {
                        "key": key,
                        "label": key,
                        "title": "",
                        "journal": "",
                        "authors": "",
                        "year": "",
                    }
                )
        # sort citations by year desc, then first-author last name
        citing.sort(
            key=lambda c: (-(int(c["year"]) if c["year"].isdigit() else 0), c["label"])
        )

        out.append(
            {
                "id": num,
                "name": name,
                "owner": owner,
                "links": parse_links(link_blob),
                "link_raw": link_blob,
                "description": description,
                "category": category,
                "accessibility": accessibility,
                "years": years,
                "cadence": cadence,
                "level": level,
                "countries_raw": countries_raw,
                "countries": normalize_country(countries_raw),
                "freq": freq_n,
                "citing": citing,
            }
        )
    return out


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>IHRM Database Metadata &mdash; Penn State</title>
<style>
  :root {
    --psu-navy: #1E407C;
    --psu-navy-dark: #14305c;
    --psu-accent: #96BEE6;
    --ink: #1a1a1a;
    --muted: #5a6470;
    --rule: #e1e5ec;
    --rule-strong: #c9d1dc;
    --bg: #ffffff;
    --bg-soft: #f6f8fb;
    --row-stripe: #fafbfd;
    --row-hover: #eef3fb;
    --chip-bg: #eef3fb;
    --chip-ink: #1E407C;
    --serif: Georgia, "Times New Roman", Cambria, serif;
    --sans: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  }
  .ihrm-app * { box-sizing: border-box; }
  .ihrm-app {
    font-family: var(--sans);
    color: var(--ink);
    background: var(--bg);
    line-height: 1.5;
    font-size: 15px;
    max-width: 1200px;
    margin: 0 auto;
    padding: 24px 18px 64px;
  }
  .ihrm-app h1, .ihrm-app h2, .ihrm-app h3 { font-family: var(--serif); color: var(--psu-navy); margin: 0; font-weight: 600; }
  .ihrm-app a { color: var(--psu-navy); }
  .ihrm-app a:hover { color: var(--psu-navy-dark); }

  .ihrm-header {
    border-top: 4px solid var(--psu-navy);
    padding-top: 18px;
    margin-bottom: 18px;
  }
  .ihrm-header h1 { font-size: 26px; letter-spacing: 0.2px; }
  .ihrm-header .sub { color: var(--muted); font-size: 14px; margin-top: 6px; }

  .ihrm-stats {
    display: flex;
    flex-wrap: wrap;
    gap: 18px;
    padding: 12px 14px;
    background: var(--bg-soft);
    border: 1px solid var(--rule);
    border-radius: 6px;
    margin: 14px 0 18px;
    font-size: 13.5px;
  }
  .ihrm-stat strong { color: var(--psu-navy); font-size: 17px; font-family: var(--serif); display: block; }
  .ihrm-stat span { color: var(--muted); }

  .ihrm-controls {
    display: grid;
    grid-template-columns: 2fr 1fr 1fr 1fr;
    gap: 10px;
    margin-bottom: 14px;
  }
  @media (max-width: 720px) { .ihrm-controls { grid-template-columns: 1fr 1fr; } }
  .ihrm-controls input, .ihrm-controls select {
    width: 100%;
    padding: 9px 11px;
    border: 1px solid var(--rule-strong);
    border-radius: 5px;
    font-size: 14px;
    background: #fff;
    color: var(--ink);
    font-family: inherit;
  }
  .ihrm-controls input:focus, .ihrm-controls select:focus {
    outline: 2px solid var(--psu-accent);
    border-color: var(--psu-navy);
  }
  .ihrm-result-count { color: var(--muted); font-size: 13px; margin: 4px 2px 8px; }

  table.ihrm-table {
    width: 100%;
    border-collapse: collapse;
    border-top: 2px solid var(--psu-navy);
    border-bottom: 2px solid var(--psu-navy);
    table-layout: fixed;
  }
  table.ihrm-table td, table.ihrm-table th { overflow-wrap: anywhere; word-break: normal; }
  table.ihrm-table thead th {
    text-align: left;
    background: var(--bg);
    border-bottom: 1px solid var(--rule-strong);
    color: var(--psu-navy);
    font-family: var(--serif);
    font-weight: 600;
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    padding: 10px 12px;
    cursor: pointer;
    user-select: none;
    white-space: nowrap;
  }
  table.ihrm-table thead th .arrow { font-size: 10px; opacity: 0.25; margin-left: 4px; display: inline-block; }
  table.ihrm-table thead th.sorted { color: var(--psu-navy-dark); }
  table.ihrm-table thead th.sorted .arrow { opacity: 1; }
  table.ihrm-table thead th:hover { background: var(--bg-soft); }
  table.ihrm-table tbody td {
    padding: 11px 12px;
    border-bottom: 1px solid var(--rule);
    vertical-align: top;
  }
  table.ihrm-table tbody tr.main-row { cursor: pointer; }
  table.ihrm-table tbody tr.main-row:nth-child(4n+1) { background: var(--row-stripe); }
  table.ihrm-table tbody tr.main-row:hover { background: var(--row-hover); }
  table.ihrm-table tbody tr.detail-row > td { padding: 0; border-bottom: 1px solid var(--rule-strong); background: var(--bg-soft); }
  td.col-num { color: var(--muted); text-align: right; font-variant-numeric: tabular-nums; }
  td.col-name {
    font-weight: 600; font-family: var(--serif); font-size: 15.5px;
    color: var(--psu-navy-dark);
    text-decoration: underline;
    text-decoration-color: transparent;
    text-decoration-thickness: 1.5px;
    text-underline-offset: 3px;
    transition: text-decoration-color 0.12s;
  }
  tr.main-row:hover td.col-name { text-decoration-color: var(--psu-navy); }
  tr.main-row.open td.col-name { text-decoration-color: var(--psu-navy); }
  td.col-category {
    color: #2a3340;
    font-size: 13px;
    line-height: 1.4;
  }
  td.col-category .cat-chip {
    display: inline-block;
    background: var(--chip-bg);
    color: var(--chip-ink);
    padding: 1px 7px;
    border-radius: 10px;
    margin: 1px 3px 1px 0;
    font-size: 12px;
    line-height: 1.5;
    white-space: nowrap;
  }
  td.col-level { color: #2a3340; font-size: 13px; line-height: 1.4; text-transform: capitalize; }
  td.col-countries { color: #2a3340; font-size: 13.5px; line-height: 1.4; }
  td.col-years { color: var(--muted); font-size: 13px; font-variant-numeric: tabular-nums; line-height: 1.4; }
  td.col-freq { text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; color: var(--psu-navy); }
  td.col-caret {
    text-align: center; color: var(--psu-navy);
    font-size: 14px; user-select: none;
  }
  td.col-caret .caret {
    display: inline-block;
    transition: transform 0.18s ease;
    opacity: 0.55;
  }
  tr.main-row:hover td.col-caret .caret { opacity: 1; }
  tr.main-row.open td.col-caret .caret { transform: rotate(180deg); opacity: 1; }

  .ihrm-row-hint {
    color: var(--muted);
    font-size: 15px;
    margin: 0 0 8px;
    padding-left: 2px;
  }
  .ihrm-row-hint .caret-mark {
    display: inline-block;
    color: var(--psu-navy);
    margin: 0 2px;
  }

  .detail-pad { padding: 16px 18px 22px 60px; }
  .detail-grid {
    display: grid;
    grid-template-columns: minmax(0, 2fr) minmax(0, 1fr);
    gap: 24px;
  }
  @media (max-width: 820px) { .detail-grid { grid-template-columns: 1fr; } }
  .detail h3 {
    font-size: 12px; text-transform: uppercase; letter-spacing: 0.7px; color: var(--psu-navy);
    margin: 0 0 6px; font-family: var(--sans); font-weight: 700;
  }
  .detail .block { margin-bottom: 16px; }
  .detail .desc { white-space: pre-wrap; color: #2a3340; font-size: 14.5px; line-height: 1.55; }
  .detail .meta-line { color: #2a3340; font-size: 14px; }
  .access-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 12.5px;
    font-weight: 600;
    letter-spacing: 0.2px;
  }
  .access-public  { background: #dff3e3; color: #1f6b32; }
  .access-collab  { background: #dde9f7; color: #1E407C; }
  .access-restr   { background: #fbe1cf; color: #8b3e0c; }
  .access-unspec  { background: #ebeef3; color: #5a6470; }
  .detail .link-list { list-style: none; padding: 0; margin: 0; }
  .detail .link-list li { margin: 3px 0; font-size: 13.5px; word-break: break-all; }
  .detail .pre { white-space: pre-wrap; color: #2a3340; font-size: 13.5px; }

  .citings { display: flex; flex-direction: column; gap: 4px; }
  .cite {
    display: block;
    padding: 6px 8px;
    border-left: 3px solid var(--psu-accent);
    background: #fff;
    border-radius: 0 4px 4px 0;
    font-size: 13.5px;
  }
  .cite .label { color: var(--psu-navy-dark); font-weight: 600; }
  .cite .title { color: #2a3340; }
  .cite .journal { color: var(--muted); font-style: italic; font-size: 12.5px; }
  .show-all-btn {
    margin-top: 6px;
    background: none;
    border: 1px solid var(--rule-strong);
    color: var(--psu-navy);
    padding: 5px 10px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 12.5px;
    font-family: inherit;
  }
  .show-all-btn:hover { background: var(--chip-bg); }

  mark.hit { background: #fff3a8; padding: 0 1px; border-radius: 2px; }

  .empty-state {
    text-align: center; padding: 40px 20px; color: var(--muted); font-style: italic;
    border: 1px dashed var(--rule-strong); border-radius: 6px; margin-top: 14px;
  }

  .footer-note {
    margin-top: 28px;
    padding-top: 14px;
    border-top: 1px solid var(--rule);
    color: var(--muted);
    font-size: 12.5px;
    line-height: 1.55;
  }
</style>
</head>
<body>
<div class="ihrm-app" id="ihrm-root">

  <header class="ihrm-header">
    <h1>IHRM Empirical Datasets &mdash; Meta-Database</h1>
    <div class="sub">Center for International HR Studies &middot; Penn State University &middot; Snapshot __SNAPSHOT__</div>
  </header>

  <div class="ihrm-stats">
    <div class="ihrm-stat"><strong id="stat-datasets">0</strong><span>datasets cataloged</span></div>
    <div class="ihrm-stat"><strong id="stat-citing">0</strong><span>citing journal articles</span></div>
    <div class="ihrm-stat"><strong id="stat-countries">0</strong><span>countries / regions covered</span></div>
  </div>

  <div class="ihrm-controls">
    <input type="search" id="q" placeholder="Search name, owner, description, country&hellip;" autocomplete="off">
    <select id="country-filter"><option value="">All countries</option></select>
    <select id="sort">
      <option value="freq-desc">Sort: most cited</option>
      <option value="freq-asc">Sort: least cited</option>
      <option value="name-asc">Sort: name (A&ndash;Z)</option>
      <option value="name-desc">Sort: name (Z&ndash;A)</option>
      <option value="id-asc">Sort: dataset # (asc)</option>
      <option value="id-desc">Sort: dataset # (desc)</option>
      <option value="category-asc">Sort: category (A&ndash;Z)</option>
      <option value="category-desc">Sort: category (Z&ndash;A)</option>
      <option value="level-asc">Sort: level of analysis (A&ndash;Z)</option>
      <option value="level-desc">Sort: level of analysis (Z&ndash;A)</option>
      <option value="countries-asc">Sort: countries (A&ndash;Z)</option>
      <option value="countries-desc">Sort: countries (Z&ndash;A)</option>
      <option value="years-asc">Sort: earliest year first</option>
      <option value="years-desc">Sort: latest year first</option>
    </select>
    <select id="page-size">
      <option value="25">25 per page</option>
      <option value="50" selected>50 per page</option>
      <option value="100">100 per page</option>
      <option value="9999">Show all</option>
    </select>
  </div>

  <div class="ihrm-result-count" id="result-count"></div>
  <div class="ihrm-row-hint">
    Click any row <span class="caret-mark">&#9662;</span> to view the dataset description, links, and citing articles.
  </div>

  <table class="ihrm-table" id="table">
    <colgroup>
      <col style="width: 52px">
      <col style="width: 30%">
      <col style="width: 22%">
      <col style="width: 15%">
      <col style="width: 17%">
      <col style="width: 11%">
      <col style="width: 36px">
    </colgroup>
    <thead>
      <tr>
        <th data-sort="id">#<span class="arrow"></span></th>
        <th data-sort="name">Dataset<span class="arrow"></span></th>
        <th data-sort="category">Category<span class="arrow"></span></th>
        <th data-sort="level">Level<span class="arrow"></span></th>
        <th data-sort="countries">Countries<span class="arrow"></span></th>
        <th data-sort="years">Years<span class="arrow"></span></th>
        <th aria-label="Expand"></th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
  <div id="empty" class="empty-state" style="display:none">No datasets match your filters.</div>

  <div class="footer-note">
    Click any row to view the dataset description, public documentation links, and the journal
    articles that cite it. Citing-article labels follow APA in-text style:
    &ldquo;Smith (Year)&rdquo;, &ldquo;Smith &amp; Jones (Year)&rdquo;, or &ldquo;Smith et al. (Year).&rdquo;
    Frequency counts reflect appearances in our corpus of approximately 13,587 international HRM journal articles.
  </div>
</div>

<script>
(function () {
  const DATA = __DATA__;

  const tbody = document.getElementById("tbody");
  const qEl = document.getElementById("q");
  const countryEl = document.getElementById("country-filter");
  const sortEl = document.getElementById("sort");
  const pageSizeEl = document.getElementById("page-size");
  const resultCountEl = document.getElementById("result-count");
  const emptyEl = document.getElementById("empty");

  // Stats
  document.getElementById("stat-datasets").textContent = DATA.length.toLocaleString();
  const allCiting = new Set();
  const allCountries = new Set();
  DATA.forEach(d => {
    d.citing.forEach(c => allCiting.add(c.key));
    d.countries.forEach(c => allCountries.add(c));
  });
  document.getElementById("stat-citing").textContent = allCiting.size.toLocaleString();
  // Exclude group/region labels so the stat counts individual countries only
  const GROUP_LABELS = new Set(["Global", "European countries", "Asia", "North America", "Multiple countries", "Multiple / varies"]);
  const realCountries = new Set([...allCountries].filter(c => !GROUP_LABELS.has(c)));
  document.getElementById("stat-countries").textContent = realCountries.size.toLocaleString();

  // Country dropdown
  const sortedCountries = [...allCountries].sort((a, b) => a.localeCompare(b));
  sortedCountries.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c; opt.textContent = c;
    countryEl.appendChild(opt);
  });

  const state = {
    query: "",
    country: "",
    sort: "freq-desc",
    pageSize: 50,
    openIds: new Set(),
  };

  function escapeHtml(s) {
    if (s == null) return "";
    return String(s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function escapeAttr(s) { return escapeHtml(s); }

  function highlight(text, q) {
    const safe = escapeHtml(text);
    if (!q) return safe;
    const re = new RegExp("(" + q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "ig");
    return safe.replace(re, '<mark class="hit">$1</mark>');
  }

  function matches(d, q) {
    if (!q) return true;
    const hay = [
      d.name, d.owner, d.description, d.category, d.accessibility,
      d.level, d.countries_raw, d.years, d.cadence,
      ...d.citing.map(c => c.label + " " + c.title + " " + c.authors + " " + c.journal),
    ].join("  ").toLowerCase();
    return hay.includes(q);
  }

  function firstYear(s) {
    if (!s) return Number.POSITIVE_INFINITY;
    const m = String(s).match(/\d{4}/);
    return m ? parseInt(m[0], 10) : Number.POSITIVE_INFINITY;
  }
  function lastYear(s) {
    if (!s) return Number.NEGATIVE_INFINITY;
    const all = String(s).match(/\d{4}/g);
    return all ? Math.max(...all.map(Number)) : Number.NEGATIVE_INFINITY;
  }
  const SORT_KEYS = {
    "id":        d => (d.id == null ? Number.POSITIVE_INFINITY : d.id),
    "name":      d => d.name.toLowerCase(),
    "category":  d => (d.category || "￿").toLowerCase(),
    "level":     d => (d.level || "￿").toLowerCase(),
    "countries": d => (d.countries_raw || "￿").toLowerCase(),
    "freq":      d => d.freq,
  };
  function compare(a, b, sortKey) {
    const dashIdx = sortKey.lastIndexOf("-");
    const col = sortKey.slice(0, dashIdx);
    const dir = sortKey.slice(dashIdx + 1) === "desc" ? -1 : 1;
    let av, bv;
    if (col === "years") {
      // asc = earliest year first; desc = latest year first
      av = dir === 1 ? firstYear(a.years) : lastYear(a.years);
      bv = dir === 1 ? firstYear(b.years) : lastYear(b.years);
    } else {
      const fn = SORT_KEYS[col] || SORT_KEYS["freq"];
      av = fn(a); bv = fn(b);
    }
    let cmp;
    if (typeof av === "number" && typeof bv === "number") cmp = av - bv;
    else cmp = String(av).localeCompare(String(bv));
    if (cmp === 0) cmp = a.name.localeCompare(b.name);
    return dir === -1 ? -cmp : cmp;
  }
  function updateSortIndicators() {
    const dashIdx = state.sort.lastIndexOf("-");
    const activeCol = state.sort.slice(0, dashIdx);
    const dir = state.sort.slice(dashIdx + 1);
    document.querySelectorAll("th[data-sort]").forEach(th => {
      const isActive = th.dataset.sort === activeCol;
      th.classList.toggle("sorted", isActive);
      const arrow = th.querySelector(".arrow");
      if (!arrow) return;
      if (isActive) arrow.innerHTML = dir === "desc" ? "&#9660;" : "&#9650;";
      else arrow.innerHTML = "&#9650;&#9660;";
    });
  }

  function categoryChipsHtml(catRaw, q) {
    if (!catRaw) return "";
    const parts = String(catRaw).split(/[;,]/).map(s => s.trim()).filter(Boolean);
    if (!parts.length) return escapeHtml(catRaw);
    return parts.map(p => `<span class="cat-chip">${highlight(p, q)}</span>`).join("");
  }
  function accessClass(s) {
    const k = (s || "").toLowerCase();
    if (k.includes("public"))     return "access-public";
    if (k.includes("collab"))     return "access-collab";
    if (k.includes("restrict"))   return "access-restr";
    return "access-unspec";
  }
  function accessBadgeHtml(s) {
    if (!s) return '<span class="access-badge access-unspec">Not Specified</span>';
    return `<span class="access-badge ${accessClass(s)}">${escapeHtml(s)}</span>`;
  }

  function citingHtml(d) {
    if (!d.citing.length) {
      return '<div class="meta-line" style="color:var(--muted)">No citing articles recorded.</div>';
    }
    const renderOne = c => {
      const label = escapeHtml(c.label);
      const title = c.title ? `<div class="title">${escapeHtml(c.title)}</div>` : "";
      const journal = c.journal ? `<div class="journal">${escapeHtml(c.journal)}</div>` : "";
      return `<div class="cite"><span class="label">${label}</span>${title}${journal}</div>`;
    };
    const initial = d.citing.slice(0, 5).map(renderOne).join("");
    if (d.citing.length <= 5) {
      return `<div class="citings">${initial}</div>`;
    }
    const rest = d.citing.slice(5).map(renderOne).join("");
    return (
      `<div class="citings" data-cites="${d._key}">` +
        initial +
        `<div class="more-cites" style="display:none">${rest}</div>` +
        `<button class="show-all-btn" data-toggle-cites="${d._key}">` +
          `Show all ${d.citing.length} citing articles` +
        `</button>` +
      `</div>`
    );
  }

  function detailHtml(d) {
    const desc = d.description
      ? `<div class="block"><h3>Description</h3><div class="desc">${escapeHtml(d.description)}</div></div>`
      : "";
    const yearsBlock = d.years
      ? `<div class="block"><h3>Years available</h3><div class="meta-line">${escapeHtml(d.years)}</div></div>`
      : "";
    const cadenceBlock = d.cadence
      ? `<div class="block"><h3>Collection cadence</h3><div class="meta-line">${escapeHtml(d.cadence)}</div></div>`
      : "";
    const countriesBlock = d.countries_raw
      ? `<div class="block"><h3>Countries</h3><div class="meta-line">${escapeHtml(d.countries_raw)}</div></div>`
      : "";
    const ownerBlock = d.owner
      ? `<div class="block"><h3>Owner</h3><div class="meta-line">${escapeHtml(d.owner)}</div></div>`
      : "";
    const accessBlock =
      `<div class="block"><h3>Accessibility</h3><div class="meta-line">${accessBadgeHtml(d.accessibility)}</div></div>`;
    let linksHtml = "";
    if (d.links.length) {
      const items = d.links.map(u => `<li><a href="${escapeAttr(u)}" target="_blank" rel="noopener noreferrer">${escapeHtml(u)}</a></li>`).join("");
      linksHtml = `<div class="block"><h3>Owner links</h3><ul class="link-list">${items}</ul></div>`;
    } else if (d.link_raw) {
      linksHtml = `<div class="block"><h3>Owner links</h3><div class="pre">${escapeHtml(d.link_raw)}</div></div>`;
    }
    const citingBlock = `<div class="block"><h3>Citing articles (${d.citing.length})</h3>${citingHtml(d)}</div>`;

    return (
      `<div class="detail-pad detail">` +
        `<div class="detail-grid">` +
          `<div>${desc}${yearsBlock}${cadenceBlock}${countriesBlock}${citingBlock}</div>` +
          `<div>${ownerBlock}${accessBlock}${linksHtml}</div>` +
        `</div>` +
      `</div>`
    );
  }

  function rowHtml(d, q) {
    const isOpen = state.openIds.has(d._key);
    const main =
      `<tr class="main-row${isOpen ? " open" : ""}" data-id="${d._key}" title="Click for details">` +
        `<td class="col-num">${escapeHtml(d.id != null ? String(d.id) : "")}</td>` +
        `<td class="col-name">${highlight(d.name, q)}</td>` +
        `<td class="col-category">${categoryChipsHtml(d.category, q)}</td>` +
        `<td class="col-level">${highlight(d.level, q)}</td>` +
        `<td class="col-countries">${highlight(d.countries_raw, q)}</td>` +
        `<td class="col-years">${highlight(d.years, q)}</td>` +
        `<td class="col-caret"><span class="caret">&#9662;</span></td>` +
      `</tr>`;
    if (!isOpen) return main;
    return main + `<tr class="detail-row"><td colspan="7">${detailHtml(d)}</td></tr>`;
  }

  function render() {
    const q = state.query.trim().toLowerCase();
    const country = state.country;
    let filtered = DATA.filter(d => {
      if (country && !d.countries.includes(country)) return false;
      return matches(d, q);
    });
    filtered.sort((a, b) => compare(a, b, state.sort));

    const total = filtered.length;
    const limited = filtered.slice(0, state.pageSize);
    // Only show the count line when a filter narrows the set
    // (otherwise it duplicates the per-page dropdown).
    resultCountEl.textContent =
      total === DATA.length
        ? ""
        : `Showing ${Math.min(limited.length, total)} of ${total} matching datasets (of ${DATA.length} total)`;

    if (!total) {
      tbody.innerHTML = "";
      emptyEl.style.display = "";
      return;
    }
    emptyEl.style.display = "none";
    updateSortIndicators();
    tbody.innerHTML = limited.map(d => rowHtml(d, q)).join("");
    // Reset stripe parity since detail rows interleave
    let stripeIdx = 0;
    [...tbody.querySelectorAll("tr.main-row")].forEach((tr) => {
      if (stripeIdx % 2 === 0) tr.style.background = "var(--row-stripe)";
      else tr.style.removeProperty("background");
      stripeIdx++;
    });
  }

  // Attach data keys for state tracking
  DATA.forEach((d, i) => { d._key = "ds" + i; });

  // Events
  let qTimer = null;
  qEl.addEventListener("input", () => {
    clearTimeout(qTimer);
    qTimer = setTimeout(() => { state.query = qEl.value; render(); }, 80);
  });
  countryEl.addEventListener("change", () => { state.country = countryEl.value; render(); });
  sortEl.addEventListener("change", () => { state.sort = sortEl.value; render(); });
  pageSizeEl.addEventListener("change", () => { state.pageSize = parseInt(pageSizeEl.value, 10) || 50; render(); });

  // Column-header sorting: click toggles direction on active column,
  // or switches to that column with a sensible default direction.
  const DEFAULT_DIR = { freq: "desc", id: "asc", name: "asc", category: "asc", level: "asc", countries: "asc", years: "asc" };
  document.querySelectorAll("th[data-sort]").forEach(th => {
    th.addEventListener("click", () => {
      const col = th.dataset.sort;
      const dashIdx = state.sort.lastIndexOf("-");
      const activeCol = state.sort.slice(0, dashIdx);
      const activeDir = state.sort.slice(dashIdx + 1);
      const nextDir = (col === activeCol)
        ? (activeDir === "asc" ? "desc" : "asc")
        : (DEFAULT_DIR[col] || "asc");
      state.sort = col + "-" + nextDir;
      // Keep dropdown in sync when the new sort is one of its options
      if ([...sortEl.options].some(o => o.value === state.sort)) sortEl.value = state.sort;
      render();
    });
  });

  // Row clicks (expand) and citation toggle
  tbody.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-toggle-cites]");
    if (btn) {
      const wrap = btn.parentElement;
      const more = wrap.querySelector(".more-cites");
      if (more) {
        const open = more.style.display !== "none";
        more.style.display = open ? "none" : "";
        btn.textContent = open
          ? "Show all " + (wrap.querySelectorAll(".cite").length) + " citing articles"
          : "Show fewer";
      }
      e.stopPropagation();
      return;
    }
    // Don't toggle row when clicking a link in the detail panel
    if (e.target.closest("a")) return;
    const tr = e.target.closest("tr.main-row");
    if (!tr) return;
    const id = tr.dataset.id;
    if (state.openIds.has(id)) state.openIds.delete(id); else state.openIds.add(id);
    render();
  });

  render();
})();
</script>
</body>
</html>
"""


def main() -> None:
    articles = load_articles()
    datasets = load_datasets(articles)
    payload = json.dumps(datasets, ensure_ascii=False)
    html_out = HTML_TEMPLATE.replace("__SNAPSHOT__", SNAPSHOT_DATE).replace("__DATA__", payload)
    OUT.write_text(html_out, encoding="utf-8")
    print(f"Wrote {OUT} ({OUT.stat().st_size:,} bytes; {len(datasets)} datasets)")


if __name__ == "__main__":
    main()
