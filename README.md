# IHRM Empirical Datasets — Meta-Database

A searchable, sortable web view of empirical datasets used in international human
resource management research. Curated from a corpus of approximately 13,587 IHRM
journal articles by the Center for International HR Studies at Penn State.

**Status:** Alpha v1.0

## What's here

| File | Purpose |
|---|---|
| `build_site.py` | Generator: reads the master xlsx + article list, writes a single self-contained HTML page. |
| `IHRM_DB_Metadata.html` | The latest built page. Open it directly in a browser, or embed it in WordPress. |
| `IHRM_DB_Metadata_Master_<date>.xlsx` | The master metadata workbook. Each dated snapshot is a revision. |
| `article_list.csv` | Index of the journal-article corpus (`Key_13587, Publication Title, Author, Title, Year`). |
| `CLAUDE.md` | Project context. |

## Build

Requires Python 3.10+ and `openpyxl`.

```bash
pip install openpyxl
python build_site.py
```

The script reads the most recent master xlsx referenced at the top of `build_site.py`
and writes `IHRM_DB_Metadata.html` in the project root.

## The page

- 190 datasets across 60 country/region facets, 328 citing journal articles
- Full-text search across name, owner, description, country, years, and citation labels
- Country filter dropdown
- Click any column header to sort (toggleable direction); click any row to expand
- Citing articles formatted as `Smith (Year)` / `Smith & Jones (Year)` / `Smith et al. (Year)`
  with full title and journal underneath
- Designed to drop into a WordPress Custom HTML block or be embedded via iframe
  on sites.psu.edu

## License

Code, data, and the generated page are released under
[Creative Commons Attribution 4.0 International (CC BY 4.0)](https://creativecommons.org/licenses/by/4.0/).
Please cite the Center for International HR Studies at Penn State when reusing.
