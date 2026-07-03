# IBGCScrapper

A Python scraper that extracts the complete directory of IBGC-certified governance professionals and exports it to a structured, analysis-ready Excel workbook.

---

## Context

The [IBGC (Instituto Brasileiro de Governança Corporativa)](https://www.ibgc.org.br) publishes a public registry of all professionals holding one of its ten certification categories — from board member credentials (CCA) to audit committee and governance officer certifications. A researcher needed this data collected, normalized, and structured into a spreadsheet with filterable columns per certificate type.

The catch: the registry spans **32 pages** of results on the website.

---

## The Technical Challenge

The naive approach — looping through 32 URLs and scraping each page — would not work here for two reasons:

1. **The server blocks direct HTTP requests.** The domain returns `403 Forbidden` to any non-browser client, ruling out `requests` + `BeautifulSoup`.

2. **There are no 32 pages to navigate.** After inspecting the page's HTML source, it became clear the site uses **[DataTables](https://datatables.net/) in client-side mode**: the entire dataset is loaded into the DOM on the very first request, and pagination is handled purely in JavaScript on the frontend. The "32 pages" exist only as a UI illusion.

This meant the right approach wasn't to paginate at all — it was to intercept the data that was already there.

---

## Solution

The scraper uses **Playwright** to launch a headless Chromium browser, load the page as a real user would, and then call the **DataTables JavaScript API** directly from Python via `page.evaluate()`:

```python
const api = new $.fn.dataTable.Api('#tableListaCCI');
return api.rows().data().toArray();
```

This single call bypasses all pagination and returns every record in the table's internal dataset — regardless of which "page" is currently displayed. All ~1,600+ professionals are extracted in one shot, with zero UI interaction.

A DOM-traversal fallback is also included for resilience:

```python
document.querySelectorAll('#tableListaCCI tbody tr').forEach(tr => { ... });
```

---

## Pipeline

```
Browser launch (Playwright)
       │
       ▼
Page load → networkidle
       │
       ▼
DataTables JS API call   ──(fallback)──▶  DOM traversal
       │
       ▼
Raw rows: [nome, estado, categorias_raw]
       │
       ▼
parse_rows()  →  normalize categories via exact CAT_MAP lookup
       │
       ▼
aggregate()   →  deduplicate by name, union certificate sets
       │
       ▼
build_excel() →  one general tab + one tab per certificate type
```

---

## Category Normalization

Each professional can hold multiple certificates, and the raw category strings from the site are verbose (e.g., `"CCA+ IBGC (Exame) (Certificado para Conselheiro de Administração Experiente IBGC)"`). The script maps every possible value to a canonical short form using an exact-match dictionary — no fragile regex patterns:

| Raw value (site) | Canonical column |
|---|---|
| CCA IBGC (...) | `CCA` |
| CCA+ IBGC (Exame) (...) | `CCA+ (Exame)` |
| CCA+ IBGC (Experiência) (...) | `CCA+ (Experiência)` |
| CCF IBGC (Exame) (...) | `CCF (Exame)` |
| CCF IBGC (Experiência) (...) | `CCF (Experiência)` |
| CCF+ IBGC (...) | `CCF+` |
| CCoAud IBGC (...) | `CCoAud` |
| CCoAud+ IBGC (...) | `CCoAud+` |
| CGO IBGC (...) | `CGO` |
| CGO+ IBGC (...) | `CGO+` |

---

## Output

The script generates `ibgc_certificados.xlsx` with the following structure:

**Tab "IBGC Certificados" (general):** every professional with a `✓` marker under each certificate they hold — designed for cross-filtering in Excel or Google Sheets.

| Nome | Estado | CCA | CCA+ (Exame) | CCA+ (Experiência) | CCF (Exame) | ... |
|---|---|---|---|---|---|---|
| Alexandre Silveira De Oliveira | São Paulo | ✓ | | | | ✓ |
| Ana Dolores Moura Carneiro De Novaes | Rio de Janeiro | | | ✓ | | |

**Tabs per certificate type:** one dedicated tab for each credential, listing only the professionals who hold it.

Additional formatting applied via `openpyxl`: header freeze, auto-filter, alternating row fills, and column widths calibrated to content.

---

## Stack

- **[Playwright](https://playwright.dev/python/)** — headless browser automation; handles JS rendering and server-side bot detection
- **[openpyxl](https://openpyxl.readthedocs.io/)** — Excel file generation with full style control

---

## Installation

```bash
pip install playwright openpyxl
playwright install chromium
```

## Usage

```bash
python ibgc_scraper.py
```

The script prints progress across its three stages and saves `ibgc_certificados.xlsx` in the current directory.

```
============================================================
  IBGC Certificados Scraper
============================================================

[1/3] Raspando dados (DataTables client-side) ...
  Abrindo https://www.ibgc.org.br/... 
  1631 linhas brutas capturadas.

[2/3] Parseando e agregando registros ...
  1598 profissionais únicos encontrados.
  Categorias: CCA, CCA+ (Exame), CCA+ (Experiência), CCF (Exame), ...

[3/3] Gerando Excel ...

✅  Arquivo salvo : ibgc_certificados.xlsx
    Profissionais  : 1598
    Categorias     : CCA, CCA+ (Exame), CCA+ (Experiência), ...
    Abas geradas   : 'IBGC Certificados' + 10 abas de categoria
```
