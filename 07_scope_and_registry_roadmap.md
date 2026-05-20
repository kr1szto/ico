# Suggested Scope and Public Registry Roadmap

## Current implemented sources

The current scraper collects from:

| Source | Purpose |
|---|---|
| ORSR | commercial register details, statutory body, shareholders, legal facts |
| RPVS | public-sector partner data and beneficial owners |
| FinStat | financial and company summary data |
| RÚZ | accounting entity data |

## MVP product scope

The MVP should be a supplier due-diligence lookup by IČO.

It should provide:

1. One simple web input field for IČO.
2. Validated IČO normalization.
3. Source-by-source scraping.
4. JSON output for raw evidence.
5. CSV output for one-row business summary.
6. XLSX output for human review.
7. Basic transparent risk flags.

## Recommended future public registries

Add these later as separate source adapters.

| Source | Purpose | Priority |
|---|---|---:|
| RPO — Register právnických osôb | canonical identity, address, legal form, status | High |
| CRZ — Centrálny register zmlúv | supplier contracts, contract values, public counterparties | High |
| Insolvency / liquidation register | insolvency, restructuring, liquidation red flags | High |
| Obchodný vestník | older legal notices and events | Medium |
| Financial Administration / tax-debtor datasets | tax/debt indicators where available | Medium |
| ÚVO / UVOstat | public procurement exposure | Medium |
| EU sanctions list | sanctions screening for company names and persons | Medium |

## Avoid in MVP

Do not start with:

- court case mining
- general web search
- social media checks
- automated negative-news screening
- OCR-heavy PDF processing
- complex ownership graphing

These add noise and false positives before the core source adapters are stable.

## Recommended next technical improvement

Refactor `scrape_subject()` so one failed source does not break the whole lookup.

Target structure:

```python
{
    "ico": "36785512",
    "sources": {
        "orsr": {"status": "ok", "data": {}},
        "rpvs": {"status": "ok", "data": {}},
        "finstat": {"status": "error", "error": "..."},
        "ruz": {"status": "ok", "data": {}}
    }
}
```

This should be done before adding CRZ, RPO, insolvency, tax-debt, ÚVO or sanctions adapters.
