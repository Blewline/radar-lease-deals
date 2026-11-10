# radar-lease-deals

Car lease deals pulled from Leasehackr brokers' public Google Sheets, plus manufacturer lease offers for direct-sale brands (snapshot Nov 10, 2026 PT, `data_month` 2026-11). Figures are copied from each source; confirm with the broker/manufacturer before quoting.

## Files
- `deals.json` – array of deal objects (the "All Deals" sheet). snake_case keys, numbers as numbers, `null` for blanks.
- `deals.csv` – same deals as CSV.
- `brokers.json` – broker directory (name, username, region, thread/sheet URLs, contact, access status, notes).
- `meta.json` – `{updated_at, data_month, source, deal_count, broker_count, posted_date_min, posted_date_max}`.
- `rejected_rows.csv` – every row that failed validation, with the reason.
- `build/build.py` – downloads/parses the broker sheets into the workbook's All Deals tab; `build/manufacturer_offers_YYYY-MM.json` – manufacturer offers copied from the maker's own page.

### Per-deal date/source fields
- `posted_date` – when the source says the deal was posted/updated: the tab's date stamp (e.g. "11.8 updated", "UPDATED: 11/04", "Last synced: Nov 10, 2026"); `YYYY-MM` when the tab/section only names a month (e.g. "NOVEMBER DEALS", "HN308's NOV 2026 DEALS"); otherwise the broker's latest post in the linked Leasehackr thread; `null` if unknown.
- `expires` – the source's stated end / "valid through" date (`null` if the source gives none).
- `source_url` – the sheet or page the row came from.
- Manufacturer rows use `broker` = `"Rivian (manufacturer)"` / `"Tesla (manufacturer)"`.

## Fetching from a website
```js
const base = 'https://raw.githubusercontent.com/Blewline/radar-lease-deals/main/';
const [deals, brokers, meta] = await Promise.all(
  ['deals.json', 'brokers.json', 'meta.json'].map(f => fetch(base + f).then(r => r.json()))
);
```
raw.githubusercontent.com sends `Access-Control-Allow-Origin: *`, so browser fetches work directly. Content is cached up to ~5 minutes after a commit.

## Refresh
```bash
python3 build/build.py --download --sheets /workspace/sheets_YYYYMMDD --threads threads.json \
  --extra build/manufacturer_offers_YYYY-MM.json --xlsx /workspace/leasehackr_broker_master.xlsx
python3 validate_and_export.py --xlsx /workspace/leasehackr_broker_master.xlsx --write-xlsx --data-month YYYY-MM
```
Money parsing is strict: a malformed figure in a sheet (e.g. PMI's Escalade "$2.,245") is flagged and rejected, never guessed.

## Data validation (`validate_and_export.py`)
The site does no filtering, so every published row is validated first. Rows are rejected when:
- year is missing, or outside 2024–2028 (never inferred if the sheet doesn't state it)
- monthly payment is missing, under $99 or over $5,000, or the source amount is malformed
- MSRP (when present) is under $15k or over $500k, or payment/MSRP is outside 0.3%–4%
- term is not 12–60 months; miles/yr (when stated) is outside 5k–30k
- make/model/trim is missing, ends in a stray digit, or is header/section text (e.g. "Model", "OCTOBER")
- status or notes say SOLD, SOLD OUT or PROGRAM ENDED
- a 48-mo payment is >1.5× the same car's 36-mo payment (copied cell in the source sheet)
- `expires` is in the past, `posted_date` is in the future, or either is not ISO
- exact duplicate of an earlier row

Safe cleanups are applied first (e.g. the "- MULTIPLE UNITS" label moved from the model name into notes). **Safety stop:** if fewer than 50% of the previous `meta.json` deal_count pass, the script exits with code 2 and does not overwrite the published files.
