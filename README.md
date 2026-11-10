# radar-lease-deals

Car lease deals pulled from Leasehackr brokers' public Google Sheets, plus manufacturer lease offers for direct-sale brands (snapshot Nov 10, 2026 PT, `data_month` 2026-11). Figures are copied from each source; confirm with the broker/manufacturer before quoting.

## Files
- `deals.json` – array of deal objects (the "All Deals" sheet). snake_case keys, numbers as numbers, `null` for blanks.
- `deals.csv` – same deals as CSV.
- `brokers.json` – broker directory (name, username, region, thread/sheet URLs, contact, access status, notes).
- `meta.json` – `{updated_at, data_month, source, deal_count, broker_count, year_inferred_count, posted_date_min, posted_date_max}`.
- `rejected_rows.csv` – every row that failed validation, with the reason.
- `build/build.py` – downloads/parses the broker sheets into the workbook's All Deals tab; `build/manufacturer_offers_YYYY-MM.json` – manufacturer offers copied from the maker's own page.

### Per-deal date/source fields
- `posted_date` – when the source says the deal was posted/updated: the tab's date stamp (e.g. "11.8 updated", "UPDATED: 11/04", "Last synced: Nov 10, 2026"); `YYYY-MM` when the tab/section only names a month (e.g. "NOVEMBER DEALS", "HN308's NOV 2026 DEALS"); otherwise the broker's latest post in the linked Leasehackr thread; `null` if unknown.
- `expires` – the source's stated end / "valid through" date (`null` if the source gives none).
- `source_url` – the sheet or page the row came from.
- `year_inferred` – `true` when the sheet did not state the model year and `year` was inferred; `false` otherwise (CSV: `true`/`false`). Inference order: (1) a year section label above the row (e.g. Autologicx "25 Sentra"); (2) the same model with the same MSRP stated with a single year on the same tab; (3) a tab where every stated year (3+ rows) is the same; (4) otherwise the current model year for the sheet's month (2026 for the Nov 2026 sheets). The row's `notes` say which rule was used.
- Manufacturer rows use `broker` = `"Rivian (manufacturer)"` / `"Tesla (manufacturer)"`.

## Brokers
NewView Auto, Platinum Motor Imports, Lease Wizards, HN308, Lightning Auto, Genesis & Hyundai of Cumming GA, Autologicx, NYCLIFE, Stellar Auto (added Nov 10, 2026), plus Rivian manufacturer offers. Each row's `payment_basis` says whether the payment is pre-tax or includes tax; see `brokers.json` notes for each broker's quirks.

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
- year is outside 2024–2028 (a missing year is inferred, never left blank; see `year_inferred`)
- monthly payment is missing, under $99 or over $5,000, or the source amount is malformed
- MSRP (when present) is under $15k or over $500k, or payment/MSRP is outside 0.3%–4%
- term is not 12–60 months; miles/yr (when stated) is outside 5k–30k
- make/model/trim is missing, ends in a stray digit, or is header/section text (e.g. "Model", "OCTOBER")
- status or notes say SOLD, SOLD OUT or PROGRAM ENDED
- a 48-mo payment is >1.5× the same car's 36-mo payment (copied cell in the source sheet)
- `expires` is in the past, `posted_date` is in the future, or either is not ISO
- exact duplicate of an earlier row
- MSRP equals miles/yr (MSRP and mileage columns confused)
- build.py marked the row "⚠ … (not guessed)" because it could not verify it: find/replace-damaged cells (Autologicx "$39Mo/10K9"), a term string that disagrees with the Mileage column, a pre-tax payment higher than the with-tax one, a row that disagrees with its own Leasehackr calculator link (term/miles, MSRP more than 3% off, or Stellar payment more than $15 off), a model copied from another brand's tab, a tab the broker marks PENDING or labels with a month older than last month, Stellar inventory "NA" (none available), or NYCLIFE BMW price tiers out of order
- header text check uses whole month names, so a model like "Maybach" is not mistaken for "May"

Safe cleanups are applied first (e.g. the "- MULTIPLE UNITS" label moved from the model name into notes). **Safety stop:** if fewer than 50% of the previous `meta.json` deal_count pass, the script exits with code 2 and does not overwrite the published files.
