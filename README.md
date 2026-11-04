# radar-lease-deals

Car lease deals pulled from Leasehackr brokers' public Google Sheets (snapshot Nov 4, 2026 PT). Figures are copied from each broker's own sheet; confirm with the broker before quoting.

## Files
- `deals.json` – array of deal objects (the "All Deals" sheet). snake_case keys, numbers as numbers, `null` for blanks.
- `deals.csv` – same deals as CSV.
- `brokers.json` – broker directory (name, username, region, thread/sheet URLs, contact, access status, notes).
- `meta.json` – `{updated_at, source, deal_count, broker_count}`.

## Fetching from a website
```js
const base = 'https://raw.githubusercontent.com/Blewline/radar-lease-deals/main/';
const [deals, brokers, meta] = await Promise.all(
  ['deals.json', 'brokers.json', 'meta.json'].map(f => fetch(base + f).then(r => r.json()))
);
```
raw.githubusercontent.com sends `Access-Control-Allow-Origin: *`, so browser fetches work directly. Content is cached up to ~5 minutes after a commit.

## Data validation (`validate_and_export.py`)
The site does no filtering, so every published row is validated first. After rebuilding the workbook's **All Deals** tab, run:
```bash
python3 validate_and_export.py --xlsx /workspace/leasehackr_broker_master.xlsx --write-xlsx
# or: python3 validate_and_export.py --json some_deals.json
```
It writes `deals.json`, `deals.csv`, `meta.json` (same schema/keys) and `rejected_rows.csv` (each rejected row with its reason). `--write-xlsx` also rewrites the All Deals tab with the clean rows.

Rows are rejected when:
- year is missing, or outside 2024–2028 (never inferred if the sheet doesn't state it)
- monthly payment is missing, under $99 or over $5,000
- MSRP (when present) is under $15k or over $500k, or payment/MSRP is outside 0.3%–4%
- term is not 12–60 months; miles/yr (when stated) is outside 5k–30k
- make/model/trim is missing, ends in a stray digit, or is header/section text (e.g. "Model", "OCTOBER")
- status or notes say SOLD, SOLD OUT or PROGRAM ENDED
- a 48-mo payment is >1.5× the same car's 36-mo payment (copied cell in the source sheet)
- exact duplicate of an earlier row

Safe cleanups are applied first (e.g. the "- MULTIPLE UNITS" label moved from the model name into notes). **Safety stop:** if fewer than 50% of the previous `meta.json` deal_count pass, the script exits with code 2 and does not overwrite the published files.
