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
