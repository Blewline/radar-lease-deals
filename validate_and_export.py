#!/usr/bin/env python3
"""Validate the rebuilt 'All Deals' data and export clean website files.

Usage:
  python3 validate_and_export.py [--xlsx leasehackr_broker_master.xlsx | --json deals_in.json]
                                 [--out DIR] [--write-xlsx]

Writes to --out (default: this script's folder):
  deals.json, deals.csv, meta.json  -> clean rows only (same schema/keys as before)
  rejected_rows.csv                 -> every rejected row + reason(s)
Refuses to publish (exit 2, nothing overwritten) if fewer than 50% of the
previous meta.json deal_count pass validation.
--write-xlsx also rewrites the workbook's 'All Deals' tab with the clean rows.
"""
import argparse, csv, json, os, re, sys, datetime

COLS = ['Broker','LH Username','Source Tab','Tab Updated (per sheet)','Unit Type / Status','Year','Make','Model',
        'Trim / Description','MSRP (as listed)','MSRP ($)','MSRP Approx?','Monthly Payment ($)','Payment Basis',
        'Due at Signing (as listed)','DAS ($)','Term (mo)','Miles/yr','Region','Broker/Service Fee','Notes',
        'Effective Monthly ($)','Effective Monthly incl. Broker Fee ($)','% of MSRP (payment/MSRP)','Spreadsheet URL',
        'Posted Date','Expires','Source URL']
KEYS = ['broker','lh_username','source_tab','tab_updated_per_sheet','unit_type_status','year','make','model',
        'trim_description','msrp_as_listed','msrp_usd','msrp_approx','monthly_payment_usd','payment_basis',
        'due_at_signing_as_listed','das_usd','term_mo','miles_yr','region','broker_service_fee','notes',
        'effective_monthly_usd','effective_monthly_incl_broker_fee_usd','pct_of_msrp_payment_msrp','spreadsheet_url',
        'posted_date','expires','source_url']
COL2KEY = dict(zip(COLS, KEYS))
INTS = {'year','msrp_usd','das_usd','term_mo','miles_yr'}
# posted_date / expires: ISO 'YYYY-MM-DD', or 'YYYY-MM' when the source only names the month; null if unknown
NUMS = INTS | {'monthly_payment_usd','effective_monthly_usd','effective_monthly_incl_broker_fee_usd','pct_of_msrp_payment_msrp'}

SOLD_RE = re.compile(r'\b(sold(\s*out)?|program\s+ended)\b', re.I)
LABEL_RE = re.compile(r'\s*-\s*MULTIPLE UNITS\b', re.I)   # section label pasted into model text
HEADER_RE = re.compile(r'^(model|trim|payment|msrp|term|stock\s*#?|year|make|notes?|location|broker fee|'
                       r'(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*( deals)?|norcal|socal|out of stock)$', re.I)

def to_num(v, key):
    if v is None or (isinstance(v, str) and v.strip() == ''): return None
    try: f = float(str(v).replace(',', '').replace('$', ''))
    except ValueError: return None
    return int(round(f)) if key in INTS else (int(f) if f.is_integer() and key != 'pct_of_msrp_payment_msrp' else f)

def normalize(r):
    out = {}
    for k in KEYS:
        v = r.get(k)
        if k in NUMS: v = to_num(v, k)
        elif isinstance(v, str): v = v.strip() or None
        out[k] = v
    return out

def load(args):
    if args.json:
        return [normalize(r) for r in json.load(open(args.json))]
    import openpyxl
    ws = openpyxl.load_workbook(args.xlsx, data_only=True)['All Deals']
    it = ws.iter_rows(values_only=True); hdr = list(next(it))
    return [normalize({COL2KEY[h]: v for h, v in zip(hdr, row) if h in COL2KEY}) for row in it if any(c is not None for c in row)]

def clean(r):
    """Deterministic, source-preserving cleanups. Returns list of fixes applied."""
    fixes = []
    for k in ('model', 'trim_description'):
        if r[k] and LABEL_RE.search(r[k]):
            r[k] = LABEL_RE.sub('', r[k]).strip(); fixes.append(f'{k}: removed "MULTIPLE UNITS" label')
    if any('MULTIPLE UNITS' in f for f in fixes):
        r['notes'] = '; '.join(x for x in [r['notes'], 'Multiple units available (per sheet)'] if x)
    p, m = r['monthly_payment_usd'], r['msrp_usd']
    if p and m:  # keep derived column consistent
        r['pct_of_msrp_payment_msrp'] = round(p / m, 4)
    return fixes

def reasons(r):
    out = []
    y, p, m, t, mi = r['year'], r['monthly_payment_usd'], r['msrp_usd'], r['term_mo'], r['miles_yr']
    if y is None: out.append('year missing (not stated in source sheet)')
    elif not 2024 <= y <= 2028: out.append(f'year {y} outside 2024-2028')
    if p is None: out.append('payment missing')
    elif p < 99 or p > 5000: out.append(f'payment ${p} outside $99-$5,000 (misparsed)')
    if m is not None and not 15000 <= m <= 500000: out.append(f'MSRP ${m} outside $15k-$500k')
    if p and m and not 0.003 <= p / m <= 0.04: out.append(f'payment/MSRP {p/m:.2%} outside 0.3%-4%')
    if t is None or not 12 <= t <= 60: out.append(f'term {t} not in 12-60')
    if mi is not None and not 5000 <= mi <= 30000: out.append(f'miles {mi} outside 5k-30k')
    for k in ('make', 'model', 'trim_description'):
        v = r[k] or ''
        if k != 'trim_description' and not v: out.append(f'{k} missing')
        elif re.search(r'(?<![\d.])\s\d$', v) and not re.search(r'\s\d$', (r['trim_description'] or '') + ' ' if k == 'model' else ''):
            out.append(f'{k} "{v}" has trailing stray digit')
        elif HEADER_RE.match(v): out.append(f'{k} "{v}" is header/section text')
    if SOLD_RE.search(r['unit_type_status'] or '') or SOLD_RE.search(r['notes'] or ''):
        out.append('SOLD / SOLD OUT / PROGRAM ENDED')
    if re.search(r'malformed amount', r['notes'] or ''):
        out.append('malformed amount in source sheet (flagged by build.py, not guessed)')
    today = datetime.date.today().isoformat()
    for k in ('posted_date', 'expires'):
        v = r.get(k)
        if v is not None and not re.fullmatch(r'20\d\d-\d\d(-\d\d)?', str(v)): out.append(f'{k} "{v}" not ISO')
    if r.get('posted_date') and str(r['posted_date']) > today: out.append(f'posted_date {r["posted_date"]} in the future')
    if r.get('expires') and len(str(r['expires'])) == 10 and str(r['expires']) < today: out.append(f'offer expired {r["expires"]}')
    return out

def carryover_check(rows, rej):
    """Same car listed at 36 and 48 mo: a 48-mo payment >1.5x the 36-mo one is a copied cell."""
    idx = {}
    for i, r in enumerate(rows):
        idx.setdefault((r['broker'], r['source_tab'], r['year'], r['model'], r['trim_description'], r['msrp_usd']), {})[r['term_mo']] = i
    for g in idx.values():
        if 36 in g and 48 in g:
            a, b = rows[g[36]]['monthly_payment_usd'], rows[g[48]]['monthly_payment_usd']
            if a and b and b > 1.5 * a:
                rej.setdefault(g[48], []).append(f'48-mo payment ${b} vs 36-mo ${a}: copied/carry-over cell in sheet')

def main():
    here = os.path.dirname(os.path.abspath(__file__))
    ap = argparse.ArgumentParser()
    ap.add_argument('--xlsx', default='/workspace/leasehackr_broker_master.xlsx')
    ap.add_argument('--json'); ap.add_argument('--out', default=here)
    ap.add_argument('--write-xlsx', action='store_true')
    ap.add_argument('--data-month', help="YYYY-MM written to meta.json (default: current month)")
    a = ap.parse_args()
    rows = load(a)
    fixes = {i: clean(r) for i, r in enumerate(rows)}
    rej = {}
    for i, r in enumerate(rows):
        rs = reasons(r)
        if rs: rej[i] = rs
    carryover_check(rows, rej)
    seen = {}
    for i, r in enumerate(rows):
        if i in rej: continue
        k = json.dumps(r, sort_keys=True)
        if k in seen: rej[i] = [f'duplicate of row {seen[k] + 2}']
        else: seen[k] = i
    clean_rows = [r for i, r in enumerate(rows) if i not in rej]

    meta_path = os.path.join(a.out, 'meta.json')
    prev = json.load(open(meta_path)).get('deal_count', 0) if os.path.exists(meta_path) else 0
    with open(os.path.join(a.out, 'rejected_rows.csv'), 'w', newline='') as f:
        w = csv.writer(f); w.writerow(['source_row', 'reason'] + KEYS)
        for i in sorted(rej): w.writerow([i + 2, '; '.join(rej[i])] + [rows[i][k] for k in KEYS])
    print(f'input {len(rows)} | pass {len(clean_rows)} | rejected {len(rej)} | fixed {sum(1 for i,f in fixes.items() if f and i not in rej)} | previous {prev}')
    if prev and len(clean_rows) < 0.5 * prev:
        print(f'REFUSING TO PUBLISH: {len(clean_rows)} < 50% of previous {prev}', file=sys.stderr); sys.exit(2)

    json.dump(clean_rows, open(os.path.join(a.out, 'deals.json'), 'w'), indent=1, ensure_ascii=False)
    with open(os.path.join(a.out, 'deals.csv'), 'w', newline='') as f:
        w = csv.DictWriter(f, KEYS); w.writeheader()
        for r in clean_rows: w.writerow({k: ('' if v is None else v) for k, v in r.items()})
    pds = sorted(str(r['posted_date']) for r in clean_rows if r.get('posted_date'))
    json.dump({'updated_at': datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(),
               'data_month': a.data_month or datetime.date.today().strftime('%Y-%m'),
               'source': 'Leasehackr broker public Google Sheets + manufacturer lease offer pages', 'deal_count': len(clean_rows),
               'broker_count': len({r['broker'] for r in clean_rows}),
               'posted_date_min': pds[0] if pds else None, 'posted_date_max': pds[-1] if pds else None},
              open(meta_path, 'w'), indent=1)
    if a.write_xlsx and not a.json:
        import openpyxl
        wb = openpyxl.load_workbook(a.xlsx); ws = wb['All Deals']
        ws.delete_rows(2, ws.max_row)
        for r in clean_rows: ws.append([r[COL2KEY[c]] for c in COLS])
        for j, c in enumerate(COLS, 1): ws.cell(1, j, c)
        ws.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(COLS))}{ws.max_row}'
        wb.save(a.xlsx); print('updated All Deals tab in', a.xlsx)

if __name__ == '__main__':
    main()
