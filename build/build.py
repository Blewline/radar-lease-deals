#!/usr/bin/env python3
"""Rebuild the master workbook's 'All Deals' tab (+ raw tabs) from the brokers' public Google Sheets.

Usage: python3 build.py [--sheets DIR] [--download] [--threads threads.json] [--extra extra_deals.json]
       [--xlsx /workspace/leasehackr_broker_master.xlsx] [--out-json all_deals_raw.json]

* --download re-exports each sheet via https://docs.google.com/spreadsheets/d/<id>/export?format=xlsx
* --threads: forum thread dump (from lh gather) used for posted_date fallbacks
* --extra: JSON list of already-shaped deals (e.g. manufacturer offers) appended as-is
Money parsing is strict: a malformed figure such as "$2.,245" is NOT guessed; the row is kept with
payment=None and a flag in notes so validate_and_export.py rejects it with a clear reason.
"""
import argparse, datetime, json, os, re, subprocess, sys, urllib.parse
import openpyxl

TODAY = datetime.date(2026, 11, 10)
SHEETS = {
    'nev': '1a2LGPqckDi_qBiDVECBrroh5M2v5YcqrAIaFGN6Sl5U', 'pmi': '1lJp8CvddZHpt4iDFJGXCSHxvUqzDtWOqvrgCla5145Y',
    'wiz': '17XIc0WpYVeQ5noZ66pkJrPoIYMXQ92cjj43yvMOOBv4', 'hn': '1wj42phKxrEIMgMo8BPqsweSq0zwZR-gA6e2IK3Cy4JY',
    'lig': '1pq7luu8sK1-BB63sKdiGzEnNta-qvxMjjdNhnmL5m3c', 'gen': '1pRCqr7-fkwZagi2lBZdTTPTW-sPleLBa3buUK8BRlVI',
    'hyu': '1iqg9AjbPU1oPn-KUFsGUDq5om87-RFXKyO7j9Cl4Ns0'}
def surl(k): return f'https://docs.google.com/spreadsheets/d/{SHEETS[k]}/edit'
MONTHS = {m: i + 1 for i, m in enumerate(['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'])}
FLAGS = []  # (broker, tab, row, message)

# ---------------- strict number parsing ----------------
MONEY_TOKEN = re.compile(r'\$?\s*(\d[\d.,]*)')
WELLFORMED = re.compile(r'^(\d{1,3}(,\d{3})+|\d+)(\.\d{1,2})?$')
def parse_money(v, ctx=None):
    """Return (number|None, problem|None). Numbers pass through. Text must hold one well-formed amount."""
    if v is None: return None, None
    if isinstance(v, (int, float)): return (round(float(v), 2), None)
    s = str(v).strip()
    m = MONEY_TOKEN.search(s)
    if not m: return None, None
    tok = m.group(1).rstrip('.,')
    trailing = m.group(1)[len(tok):]
    if not WELLFORMED.match(tok) or (trailing and s[m.end():m.end() + 1].isdigit()):
        return None, f'malformed amount "{s}" in source sheet (not guessed)'
    # "$2.,245" -> group "2.,245" -> tok "2.,245" fails WELLFORMED -> flagged (old parser read $2)
    return float(tok.replace(',', '')), None

def num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    n, _ = parse_money(v); return n

def k_amount(s):
    """'$3k' -> 3000, '$3,500' -> 3500, '3500' -> 3500"""
    m = re.match(r'\$?\s*([\d,]+(?:\.\d+)?)\s*([kK])?', s.strip())
    if not m: return None
    n = float(m.group(1).replace(',', ''))
    return n * 1000 if m.group(2) else n

def year_of(text):
    m = re.match(r"^\s*'?(20(2[4-9]))\b|^\s*'?(2[4-8])\s", text or '')
    if not m: return None
    return int(m.group(1)) if m.group(1) else 2000 + int(m.group(3))

def strip_year(text):
    return re.sub(r"^\s*'?(20\d\d|2[4-8])\s+", '', text or '').strip()

def iso(y, mo, d):
    try: return datetime.date(y, mo, d).isoformat()
    except ValueError: return None

def md_date(mo, d):
    """month/day stamp with no year: assume 2026 unless that is in the future."""
    y = TODAY.year
    if (mo, d) > (TODAY.month, TODAY.day): y -= 1
    return iso(y, mo, d)

def month_label(text):
    """'NOVEMBER DEALS' / 'November 2026 Ford' / "HN308's NOV 2026 DEALS" -> 'YYYY-MM'"""
    m = re.search(r'\b(january|february|march|april|may|june|july|august|september|sept|october|november|december|'
                  r'jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\b\.?\s*(20\d\d)?', text or '', re.I)
    if not m: return None
    mo = MONTHS[m.group(1).lower()[:3]]
    y = int(m.group(2)) if m.group(2) else (TODAY.year if mo <= TODAY.month else TODAY.year - 1)
    return f'{y}-{mo:02d}'

def money_str(v):
    if v is None: return None
    return str(int(v)) if float(v).is_integer() else str(v)

def cell_str(v):
    if v is None: return None
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    s = str(v).strip()
    return s or None

# ---------------- row factory ----------------
def deal(**kw):
    d = dict(broker=None, lh_username=None, source_tab=None, tab_updated_per_sheet=None, unit_type_status=None, year=None,
             make=None, model=None, trim_description=None, msrp_as_listed=None, msrp_usd=None, msrp_approx=None,
             monthly_payment_usd=None, payment_basis=None, due_at_signing_as_listed=None, das_usd=None, term_mo=None,
             miles_yr=None, region=None, broker_service_fee=None, notes=None, effective_monthly_usd=None,
             effective_monthly_incl_broker_fee_usd=None, pct_of_msrp_payment_msrp=None, spreadsheet_url=None,
             posted_date=None, expires=None, source_url=None)
    fee_num = kw.pop('fee_num', None); das_first_only = kw.pop('das_first_only', False)
    d.update(kw)
    for k in ('model', 'trim_description'):
        if isinstance(d[k], str): d[k] = re.sub(r'\s+', ' ', d[k]).strip() or None
    p, das, t, m = d['monthly_payment_usd'], d['das_usd'], d['term_mo'], d['msrp_usd']
    if p is not None: d['monthly_payment_usd'] = round(p, 2) if not float(p).is_integer() else int(p)
    if p is not None and t:
        if das is not None: d['effective_monthly_usd'] = round(p + das / t, 2)
        elif das_first_only: d['effective_monthly_usd'] = p
        if d['effective_monthly_usd'] is not None and fee_num:
            d['effective_monthly_incl_broker_fee_usd'] = round(d['effective_monthly_usd'] + fee_num / t, 2)
    if p and m: d['pct_of_msrp_payment_msrp'] = round(p / m, 4)
    if d['effective_monthly_usd'] is not None and float(d['effective_monthly_usd']).is_integer():
        d['effective_monthly_usd'] = int(d['effective_monthly_usd'])
    if d['source_url'] is None: d['source_url'] = d['spreadsheet_url']
    return d

def join(*parts):
    out = [p.strip() for p in parts if p and str(p).strip()]
    return '; '.join(out) or None

def rows_of(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]

def find_header(rows, must):
    for i, r in enumerate(rows):
        cells = [str(c).strip().lower() if c is not None else '' for c in r]
        if all(any(re.search(m, c) for c in cells) for m in must): return i, cells
    return None, None

def col(cells, pat):
    for j, c in enumerate(cells):
        if re.search(pat, c): return j
    return None

# ---------------- thread dates ----------------
THREADS = {}
def thread_latest(tids, user):
    """latest post date (UTC->PT date) by `user` among the given thread ids."""
    best = None
    for tid in tids:
        t = THREADS.get(str(tid))
        if not t: continue
        for u, ts, _ in t.get('posts', []):
            if u.lower() == user.lower():
                dt = datetime.datetime.fromisoformat(ts.replace('Z', '+00:00')) - datetime.timedelta(hours=8)
                if best is None or dt > best: best = dt
    return best.date().isoformat() if best else None

# ---------------- NewView ----------------
NEV_MAKE = {'Mercedes-Benz': 'Mercedes-Benz', 'BMW': 'BMW', 'Lexus': 'Lexus', 'Audi': 'Audi', 'Hyundai': 'Hyundai',
            'Porsche': 'Porsche', 'Volkswagen': 'Volkswagen', 'Toyota': 'Toyota', 'Kia': 'Kia', 'Cadillac': 'Cadillac'}
def p_newview(wb):
    out = []
    B = dict(broker='NewView Auto (neviewauto)', lh_username='neviewauto', spreadsheet_url=surl('nev'), payment_basis='Pre-tax')
    for ws in wb.worksheets:
        if ws.title not in NEV_MAKE: continue   # skips 'Copy of Audi' (old 2024/25 blowout copy) + empty SoCal tabs
        rows = rows_of(ws)
        stamp, posted, tabnote = None, None, None
        for r in rows[:6]:
            for c in r:
                s = cell_str(c) or ''
                m = re.match(r'^(\d{1,2})\.(\d{1,2})\s*updated', s, re.I)
                if m: stamp, posted = s, md_date(int(m.group(1)), int(m.group(2)))
                elif s.startswith('*'): tabnote = s.strip()
        hi, cells = find_header(rows, [r'^model$', r'^msrp$'])
        if hi is None: continue
        cm, cmsrp, cterm, cdas = col(cells, r'^model$'), col(cells, r'^msrp$'), col(cells, r'^term'), col(cells, r'driveoff')
        cpay = col(cells, r'monthly')
        region = 'CA (SoCal-based; NorCal & SoCal)'
        for i in range(hi + 1, len(rows)):
            r = rows[i]
            name = cell_str(r[cm]) if cm < len(r) else None
            if name and all(c is None for j, c in enumerate(r) if j != cm):
                if name.lower() in ('norcal', 'socal'): region = name
                continue
            if not name: continue
            pay, prob = parse_money(r[cpay])
            if prob: FLAGS.append((B['broker'], ws.title, i + 1, prob))
            if pay is None and not prob: continue
            termtxt = cell_str(r[cterm]) or ''
            tm = re.match(r'(\d+)\s*m?\s*/\s*([\d.]+)\s*k', termtxt, re.I)
            msrp_l = cell_str(r[cmsrp])
            msrp = approx = None
            if msrp_l:
                mm = re.sub(r'[xX]', '0', msrp_l)
                msrp = k_amount(mm); approx = 'Y (lower bound, "xxx" in sheet)' if re.search('[xX]', msrp_l) else 'N'
            das = num(r[cdas]) if cdas is not None else None
            extras = [cell_str(c) for j, c in enumerate(r) if j > max(cm, cmsrp, cterm, cdas, cpay) and j != 7 and cell_str(c)]
            low = name.lower()
            status = 'Loaner' if 'loaner' in low else 'Demo' if 'demo' in low else 'New'
            out.append(deal(**B, source_tab=ws.title, tab_updated_per_sheet=stamp, unit_type_status=status,
                            year=year_of(name), make=NEV_MAKE[ws.title], model=strip_year(name), msrp_as_listed=msrp_l,
                            msrp_usd=int(msrp) if msrp else None, msrp_approx=approx, monthly_payment_usd=pay,
                            due_at_signing_as_listed=cell_str(r[cdas]) if cdas is not None else None,
                            das_usd=int(das) if das is not None else None,
                            term_mo=int(tm.group(1)) if tm else None, miles_yr=int(float(tm.group(2)) * 1000) if tm else None,
                            region=region if ws.title == 'Mercedes-Benz' else 'CA (SoCal-based; NorCal & SoCal)',
                            notes=join(*extras, tabnote, prob), posted_date=posted))
    return out

# ---------------- PMI ----------------
PMI_TABS = {'Mercedes (SoCal)': 'Mercedes-Benz', 'Mercedes (NorCal)': 'Mercedes-Benz', 'BMW': 'BMW', 'Audi': 'Audi',
            'ToyotaLexus': 'Toyota', 'HyundaiGenesis': 'Hyundai', 'Range Rover': 'Land Rover', 'HondaAcura': 'Honda',
            'Porsche': 'Porsche', 'MazdaSubaru': 'Mazda', 'Chevy': 'Chevrolet', 'Cadillac': 'Cadillac', 'Volkswagen': 'Volkswagen',
            'Ford': 'Ford', 'Kia': 'Kia', 'CDJR': None, 'Volvo': 'Volvo', 'GMC': 'GMC', 'Nissan': 'Nissan'}
PMI_THREADS = {'Mercedes (SoCal)': [559196], 'Mercedes (NorCal)': [559196], 'Range Rover': [559196], 'Porsche': [559196],
               'BMW': [543300], 'Audi': [543300], 'ToyotaLexus': [534682], 'HyundaiGenesis': [534676], 'MazdaSubaru': [612202], 'HondaAcura': [612202],
               'Cadillac': [535347], 'GMC': [535347], 'Chevy': [535347]}
SECTION_MAKE = {'LEXUS': 'Lexus', 'TOYOTA': 'Toyota', 'GENESIS': 'Genesis', 'HYUNDAI': 'Hyundai', 'ACURA': 'Acura',
                'HONDA': 'Honda', 'SUBARU': 'Subaru', 'MAZDA': 'Mazda', 'JAGUAR': 'Jaguar', 'RANGE ROVER': 'Land Rover'}
CDJR_MAKE = [(r'\b(wrangler|grand cherokee|gladiator|compass|cherokee|wagoneer)', 'Jeep'), (r'\bram\b', 'Ram'),
             (r'\b(durango|charger|hornet)', 'Dodge'), (r'\b(pacifica|voyager)', 'Chrysler')]
GEN_MODELS = re.compile(r'^(gv?\d{2}|electrified)', re.I)
def pmi_make(tab, section_make, model):
    if tab == 'CDJR':
        for pat, mk in CDJR_MAKE:
            if re.search(pat, model, re.I): return mk
        return None
    if tab == 'HyundaiGenesis' and GEN_MODELS.match(model): return 'Genesis'
    return section_make or PMI_TABS[tab]

PAY_RE = re.compile(r'^\s*\$?\s*\d[\d.,]*\s*(/\s*mo)?\s*\+?\s*tax', re.I)
TERM_RE = re.compile(r'(\d+)\s*(mo|mos|months?)\b', re.I)
def p_pmi(wb):
    out = []
    B = dict(broker='Platinum Motor Imports (PMI)', lh_username='PlatinumMotorImports', spreadsheet_url=surl('pmi'))
    for ws in wb.worksheets:
        if ws.title not in PMI_TABS: continue
        rows = rows_of(ws)
        hi, cells = find_header(rows, [r'payment'])
        if hi is None: continue
        cpay = col(cells, r'payment')
        cm = 0
        section, section_make, month = None, None, None
        fallback_posted = thread_latest(PMI_THREADS.get(ws.title, []), 'PlatinumMotorImports')
        hdr0 = cell_str(rows[hi][0]) or ''
        if re.search(r'under construction', hdr0, re.I): section = hdr0
        for k, v in SECTION_MAKE.items():
            if hdr0.upper().startswith(k): section_make = v
        for i in range(hi + 1, len(rows)):
            r = rows[i]
            name = cell_str(r[cm])
            texts = [cell_str(c) for c in r]
            payi = next((j for j, s in enumerate(texts) if s and PAY_RE.match(s)), None)
            onepay = next((s for s in texts if s and re.search(r'one\s*pay', s, re.I)), None)
            if name and payi is None and not onepay:
                if re.fullmatch(r'x+', name): continue
                up = name.upper()
                for k, v in SECTION_MAKE.items():
                    if up.startswith(k): section_make = v
                ml = month_label(name)
                if ml and re.search(r'DEALS|SPECIAL', up): month = (name, ml)
                section = name
                continue
            if not name or payi is None:
                if name and onepay:
                    FLAGS.append((B['broker'], ws.title, i + 1, f'one-pay lease "{onepay}" skipped (not a monthly payment)'))
                continue
            paytxt = texts[payi]
            pay, prob = parse_money(paytxt)
            if prob: FLAGS.append((B['broker'], ws.title, i + 1, prob))
            termi = next((j for j, s in enumerate(texts) if s and j != payi and TERM_RE.search(s) and '/' in s), None)
            termtxt = texts[termi] if termi is not None else ''
            tm = TERM_RE.search(termtxt)
            mi = re.search(r'([\d,.]+)\s*(k)?\s*(mi|mile|miles)\b', termtxt, re.I)
            if not mi:  # "36 mo / 7500 / $5000 drives": bare 2nd segment
                segs = [x.strip() for x in termtxt.split('/')]
                mi = re.fullmatch(r'([\d,.]+)\s*(k)?()', segs[1], re.I) if len(segs) > 2 else None
            miles = None
            if mi:
                miles = float(mi.group(1).replace(',', '')) * (1000 if mi.group(2) else 1)
            dm = re.search(r'\$?\s*([\d,]+(?:\.\d+)?\s*k?)\s*(driveoff|drive off|das|drives|due)', termtxt, re.I)
            das = k_amount(dm.group(1)) if dm else None
            loci = next((j for j, s in enumerate(texts) if s and re.fullmatch(r'(so|nor)\s*cal', s.strip(), re.I)), None)
            feei = next((j for j, s in enumerate(texts) if s and re.search(r'\bfee\b', s, re.I) and j not in (payi, termi)), None)
            fee_txt = texts[feei] if feei is not None else None
            fee_num = parse_money(fee_txt)[0] if fee_txt else None
            used = {cm, payi, termi, loci, feei}
            extras = [s for j, s in enumerate(texts) if j not in used and s and not re.fullmatch(r'[A-Z]', s)]
            MS = r'\(\s*\$?\s*([\d.]+)\s*[kK]\s*(?:MSRP)?\s*\)'
            mk = re.search(MS, name)
            msrp_l = mk.group(0).strip('()').replace(' MSRP', '').strip() if mk else None
            msrp = int(round(float(mk.group(1)) * 1000)) if mk else None
            model = re.sub(r'\s+', ' ', strip_year(re.sub(MS + r'|\(\s*[\d.]+\s*\)', '', name))).strip(' -')
            model = re.sub(r'\s+-\s*$', '', model)
            make = pmi_make(ws.title, section_make, model)
            if make and model.lower().startswith(make.lower().split('-')[0] + ' ') and ws.title in ('Chevy',):
                pass
            low = (name + ' ' + (section or '')).lower()
            status = 'Loaner' if 'loaner' in name.lower() else 'Demo' if 'demo' in name.lower() else 'New'
            posted = month[1] if month else fallback_posted
            out.append(deal(**B, source_tab=ws.title, tab_updated_per_sheet=join(section, month[0] if month and month[0] != section else None),
                            unit_type_status=status, year=year_of(name), make=make, model=model, trim_description=name,
                            msrp_as_listed=msrp_l, msrp_usd=msrp, msrp_approx='Y (rounded "k" figure)' if msrp else None,
                            monthly_payment_usd=pay, payment_basis='+ tax' if re.search('tax', paytxt, re.I) else paytxt,
                            due_at_signing_as_listed=termtxt or None, das_usd=int(das) if das is not None else None,
                            term_mo=int(tm.group(1)) if tm else None, miles_yr=int(miles) if miles else None,
                            region=f'CA - {texts[loci]}' if loci is not None else 'CA (SoCal & NorCal)',
                            broker_service_fee=fee_txt, fee_num=fee_num, notes=join(' | '.join(extras) if extras else None, prob),
                            posted_date=posted))
    return out

# ---------------- Lease Wizards ----------------
WIZ_MAKE = {'Acura': 'Acura', 'Mazda': 'Mazda', 'Nissan': 'Nissan', 'Hyundai': 'Hyundai', 'Honda': 'Honda', 'Volkswagen': 'Volkswagen',
            'Infiniti': 'Infiniti', 'Volvo': 'Volvo', 'KIA': 'Kia', 'Audi': 'Audi', 'BMW': 'BMW', 'Toyota': 'Toyota',
            'Mercedes': 'Mercedes-Benz', 'LEXUS': 'Lexus', 'Land Rover': 'Land Rover', 'CJDR': None, 'CPODEMOSPECIALS': None}
MAKE_NORM = {'BENZ': 'Mercedes-Benz', 'MERCEDES': 'Mercedes-Benz', 'VOLKSWAGEN': 'Volkswagen', 'VW': 'Volkswagen', 'BMW': 'BMW',
             'AUDI': 'Audi', 'PORSCHE': 'Porsche', 'TOYOTA': 'Toyota', 'LEXUS': 'Lexus', 'HONDA': 'Honda', 'ACURA': 'Acura',
             'KIA': 'Kia', 'HYUNDAI': 'Hyundai', 'GENESIS': 'Genesis', 'VOLVO': 'Volvo', 'JEEP': 'Jeep', 'RAM': 'Ram',
             'NISSAN': 'Nissan', 'INFINITI': 'Infiniti', 'MAZDA': 'Mazda', 'LAND ROVER': 'Land Rover', 'CADILLAC': 'Cadillac'}
def p_wiz(wb):
    out = []
    B = dict(broker='Lease Wizards (Lease.Wiz)', lh_username='Lease.Wiz', spreadsheet_url=surl('wiz'),
             region='NY, NJ, CT, PA, RI, DE, FL (home delivery)')
    for ws in wb.worksheets:
        if ws.title not in WIZ_MAKE: continue
        rows = rows_of(ws)
        stamp = posted = None
        for r in rows[:6]:
            for c in r:
                m = re.match(r'UPDATED:\s*(\d{1,2})/(\d{1,2})', cell_str(c) or '', re.I)
                if m: stamp, posted = cell_str(c), md_date(int(m.group(1)), int(m.group(2)))
        hi, cells = find_header(rows, [r'^year', r'monthly payment'])
        if hi is None: continue
        cy, cmk, cmo, ctr, cms = col(cells, r'^year'), col(cells, r'^make'), col(cells, r'^model'), col(cells, r'^trim'), col(cells, r'^msrp')
        cte, cmi, cpa, cda, cal = col(cells, r'^term'), col(cells, r'^mileage'), col(cells, r'monthly payment'), col(cells, r'due at signing'), col(cells, r'all in')
        section_make = None
        for i in range(hi + 1, len(rows)):
            r = rows[i] + [None] * 16
            yv = r[cy]
            if not isinstance(yv, (int, float)):
                s = cell_str(yv)
                if s and s.upper() in MAKE_NORM and all(c is None for j, c in enumerate(r) if j != cy): section_make = MAKE_NORM[s.upper()]
                continue
            model = cell_str(r[cmo])
            if not model: continue
            paytxt = r[cpa]
            pay, prob = parse_money(paytxt)
            if prob: FLAGS.append((B['broker'], ws.title, i + 1, prob))
            if pay is None and not prob: continue  # '--' / '-' / 'AVAIL FOR PRE ORDER' => no quoted payment
            make = MAKE_NORM.get((cell_str(r[cmk]) or '').upper()) if cmk is not None else None
            make = make or WIZ_MAKE[ws.title] or section_make
            if ws.title == 'CJDR':
                make = next((mk for pat, mk in CDJR_MAKE if re.search(pat, model, re.I)), None)
            pre = [cell_str(r[j]) for j in range(0, cy) if cell_str(r[j]) and cell_str(r[j]) not in ('l',)]
            status_vals = [p for p in pre if re.fullmatch(r'\*?(SPECIAL|DEMO|CPO|LOANER|NEW)\*?', p, re.I)]
            other_pre = [p for p in pre if p not in status_vals]
            post = [cell_str(r[j]) for j in range((cal if cal is not None else cda) + 1, len(r)) if cell_str(r[j])]
            trim = cell_str(r[ctr]); opt = cell_str(r[ctr + 1]) if cms is not None and ctr + 1 < cms else None
            msrp = num(r[cms])
            allin = num(r[cal]) if cal is not None else None
            out.append(deal(**B, source_tab=ws.title, tab_updated_per_sheet=stamp,
                            unit_type_status=(status_vals[0].strip('*').title() if status_vals else 'New'),
                            year=int(yv), make=make, model=model, trim_description=join(trim, opt),
                            msrp_as_listed=cell_str(r[cms]), msrp_usd=int(msrp) if msrp else None, msrp_approx='N' if msrp else None,
                            monthly_payment_usd=pay, payment_basis='Pre-tax (top-tier credit, all rebates excl. affinity)',
                            due_at_signing_as_listed=cell_str(r[cda]), das_usd=None,
                            term_mo=int(num(r[cte])) if num(r[cte]) else None, miles_yr=int(num(r[cmi])) if num(r[cmi]) else None,
                            notes=join(*other_pre, f'All-in payment (per sheet): ${money_str(allin)}' if allin else None, *post, prob),
                            posted_date=posted))
    return out

# ---------------- HN308 ----------------
HN_TABS = {'New BMW': ('BMW', 'New'), 'Loaner BMW': ('BMW', 'Loaner'), 'Loaner Mercedes': ('Mercedes-Benz', 'Loaner'),
           'New Mercedes': ('Mercedes-Benz', 'New'), 'Cadillac EVs': ('Cadillac', 'New')}
def p_hn(wb):
    out = []
    B = dict(broker='HN308', lh_username='HN308', spreadsheet_url=surl('hn'), region='NY, NJ, CT, PA (limited out-of-state shipping)')
    for ws in wb.worksheets:
        if ws.title not in HN_TABS: continue   # skips 'Sold', ' Loaner cars', Jul-2024 'Lexus'
        make, kind = HN_TABS[ws.title]
        rows = rows_of(ws)
        title = next((cell_str(c) for r in rows[:2] for c in r if cell_str(c) and 'HN308' in cell_str(c)), None)
        posted = month_label(title)
        hi, cells = find_header(rows, [r'^year', r'^model', r'^monthly'])
        if hi is None: continue
        cy, cm, cp, cms = col(cells, r'^year'), col(cells, r'^model'), col(cells, r'^monthly'), col(cells, r'^msrp')
        cdi, cfe, cco, cte = col(cells, r'^discount'), col(cells, r'^broker fee'), col(cells, r'^color'), col(cells, r'^terms')
        cop, cst = col(cells, r'key options'), col(cells, r'state for reg')
        cbens = [(j, c.replace('\n', ' ').strip()) for j, c in enumerate(cells) if re.search(r'loyalty|benefit|conquest', c)]
        links = [j for j, c in enumerate(cells) if 'calc' in c and 'tax and fees upfront' in c] or [j for j, c in enumerate(cells) if 'calc' in c]
        prev_note = None
        for i in range(hi + 1, len(rows)):
            r = rows[i] + [None] * 16
            yv = r[cy]; model = cell_str(r[cm])
            ys = cell_str(yv) or ''
            if not model:
                if ys and not isinstance(yv, (int, float)): prev_note = ys
                continue
            sold = bool(re.match(r'sold', ys, re.I))
            year = int(yv) if isinstance(yv, (int, float)) else None
            pay = num(r[cp])
            if pay is None: continue  # no quoted monthly (custom order / sold-out placeholder)
            link = next((cell_str(r[j]) for j in links if cell_str(r[j]) and 'calculator.leasehackr.com' in cell_str(r[j])), None)
            q = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(link).query)) if link else {}
            das = float(q['lease_das']) if q.get('lease_das') else None
            pretax = float(q['pretax_monPmt']) if q.get('pretax_monPmt') else None
            basis = 'Pre-tax (with inceptions upfront)'
            if pay is not None and pretax is not None and abs(pretax - pay) > 1:
                basis = f'As listed in sheet "Monthly" column (sheet calculator link shows ${money_str(pretax)} pre-tax)'
            tt = cell_str(r[cte]) or ''
            tm = re.match(r'(\d+)\s*/\s*([\d.]+)\s*k?\s*$', tt, re.I)  # '39/10K' or '39/10'
            fee = num(r[cfe]) if cfe is not None else None
            msrp = num(r[cms])
            note = [f'Discount off MSRP: {round(num(r[cdi]) * 100, 1)}%' if cdi is not None and num(r[cdi]) is not None and num(r[cdi]) < 1 else None,
                    f'Color: {cell_str(r[cco])}' if cco is not None and cell_str(r[cco]) else None]
            note += [f'{n.lower()}: {cell_str(r[j])}' for j, n in cbens if cell_str(r[j])]
            if cst is not None and cell_str(r[cst]): note.append(f'state for reg: {cell_str(r[cst])}')
            if cop is not None and cell_str(r[cop]): note.append(f'Options: {cell_str(r[cop])}')
            if ws.title == 'Loaner BMW' and cell_str(r[11]) and not (cell_str(r[11]) or '').startswith('http'): note.append(f'Options: {cell_str(r[11])}')
            if sold: note.append('SOLD (per sheet)' if ys.lower() == 'sold' else f'{ys.upper()} (per sheet)')
            if prev_note: note.append(prev_note)
            out.append(deal(**B, source_tab=ws.title, tab_updated_per_sheet=title,
                            unit_type_status=f'{kind} - ' + ('Sold' if sold else 'Available (listed)') if kind != 'Loaner' or ws.title != 'Loaner BMW' else 'Loaner - Loaner (<15k mi per sheet)' + (' - Sold' if sold else ''),
                            year=year, make=make, model=model.replace('\n', ' '), msrp_as_listed=cell_str(r[cms]),
                            msrp_usd=int(msrp) if msrp else None, msrp_approx='N' if msrp else None, monthly_payment_usd=pay,
                            payment_basis=basis,
                            due_at_signing_as_listed=f"${int(das):,} (from sheet's LH calculator link, tax & fees upfront)" if das is not None else None,
                            das_usd=int(das) if das is not None else None, term_mo=int(tm.group(1)) if tm else None,
                            miles_yr=int(float(tm.group(2)) * 1000) if tm else None,
                            broker_service_fee=f'${int(fee)}' if fee else None, fee_num=fee, notes=join(*note), posted_date=posted))
    return out

# ---------------- Lightning ----------------
def p_lig(wb):
    out = []
    B = dict(broker='Lightning Auto', lh_username='LightningAuto', spreadsheet_url=surl('lig'), region='East Coast (free delivery)',
             broker_service_fee='$499')
    for ws in wb.worksheets:
        rows = rows_of(ws)
        title = next((cell_str(c) for c in rows[0] if cell_str(c)), None)
        posted = month_label(title)
        hi, cells = find_header(rows, [r'model year', r'36m'])
        if hi is None: continue
        cy, cm, ct, cms, cdi, cre = 0, 1, 2, 3, col(cells, 'discount'), col(cells, 'rebates')
        c36, c48, cda, cco, cno = col(cells, '36m'), col(cells, '48m'), col(cells, 'due at signing'), col(cells, '^comment'), col(cells, '^note$')
        tabmake = 'Ford' if 'Ford' in ws.title else 'Mazda' if 'Mazda' in ws.title else None
        cur_y = cur_m = None
        for i in range(hi + 1, len(rows)):
            r = rows[i] + [None] * 16
            if cell_str(r[0]) and cell_str(r[0]).lower().startswith('disclaimer'): break
            if isinstance(r[cy], (int, float)): cur_y = int(r[cy])
            if cell_str(r[cm]): cur_m = cell_str(r[cm])
            trim = cell_str(r[ct])
            if not trim or not cur_m: continue
            make = tabmake or next((mk for pat, mk in CDJR_MAKE if re.search(pat, cur_m, re.I)), None)
            msrp = num(r[cms])
            comment = join(cell_str(r[cco]) if cco is not None else None, cell_str(r[cno]) if cno is not None else None)
            fire = '🔥' in ' '.join(str(c) for c in r if c)
            trim_clean = trim.replace('🔥', '').strip()
            for cc, term in ((c36, 36), (c48, 48)):
                if cc is None: continue
                txt = cell_str(r[cc])
                if not txt: continue
                pay, prob = parse_money(txt)
                if prob: FLAGS.append((B['broker'], ws.title, i + 1, prob))
                if pay is None and not prob: continue
                disc = num(r[cdi]); reb = num(r[cre])
                notes = join(f'Discount {disc * 100:.2f}%' if disc is not None else None,
                             f'rebates ${reb:,.0f}' if reb is not None else None, comment,
                             '🔥 flagged by broker' if fire else None,
                             'Effective = payment (DAS is 1st payment only; excludes tax/TT&L)', prob)
                out.append(deal(**B, source_tab=ws.title, tab_updated_per_sheet=title.strip() if title else None, unit_type_status='New',
                                year=cur_y, make=make, model=cur_m.strip(), trim_description=trim_clean,
                                msrp_as_listed=cell_str(r[cms]), msrp_usd=int(msrp) if msrp else None, msrp_approx='N' if msrp else None,
                                monthly_payment_usd=pay,
                                payment_basis=f'Estimate per sheet disclaimer (sheet column: "{"36M Calculator" if term == 36 else "48M calculator"}")',
                                due_at_signing_as_listed=cell_str(r[cda]), das_usd=None, das_first_only=True, term_mo=term,
                                fee_num=499, notes=notes, posted_date=posted))
    return out

# ---------------- Genesis / Hyundai of Cumming ----------------
def p_cumming(wb, which):
    out = []
    name = 'Genesis of Cumming GA (Aaron, amasih)' if which == 'gen' else 'Hyundai of Cumming GA (Aaron, amasih)'
    make = 'Genesis' if which == 'gen' else 'Hyundai'
    B = dict(broker=name, lh_username='amasih', spreadsheet_url=surl(which), region='Cumming, GA - nationwide shipping',
             broker_service_fee='None (direct dealer)', payment_basis='+ tax',
             due_at_signing_as_listed='1st payment + tax (GA residents); OOS +~$150 + local TT&L')
    for ws in wb.worksheets:
        rows = rows_of(ws)
        if not any(any(c is not None for c in r) for r in rows): continue
        synced = next((cell_str(c) for r in rows[:3] for c in r if cell_str(c) and cell_str(c).startswith('Last synced')), None)
        posted = None
        if synced:
            m = re.search(r'(\w{3}) (\d{1,2}), (\d{4}) (\d{1,2}):(\d{2}) (AM|PM) EST', synced)
            if m:
                h = int(m.group(4)) % 12 + (12 if m.group(6) == 'PM' else 0)
                dt = datetime.datetime(int(m.group(3)), MONTHS[m.group(1).lower()], int(m.group(2)), h, int(m.group(5))) - datetime.timedelta(hours=3)
                posted = dt.date().isoformat()
        posted = posted or month_label(ws.title)
        loaner = 'LOANER' in ws.title.upper()
        defterm = next((cell_str(c) for r in rows[:16] for c in r if cell_str(c) and cell_str(c).startswith('Pricing is for')), None)
        dt_m = re.search(r'(\d+)\s*months?\s*/\s*([\d,]+)\s*miles', defterm or '')
        section = None; cells = None
        for i, r in enumerate(rows):
            r = r + [None] * 12
            c0 = cell_str(r[0])
            if c0 and c0.lower().startswith('stock'):
                cells = [str(c).strip().lower() if c is not None else '' for c in r]; continue
            if c0 and all(c is None for c in r[1:]) and i > 12 and re.fullmatch(r"(20\d\d )?[A-Z0-9 \-]+", c0) and c0 != 'OUT OF STOCK':
                section = c0; cells = None; continue
            if not cells or not c0 or c0 == 'OUT OF STOCK': continue
            g = lambda pat: (lambda j: r[j] if j is not None else None)(col(cells, pat))
            pay = num(g(r'^payment$')); msrp = num(g(r'^msrp$'))
            if pay is None: continue
            termtxt = cell_str(g(r'^term$'))
            tm = re.match(r'(\d+)\s*/\s*([\d.]+)', termtxt or '')
            term = int(tm.group(1)) if tm else (int(dt_m.group(1)) if dt_m else None)
            miles = int(float(tm.group(2)) * 1000) if tm else (int(dt_m.group(2).replace(',', '')) if dt_m else None)
            status_cell = next((cell_str(c) for c in r[9:] if cell_str(c) and re.fullmatch(r'sold|pending', cell_str(c), re.I)), None)
            extras = [f'Stock {c0}', f'{cell_str(g("^exterior"))} / {cell_str(g("^interior"))}']
            for lab in ('comp/loyalty', 'lease cash', 'miles'):
                v = g('^' + re.escape(lab) + '$')
                if v is not None: extras.append(f'{lab}: {v}')
            extras += ['Dealer (no broker fee); $899 doc + $199 e-reg incl.; OOS add ~$149-151 + local TT&L' + (' (term from sheet note: 24 mo/7,500 mi)' if not tm and dt_m else ''),
                       'Effective = payment (DAS is 1st payment only; excludes tax/TT&L)']
            if status_cell: extras.append(f'{status_cell.upper()} (per sheet)')
            out.append(deal(**B, source_tab=f'{make}: {ws.title}', tab_updated_per_sheet=synced or ws.title,
                            unit_type_status=('Loaner' if loaner else 'New') + (' - Sold' if status_cell else ' - Available (listed)'),
                            year=year_of(section), make=make, model=strip_year(section), trim_description=cell_str(g('^model$')),
                            msrp_as_listed=cell_str(g('^msrp$')), msrp_usd=int(msrp) if msrp else None, msrp_approx='N' if msrp else None,
                            monthly_payment_usd=pay, das_usd=None, das_first_only=True, term_mo=term, miles_yr=miles,
                            notes=join(*extras), posted_date=posted))
    return out

PARSERS = [('nev', p_newview), ('pmi', p_pmi), ('wiz', p_wiz), ('hn', p_hn), ('lig', p_lig),
           ('gen', lambda wb: p_cumming(wb, 'gen')), ('hyu', lambda wb: p_cumming(wb, 'hyu'))]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sheets', default='/workspace/sheets_1110'); ap.add_argument('--download', action='store_true')
    ap.add_argument('--threads'); ap.add_argument('--extra')
    ap.add_argument('--xlsx'); ap.add_argument('--out-json', default='all_deals_raw.json')
    a = ap.parse_args()
    if a.threads: THREADS.update(json.load(open(a.threads)))
    os.makedirs(a.sheets, exist_ok=True)
    allrows = []; wbs = {}
    for k, fn in PARSERS:
        p = os.path.join(a.sheets, SHEETS[k] + '.xlsx')
        if a.download:
            subprocess.run(['curl', '-sfL', '-o', p, f'https://docs.google.com/spreadsheets/d/{SHEETS[k]}/export?format=xlsx'], check=True)
        wb = openpyxl.load_workbook(p, data_only=True); wbs[k] = wb
        rows = fn(wb); allrows += rows
        print(f'{k}: {len(rows)} rows', file=sys.stderr)
    if a.extra:
        ex = json.load(open(a.extra)); allrows += [deal(**e) for e in ex]; print(f'extra: {len(ex)} rows', file=sys.stderr)
    json.dump(allrows, open(a.out_json, 'w'), indent=1, ensure_ascii=False)
    for f in FLAGS: print('FLAG', *f, sep=' | ', file=sys.stderr)
    if a.xlsx: write_xlsx(a.xlsx, allrows, wbs)

RAW_TABS = {'nev': 'Raw - neviewauto', 'pmi': 'Raw - PMI', 'wiz': 'Raw - Lease Wizards', 'hn': 'Raw - HN308',
            'lig': 'Raw - Lightning Auto', 'gen': 'Raw - Genesis-Hyundai GA', 'hyu': 'Raw - Genesis-Hyundai GA'}
def write_xlsx(path, rows, wbs):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
    from validate_and_export import COLS, KEYS
    wb = openpyxl.load_workbook(path)
    ws = wb['All Deals']; ws.delete_rows(1, ws.max_row)
    ws.append(COLS)
    for r in rows: ws.append([r[k] for k in KEYS])
    ws.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(COLS))}{ws.max_row}'
    done = set()
    for k, tab in RAW_TABS.items():
        if tab not in done and tab in wb.sheetnames:
            wsr = wb[tab]; wsr.delete_rows(1, wsr.max_row); wsr.append(['Source Tab', 'Row #', 'Cells (as exported)']); done.add(tab)
        wsr = wb[tab]
        for s in wbs[k].worksheets:
            for i, r in enumerate(s.iter_rows(values_only=True), 1):
                if any(c is not None for c in r):
                    wsr.append([s.title, i] + [c if not isinstance(c, str) else c[:500] for c in r[:30]])
    wb.save(path)

if __name__ == '__main__':
    main()
