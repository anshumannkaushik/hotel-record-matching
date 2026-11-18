import openpyxl, re, json
from difflib import SequenceMatcher
from collections import Counter

def norm(s):
    s = (s or '').lower()
    s = re.sub(r'[^a-z0-9 ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def toks(s): return norm(s).split()

def name_score(a, b):
    ta, tb = toks(a), toks(b)
    if not ta or not tb: return 0.0
    sa, sb = set(ta), set(tb)
    jacc = len(sa & sb) / max(1, len(sa | sb))
    seq = SequenceMatcher(None, ' '.join(ta), ' '.join(tb)).ratio()
    sc = 0.5 * jacc + 0.5 * seq
    # substring boost: shorter name fully inside longer
    if (sa <= sb or sb <= sa):
        sc = max(sc, 0.85)
    return sc

def addr_score(a, b):
    ta, tb = toks(a), toks(b)
    if not ta or not tb: return 0.0
    sa, sb = set(ta), set(tb)
    jacc = len(sa & sb) / max(1, len(sa | sb))
    seq = SequenceMatcher(None, ' '.join(ta), ' '.join(tb)).ratio()
    return 0.6 * jacc + 0.4 * seq

def postals(s):
    return set(re.findall(r'\b\d{4,6}\b', (s or '').lower()))

xp = openpyxl.load_workbook('your_files/workbook2-3_expedia_verified.xlsx', data_only=True)
bk = openpyxl.load_workbook('your_files/workbook2-3_booking_verified.xlsx', data_only=True)
wsx, wsb = xp.active, bk.active
hx = [c.value for c in wsx[1]]
hb = [c.value for c in wsb[1]]

out = []
for i in range(2, 932):
    x = dict(zip(hx, [c.value for c in wsx[i]]))
    if x.get('verification') != 'Verified':
        continue
    b = dict(zip(hb, [c.value for c in wsb[i]]))
    bv = b.get('verification')
    xn, xa = x.get('expedia_listing_name') or '', x.get('expedia_listing_address') or ''
    bn, ba = b.get('booking_listing_name') or '', b.get('booking_listing_address') or ''
    src = ' '.join(str(x.get(k) or '') for k in ('name','address1','address2','postal_code'))
    if bv == 'Verified':
        ns, sa_ = name_score(xn, bn), addr_score(xa, ba)
        sp, lp = postals(xa), postals(ba)
        pc_bad = bool(sp and lp and not (sp & lp))
        tag = f'name {ns:.2f}, addr {sa_:.2f}'
        if not pc_bad and ((ns >= 0.80 and sa_ >= 0.30) or (ns >= 0.60 and sa_ >= 0.40)):
            verdict, note = 'Same', tag
        elif ns >= 0.45 or (sa_ >= 0.45 and not pc_bad):
            verdict, note = 'Review', 'partial agreement: ' + tag
        else:
            verdict, note = 'Bad Match', 'listings differ: ' + tag
    elif bv == 'Mismatch':
        verdict, note = 'Bad Match', 'booking audit found a different property'
    elif bv == 'Review':
        verdict, note = 'Review', 'booking audit uncertain'
    elif bv == 'Dead link':
        verdict, note = 'Review', 'booking link dead; expedia only'
    elif bv == 'Not taking reservations':
        verdict, note = 'Review', 'booking not taking reservations; expedia only'
    else:
        verdict, note = 'Review', 'no booking candidate; expedia only'
    out.append(dict(row=i-1, src=src, xp_name=xn, xp_addr=xa, bk_verdict=bv, bk_name=bn, bk_addr=ba, verdict=verdict, note=note))

c = Counter(r['verdict'] for r in out)
print('rows:', len(out), dict(c))
print('--- Bad Match ---')
for r in out:
    if r['verdict'] == 'Bad Match':
        print('row', r['row'], '|', r['note'])
        print('  XP:', r['xp_name'][:65], '|', r['xp_addr'][:60])
        print('  BK:', r['bk_name'][:65], '|', r['bk_addr'][:60])
json.dump(out, open('hidden_xp/verified_compare.json','w'), ensure_ascii=False)
