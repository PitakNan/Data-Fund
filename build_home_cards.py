# -*- coding: utf-8 -*-
"""สร้าง _home_cards.json สรุปการ์ดหน้า Home จากไฟล์ JSON รายประเด็น
รันหลังอัพเดต JSON ประเด็นใดก็ได้: python build_home_cards.py
(อ่านเฉพาะไฟล์ในโฟลเดอร์นี้ ไม่แตะ Excel — ~1 วินาที)"""
import json, os
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
def load(name):
    with open(os.path.join(BASE, name), encoding='utf-8') as f:
        return json.load(f)

TH_M = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
def ym_th(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    return f"{TH_M[m-1]}{(y+543) % 100:02d}"

def fy_of_ym(ym):
    """ปีงบไทยจาก YYYY-MM (ต.ค.–ก.ย.)"""
    y, m = int(ym[:4]), int(ym[5:7])
    fy = y + 543 + (1 if m >= 10 else 0)
    return str(fy % 100)

issues = {}

# ── NB <1,500g ── rows: [pIdx, yr, hcode, pdx, baht, adjrw, mi]
nb = load('_nb_data.json')
yearly, f23 = {}, 0
for r in nb['rows']:
    e = yearly.setdefault(r[1], {'c': 0, 'b': 0.0})
    e['c'] += 1; e['b'] += r[4]
    if nb['hosp_meta'].get(r[2], {}).get('l') in ('F2', 'F3'):
        f23 += 1
tot = sum(e['c'] for e in yearly.values()) or 1
issues['nb'] = {
    'asof': 'ข้อมูลถึง ' + ym_th(nb['months'][-1]),
    'yearly': [{'yr': y, 'c': v['c'], 'b': round(v['b'], 2)} for y, v in sorted(yearly.items())],
    'extra': {'f23': f23, 'f23_pct': round(f23 / tot * 100, 1)},
}

# ── จิตเวช ── rows: [pIdx, yr, hcode, pdx, baht, adjrw, re, mi]
ps = load('_psych_data.json')
yearly, re_n = {}, 0
for r in ps['rows']:
    e = yearly.setdefault(r[1], {'c': 0, 'b': 0.0})
    e['c'] += 1; e['b'] += r[4]; re_n += r[6]
tot = sum(e['c'] for e in yearly.values()) or 1
issues['psych'] = {
    'asof': 'ข้อมูลถึง ' + ym_th(ps['months'][-1]),
    'yearly': [{'yr': y, 'c': v['c'], 'b': round(v['b'], 2)} for y, v in sorted(yearly.items())],
    'extra': {'re_pct': round(re_n / tot * 100, 1)},
}

# ── ส่งต่อออกนอก ── rows: [srcIdx, destIdx, yr, pdx, baht, adjrw, mi]
rf = load('_refer_data.json')
yearly, dcnt = {}, {}
for r in rf['rows']:
    e = yearly.setdefault(r[2], {'c': 0, 'b': 0.0})
    e['c'] += 1; e['b'] += r[4]
    dcnt[r[1]] = dcnt.get(r[1], 0) + 1
tot = sum(e['c'] for e in yearly.values()) or 1
top_di = max(dcnt, key=dcnt.get)
issues['refer'] = {
    'asof': 'ข้อมูลถึง ' + ym_th(rf['months'][-1]),
    'yearly': [{'yr': y, 'c': v['c'], 'b': round(v['b'], 2)} for y, v in sorted(yearly.items())],
    'extra': {'top_dest': rf['dests'][top_di]['n'], 'top_dest_pct': round(dcnt[top_di] / tot * 100, 1)},
}

# ── adjRW<0.5 ── monthly: [{ym, cases, baht, ...}]; top hosp จาก provinces[].hospitals
ar = load('_adjrw_los2.json')
yearly = {}
for mrow in ar['monthly']:
    fy = fy_of_ym(mrow['ym'])
    e = yearly.setdefault(fy, {'c': 0, 'b': 0.0})
    e['c'] += mrow.get('cases', 0); e['b'] += mrow.get('baht', 0)
hosps = [h for p in ar.get('provinces', []) for h in p.get('hospitals', [])]
top_h = max(hosps, key=lambda h: h.get('cases', 0)) if hosps else {}
issues['adjlow'] = {
    'asof': 'ข้อมูลถึง ' + ym_th(ar['monthly'][-1]['ym']),
    'yearly': [{'yr': y, 'c': v['c'], 'b': round(v['b'], 2)} for y, v in sorted(yearly.items())],
    'extra': {'top_hosp': top_h.get('name', '–'), 'top_hosp_cases': top_h.get('cases', 0)},
}

# ── PPFS ── svc_yr index si*3+yi = [persons, visits, amount]; years ['2567','2568','2569']
pp = load('_ppfs_data.json')
yrs = pp['years']
yearly = {}
for si in range(len(pp['svcs'])):
    for yi, y in enumerate(yrs):
        d = pp['svc_yr'][si * 3 + yi]
        e = yearly.setdefault(y[-2:], {'c': 0, 'b': 0.0, 'p': 0})
        e['p'] += d[0]; e['c'] += d[1]; e['b'] += d[2]
b68 = yearly.get('68', {}).get('b', 0); b69 = yearly.get('69', {}).get('b', 0)
issues['ppfs'] = {
    'asof': 'ข้อมูลถึง ' + pp.get('meta', {}).get('as_of', 'ล่าสุด'),
    'yearly': [{'yr': y, 'c': v['c'], 'b': round(v['b'], 2)} for y, v in sorted(yearly.items())],
    'extra': {'p69': yearly.get('69', {}).get('p', 0),
              'ch_pct': round((b69 - b68) / b68 * 100, 1) if b68 else 0},
}

out = {'generated': date.today().isoformat(), 'issues': issues}
path = os.path.join(BASE, '_home_cards.json')
with open(path, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
print('WROTE', path, round(os.path.getsize(path) / 1024, 1), 'KB')
for k, v in issues.items():
    ys = {e['yr']: e['c'] for e in v['yearly']}
    print(f"  {k}: {ys} | {v['asof']} | {v['extra']}")
