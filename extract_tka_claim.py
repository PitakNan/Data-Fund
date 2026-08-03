# -*- coding: utf-8 -*-
"""สร้างข้อมูลฐาน 'ส่งเคลม' ตามรอบปีงบเคลมจริง 16 ก.ย.68 – 15 ก.ย.69
+ ข้อมูลฐาน 'ผ่าตัดจริง' (วันจำหน่าย) ของปีงบ 2569 สำหรับกราฟเทียบ Gap

สคริปต์นี้คือตัวสร้างค่าใน const `TKA_CLAIM` ของ portal_v2.html (ฐานส่งเคลม)
คู่กับ extract_tka.py ที่สร้างฐาน 'ผ่าตัด' (วันจำหน่าย) สำหรับเทรนด์ 5 ปี

── วิธีรีเฟรชเมื่อได้ไฟล์ CSV งวดใหม่จาก สปสช. ─────────────────────────
1. แก้ตัวแปร CSV ด้านล่างให้ชี้ไฟล์ใหม่
2. รัน  `python extract_tka_claim.py`  (ต้องตั้ง PYTHONIOENCODING=utf-8 บน Windows)
3. เอาค่าจาก claim_data.json ไปแทนใน const TKA_CLAIM ของ portal_v2.html
4. ⚠️ อย่าลืมอัปเดต 3 ค่าที่ยัง hardcode ด้วยมือใน TKA_CLAIM:
   - `cut`     = วันที่ตัดข้อมูล (สคริปต์พิมพ์ให้ตอนรัน = data cut)
   - `elapsed` = จำนวนเดือนที่ผ่านไปในรอบเคลม นับจาก 16 ก.ย. (ใช้คิด "จังหวะที่ควรเป็น"
                 ทั้งแถบความก้าวหน้าและข้อความบนการ์ด Home) — ถ้าลืม ตัวเลข % จะเพี้ยน
   - `win`     = ช่วงรอบปีงบเคลม (เปลี่ยนเมื่อขึ้นปีงบใหม่)
5. อัปเดตป้ายวันที่ในหน้า: "เคลมถึง …" (hdr-meta + tkaCard.asof)
"""
import csv, json, os
from collections import defaultdict, Counter
from datetime import datetime

TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
PROVS = ['เชียงราย','เชียงใหม่','ลำปาง','น่าน','ลำพูน','พะเยา','แพร่','แม่ฮ่องสอน']
CSV = r"D:\OneDrive\Share Rh1-New\8. ประชุม\คณะทำงาน M&E\2569\2026-06-26 ข้อมูล TKA ย้อนหลัง 5 ปี จาก สปสช\detail_data_TKA(250626).csv"
(FY,GB,HCPROV,HN,HC,_f,PID,TR,_i,AGEMAIN,_k,SEX,ADM,DIS,SEND,PDX,_o,_p,REV,ITEM,HMAIN,PROV,MNAME,MVAL)=range(24)
# FY(0)='ปีงบฯ บริการ' (= ปีงบของวันจำหน่าย ตรวจแล้วตรง 100%) · GB(1)='ปีงบฯส่งข้อมูล GB' (= ปีงบที่เคลม ตัวเลขทางการ)
# ⚠️ คนละคอลัมน์กัน อย่าสลับ: PROV(21)='จังหวัด hmain' = จังหวัดหน่วยบริการประจำของผู้ป่วย
#    (dashboard ใช้ตัวนี้ทุกที่) · HCPROV(2)='จังหวัด hcode' = จังหวัดของ รพ. ที่ผ่าตัด/ส่งเคลมจริง
#    ใช้ HCPROV เฉพาะสถิติ lag เพราะเป็นเรื่องพฤติกรรมการส่งเคลมของ รพ. และทำให้ยอด รพ. รวมได้ตรงจังหวัดพอดี

# ชื่อ รพ. — ในไฟล์ CSV บางแห่งคอลัมน์ 'หน่วยบริการ' ว่าง และบางแห่งชื่อไม่ตรงกับ key ที่ dashboard ใช้
# (ดู HCODE_NAME_MAP ใน extract_tka.py — ต้องตรงกันเสมอ ห้าม derive จากชื่อในไฟล์อย่างเดียว)
HNAME = {
    '10672':'รพ.ลำปาง','10674':'รพ.เชียงรายประชานุเคราะห์','10713':'รพ.นครพิงค์','10714':'รพ.ลำพูน',
    '10715':'รพ.แพร่','10716':'รพ.น่าน','10717':'รพ.พะเยา','10718':'รพ.เชียงคำ','10719':'รพ.ศรีสังวาลย์',
    '11119':'รพ.จอมทอง','11125':'รพ.ฝาง','11128':'รพ.สันป่าตอง','11130':'รพ.สันทราย','11169':'รพ.สูงเม่น',
    '11177':'รพ.เวียงสา','11190':'รพ.พาน','11192':'รพ.แม่จัน','11194':'รพ.แม่สาย','11453':'รพร.ปัว',
    '11512':'รพ.ค่ายสุรศักดิ์มนตรี','13780':'รพ.มหาราชนครเชียงใหม่ มช.','14550':'รพ.เชียงใหม่ใกล้หมอ',
    '14555':'รพ.ศิริเวช ลำพูน','41347':'รพ.มหาวิทยาลัยพะเยา','41509':'รพ.ศูนย์การแพทย์ ม.แม่ฟ้าหลวง',
    '42186':'รพ.เอกชน (Hcode 42186)',
}

ps = lambda s: datetime.strptime(s, '%d %B %Y') if s else None
def pd_(s):
    try: return datetime(int(s[:4])-543, int(s[4:6]), int(s[6:8]))
    except Exception: return None

with open(CSV, encoding='utf-8-sig', newline='') as f:
    r = csv.reader(f); next(r); rows = list(r)
cs = {}
for row in rows:
    if row[HMAIN] != '01': continue
    k = (row[PID], row[HC], row[ADM])
    c = cs.get(k)
    if c is None:
        try: sd = ps(row[SEND])
        except Exception: sd = None
        c = {'send': sd, 'disch': pd_(row[DIS]), 'fy': row[FY], 'gb': row[GB], 'prov': row[PROV], 'baht': 0.0,
             'hc': row[HC], 'hcprov': row[HCPROV], 'req': 0.0, 'paid': 0.0,
             'hname': HNAME.get(row[HC]) or row[HN].strip().replace('รพ. ', 'รพ.') or f'Hcode {row[HC]}'}
        cs[k] = c
    try: _v = float(row[MVAL])
    except ValueError: _v = 0.0
    # 'จำนวนขอเบิก Inst.' vs 'จำนวนจ่าย Inst.' = จำนวนอุปกรณ์ที่ขอเบิก vs ที่ สปสช. จ่ายจริง
    # ใช้ตอบคำถาม "เคลมช้า/ข้ามปีงบ แล้วได้เงินหรือไม่"
    if row[MNAME] == 'ยอดชดเชย(บาท)': c['baht'] += _v
    elif row[MNAME] == 'จำนวนขอเบิก Inst.': c['req'] += _v
    elif row[MNAME] == 'จำนวนจ่าย Inst.': c['paid'] += _v

# ══ รอบปีงบเคลม 2569 = 16 ก.ย.2568 – 15 ก.ย.2569 ══
CW0 = datetime(2025, 9, 16)
CW1 = datetime(2026, 9, 15, 23, 59, 59)
CUT = max(c['send'] for c in cs.values() if c['send'])

claim = [c for c in cs.values() if c['send'] and CW0 <= c['send'] <= CW1]
print(f"data cut = {CUT.date()}   รอบเคลมปีงบ 2569 = {CW0.date()} ถึง {CW1.date()}")
print(f"\nเคลมที่ส่งในรอบปีงบ 2569 : {len(claim):,} เคส  {sum(c['baht'] for c in claim)/1e6:.1f} ลบ.")
print("  แยกตามปีงบที่ผ่าตัด:", dict(sorted(Counter(c['fy'] for c in claim).items())))
print(f"\nโควตาเขต 3,200 → คงเหลือ {3200-len(claim):,} เข่า")

# ── รายเดือน (ตามวันส่งเคลม)
cm = defaultdict(lambda: [0, 0.0])
for c in claim:
    cm[(c['send'].year, c['send'].month)][0] += 1
    cm[(c['send'].year, c['send'].month)][1] += c['baht']
mk = sorted(cm)
claim_months = [f"{TH[m-1]}{(y+543)%100}" for y, m in mk]
claim_cases = [cm[k][0] for k in mk]
claim_baht = [round(cm[k][1]) for k in mk]
print("\nรายเดือน (ส่งเคลม):")
run = 0
for lb, n, b in zip(claim_months, claim_cases, claim_baht):
    run += n
    print(f"  {lb:<8}{n:>6,} เคส  {b/1e6:>6.1f} ลบ.   สะสม {run:>6,}  เหลือ {3200-run:>6,}")

# ── รายจังหวัด
cp = defaultdict(lambda: [0, 0.0, 0])
for c in claim:
    cp[c['prov']][0] += 1; cp[c['prov']][1] += c['baht']
    if c['fy'] == '2569': cp[c['prov']][2] += 1

# ── โควตาจัดสรร 1 (สัดส่วน FY68 ของ ME)
FY68 = {'น่าน':119,'พะเยา':109,'ลำปาง':136,'ลำพูน':164,'เชียงราย':305,'เชียงใหม่':250,'แพร่':89,'แม่ฮ่องสอน':9}
tot68 = sum(FY68.values())
alloc = {p: round(3200 * FY68[p] / tot68) for p in PROVS}
diff = 3200 - sum(alloc.values())
if diff:
    mx = max(alloc, key=lambda p: alloc[p]); alloc[mx] += diff
print(f"\nโควตาจัดสรร 1 รวม = {sum(alloc.values()):,}")

print("\n" + "═"*84)
print("ตารางโควตารายจังหวัด (จัดสรร 1) vs เคลมที่ส่งแล้ว")
print("═"*84)
print(f"{'จังหวัด':<14}{'โควตาได้รับ':>12}{'ส่งเคลมแล้ว':>13}{'คงเหลือ':>10}{'ใช้ไป%':>9}{'ลบ.':>9}")
print("-"*84)
prov_rows = []
for p in PROVS:
    a = alloc[p]; n, b, n69 = cp.get(p, [0, 0.0, 0])
    prov_rows.append({'prov': p, 'alloc': a, 'claimed': n, 'left': a - n,
                      'pct': round(n / a * 100, 1), 'baht': round(b), 'svc69': n69})
    print(f"{p:<14}{a:>12,}{n:>13,}{a-n:>10,}{n/a*100:>8.1f}%{b/1e6:>9.1f}")
print("-"*84)
ta, tc, tb = sum(alloc.values()), len(claim), sum(c['baht'] for c in claim)
print(f"{'รวมเขต 1':<14}{ta:>12,}{tc:>13,}{ta-tc:>10,}{tc/ta*100:>8.1f}%{tb/1e6:>9.1f}")

# ── กราฟ Gap: ปีงบบริการ 2569 — ผ่าตัด(จำหน่าย) vs ส่งเคลม รายเดือน
svc69 = [c for c in cs.values() if c['fy'] == '2569']
dm = defaultdict(int); sm = defaultdict(int)
for c in svc69:
    if c['disch']: dm[(c['disch'].year, c['disch'].month)] += 1
    if c['send']: sm[(c['send'].year, c['send'].month)] += 1
gk = sorted(set(dm) | set(sm))
gap_months = [f"{TH[m-1]}{(y+543)%100}" for y, m in gk]
gap_disch = [dm.get(k, 0) for k in gk]
gap_claim = [sm.get(k, 0) for k in gk]
print("\n" + "═"*72)
print("Gap: เคสปีงบบริการ 2569 — ผ่าตัด(จำหน่าย) vs ส่งเคลม")
print("═"*72)
print(f"{'เดือน':<10}{'ผ่าตัด':>9}{'ส่งเคลม':>10}{'สะสมผ่าตัด':>13}{'สะสมเคลม':>12}{'ค้างเคลม':>11}")
print("-"*72)
rd = rc = 0
for lb, d, s in zip(gap_months, gap_disch, gap_claim):
    rd += d; rc += s
    print(f"{lb:<10}{d:>9,}{s:>10,}{rd:>13,}{rc:>12,}{rd-rc:>11,}")

# ══════════════════════════════════════════════════════════════════════
# ⏱️ ระยะเวลา จำหน่าย → ส่งเคลม (lag) — เขต / จังหวัด / รพ.
# ══════════════════════════════════════════════════════════════════════
# กลุ่มตัวอย่าง = เคสที่ "ส่งเคลมในรอบปีงบ 2569" (ชุดเดียวกับ TKA_CLAIM.total)
# รวมเคสผ่าตัดปีงบ 68 ที่เพิ่งส่งเคลมปีนี้ด้วย เพราะเป็นตัวที่ทำให้ค่าเฉลี่ยยืดจริง
def stats(v):
    v = sorted(v); n = len(v)
    q = lambda p: v[min(n-1, int(round(p*(n-1))))]
    return {'n': n, 'mean': round(sum(v)/n, 1), 'med': q(.5), 'p90': q(.9), 'max': v[-1],
            'w30': round(sum(1 for x in v if x <= 30)/n*100, 1)}

lag_cases = [(c, (c['send']-c['disch']).days) for c in claim if c['disch']]
bad = [d for _, d in lag_cases if d < 0]
print(f"\n{'═'*72}\n⏱️  ระยะเวลา จำหน่าย → ส่งเคลม\n{'═'*72}")
print(f"เคสที่คำนวณได้ {len(lag_cases):,}/{len(claim):,} · lag ติดลบ {len(bad)} เคส (ตัดออก)")
lag_cases = [(c, d) for c, d in lag_cases if d >= 0]

lag_all = stats([d for _, d in lag_cases])
print(f"ทั้งเขต: เฉลี่ย {lag_all['mean']} วัน · มัธยฐาน {lag_all['med']} · p90 {lag_all['p90']} · "
      f"ช้าสุด {lag_all['max']} · ส่งใน 30 วัน {lag_all['w30']}%")
for fy in sorted({c['fy'] for c, _ in lag_cases}):
    s = stats([d for c, d in lag_cases if c['fy'] == fy])
    print(f"   ปีงบบริการ {fy}: n={s['n']:,} เฉลี่ย {s['mean']} มัธยฐาน {s['med']} ช้าสุด {s['max']}")

byp, byh = defaultdict(list), defaultdict(list)
hmeta = {}
for c, d in lag_cases:
    byp[c['hcprov']].append(d)
    byh[c['hc']].append(d)
    hmeta[c['hc']] = (c['hname'], c['hcprov'])
lag_prov = sorted(({'p': p, **stats(v)} for p, v in byp.items()), key=lambda r: -r['mean'])
lag_hosp = sorted(({'h': hmeta[k][0], 'p': hmeta[k][1], 'hc': k, **stats(v)} for k, v in byh.items()),
                  key=lambda r: -r['mean'])
print(f"\n{'จังหวัด (ตาม รพ. ที่ส่งเคลม)':<26}{'n':>6}{'เฉลี่ย':>9}{'มัธยฐาน':>10}{'p90':>6}{'≤30 วัน':>10}")
for r in lag_prov:
    print(f"  {r['p']:<24}{r['n']:>6,}{r['mean']:>9.1f}{r['med']:>10}{r['p90']:>6}{r['w30']:>9.1f}%")
print(f"\n{'หน่วยบริการ':<36}{'จังหวัด':<12}{'n':>6}{'เฉลี่ย':>9}{'มัธยฐาน':>10}{'p90':>6}")
for r in lag_hosp:
    print(f"  {r['h']:<34.34}{r['p']:<12}{r['n']:>6,}{r['mean']:>9.1f}{r['med']:>10}{r['p90']:>6}")

# ── ③ เคสข้ามปีงบในรอบเคลมนี้ (ผ่าตัดปีงบก่อน แต่เพิ่งส่งเคลมรอบ 2569)
# แยก "ชนเส้นแบ่งรอบ" ออกจาก "ค้างจริง": เคสที่จำหน่ายตั้งแต่ 1 ส.ค.68 คือเคสปลายปีงบ 68
# ที่ส่งเคลมตามจังหวะปกติ แต่วันส่งไปตกหลังเส้น 16 ก.ย. → ไม่ใช่การส่งช้า
EDGE = datetime(2025, 8, 1)
cross = [(c, d) for c, d in lag_cases if c['fy'] != '2569']
edge = [(c, d) for c, d in cross if c['disch'] >= EDGE]
stuck = [(c, d) for c, d in cross if c['disch'] < EDGE]
paid_pct = lambda g: round(sum(1 for c, _ in g if c['req'] > 0 and c['paid'] >= c['req'])/max(len(g),1)*100, 1)
grp = lambda g: {'n': len(g), 'baht': round(sum(c['baht'] for c, _ in g)),
                 'mean': round(sum(d for _, d in g)/max(len(g),1), 1), 'paidPct': paid_pct(g)}
cross_out = {'all': grp(cross), 'edge': grp(edge), 'stuck': grp(stuck),
             'bucket': [{'lbl': lb, 'n': sum(1 for _, d in cross if lo <= d <= hi)}
                        for lb, lo, hi in [('≤30 วัน',0,30),('31–45',31,45),('46–90',46,90),('91–180',91,180),('>180',181,99999)]],
             'topHosp': sorted(({'h': c['hname'], 'n': sum(1 for x, _ in stuck if x['hc'] == c['hc']),
                                 'mean': round(sum(d for x, d in stuck if x['hc'] == c['hc'])/max(sum(1 for x,_ in stuck if x['hc']==c['hc']),1))}
                                for c in {x['hc']: x for x, _ in stuck}.values()), key=lambda r: -r['n'])}
print(f"\n{'═'*72}\n③ เคสข้ามปีงบในรอบเคลม 2569\n{'═'*72}")
print(f"  ทั้งหมด {cross_out['all']['n']} เคส {cross_out['all']['baht']/1e6:.1f} ลบ. · จ่ายครบ {cross_out['all']['paidPct']}%")
print(f"  ชนเส้นแบ่งรอบ (จำหน่าย >= 1 ส.ค.68): {cross_out['edge']['n']} เคส เฉลี่ย {cross_out['edge']['mean']} วัน")
print(f"  ค้างจริง (จำหน่ายก่อน ส.ค.68)      : {cross_out['stuck']['n']} เคส เฉลี่ย {cross_out['stuck']['mean']} วัน "
      f"{cross_out['stuck']['baht']/1e6:.1f} ลบ.")
for r in cross_out['topHosp']: print(f"      {r['h']:<34}{r['n']:>4} เคส  เฉลี่ย {r['mean']:>4} วัน")

# ── ① ข้อมูลประกอบ: ทำไมเทียบ lag ย้อนหลัง 5 ปีตรงๆ ไม่ได้
# ปีเก่า "วันที่ส่งข้อมูล" จำนวนมากเป็นการส่งเป็นก้อน (batch) รอบปิดปีงบ ไม่ใช่วันส่งจริงรายเคส
allc = [(c, (c['send']-c['disch']).days) for c in cs.values() if c['send'] and c['disch']]
allc = [(c, d) for c, d in allc if d >= 0]
byday = defaultdict(list)
for c, d in allc: byday[c['send'].date()].append(d)
BIG = 300
bigdays = {k for k, v in byday.items() if len(v) >= BIG}
inbig = [(c, d) for c, d in allc if c['send'].date() in bigdays]
normal = [(c, d) for c, d in allc if c['send'].date() not in bigdays]
topday = max(byday.items(), key=lambda kv: len(kv[1]))
topday_cases = [c for c, _ in allc if c['send'].date() == topday[0]]
batch = {'days': len(bigdays), 'sendDays': len(byday), 'nAll': len(allc),
         'nBig': len(inbig), 'pctBig': round(len(inbig)/len(allc)*100, 1),
         'topDay': str(topday[0]), 'topN': len(topday[1]),
         'topFrom': str(min(c['disch'] for c in topday_cases).date()),
         'topTo': str(max(c['disch'] for c in topday_cases).date()),
         'bigMean': round(sum(d for _, d in inbig)/len(inbig), 1),
         'normMean': round(sum(d for _, d in normal)/len(normal), 1),
         'normMed': sorted(d for _, d in normal)[len(normal)//2],
         'curMaxDay': max(Counter(c['send'].date() for c in claim).values())}
print(f"\n{'═'*72}\n① เช็ค batch: ทำไมเทียบย้อนหลัง 5 ปีตรงๆ ไม่ได้\n{'═'*72}")
print(f"  เคสทั้ง 5 ปี {batch['nAll']:,} · วันที่มีการส่ง {batch['sendDays']:,} วัน")
print(f"  วันที่ส่งเกิน {BIG} เคส/วัน มี {batch['days']} วัน = {batch['nBig']:,} เคส ({batch['pctBig']}% ของทั้งหมด)")
print(f"  วันเดียวมากสุด {batch['topDay']} = {batch['topN']:,} เคส (จำหน่าย {batch['topFrom']} ถึง {batch['topTo']})")
print(f"  lag เฉลี่ย: ในก้อน {batch['bigMean']} วัน · ส่งตามปกติ {batch['normMean']} (มัธยฐาน {batch['normMed']})")
print(f"  รอบปีงบ 2569 วันเดียวมากสุด {batch['curMaxDay']} เคส -> ไม่มีก้อนปน ✅")

# ── ④ ตารางไขว้ ปีที่จำหน่าย × ปีที่เคลม รายหน่วยบริการ (ตามที่ CFO ขอ 2026-08-03)
#    แถว  = ปีงบที่จำหน่าย → ใช้คอลัมน์ 'ปีงบฯ บริการ' ของ สปสช. โดยตรง (ตรวจแล้วตรงกับปีงบของวันจำหน่าย 100%)
#    คอลัมน์ = ปีงบที่ส่งเคลม → ใช้คอลัมน์ 'ปีงบฯส่งข้อมูล GB' ของ สปสช. (ตัวเลขทางการ)
#    ⚠️ กฎ 16 ก.ย. ที่ dashboard ใช้ตรงกับ GB 99.05% — ต่างกัน 125 เคสที่ส่ง 16–25 ก.ย.2565
#       ซึ่ง สปสช. นับเป็นปีงบ 65 (ปีนั้นใช้เส้นตัดคนละวัน) · ปีงบ 67/68/69 ตรงกันเป๊ะทุกเคส
gb_num = lambda g: int(g.replace('ปีงบฯ ', '')) + 2500
mx_cases = [c for c in cs.values() if c['send']]
DY = sorted({int(c['fy']) for c in mx_cases})
CY = sorted({gb_num(c['gb']) for c in mx_cases})
def mx_rows(pool):
    """แต่ละแถวเก็บ mn/mx/sum ของระยะเวลา จำหน่าย→ส่งเคลม (วัน) ด้วย
       เก็บ sum (ไม่ใช่ avg) เพื่อให้ฝั่งหน้าเว็บรวมข้ามแถว/ข้าม รพ. ได้ถูกต้อง (avg = Σsum/Σtot)"""
    out_rows = []
    for dy in DY:
        g = [c for c in pool if int(c['fy']) == dy]
        if not g: continue
        cnt = Counter(gb_num(c['gb']) for c in g)
        lg = [(c['send']-c['disch']).days for c in g if c['disch']]
        r = {'dy': dy, 'tot': len(g), 'c': [cnt.get(y, 0) for y in CY]}
        if lg: r.update({'mn': min(lg), 'mx': max(lg), 'sum': sum(lg), 'nl': len(lg)})
        out_rows.append(r)
    return out_rows
mx_hosp = []
for hc in {c['hc'] for c in mx_cases}:
    pool = [c for c in mx_cases if c['hc'] == hc]
    mx_hosp.append({'h': pool[0]['hname'], 'p': pool[0]['hcprov'], 'tot': len(pool), 'rows': mx_rows(pool)})
mx_hosp.sort(key=lambda r: -r['tot'])
matrix = {'dy': DY, 'cy': CY, 'region': mx_rows(mx_cases), 'hosp': mx_hosp}
print(f"\n{'═'*72}\n④ ตารางไขว้ ปีจำหน่าย × ปีเคลม (ทั้งเขต, hmain=01)\n{'═'*72}")
print('ปีจำหน่าย\\ปีเคลม'.ljust(18) + 'รวม'.rjust(8) + ''.join(str(y).rjust(9) for y in CY)
      + 'MIN'.rjust(7) + 'MAX'.rjust(7) + 'AVG'.rjust(8))
for r in matrix['region']:
    print(f"  ปีงบ {r['dy']}".ljust(18) + f"{r['tot']:,}".rjust(8) + ''.join((f"{v:,}" if v else '-').rjust(9) for v in r['c'])
          + f"{r['mn']}".rjust(7) + f"{r['mx']}".rjust(7) + f"{r['sum']/r['nl']:.1f}".rjust(8))
R = matrix['region']
print('  รวม'.ljust(18) + f"{sum(r['tot'] for r in R):,}".rjust(8)
      + ''.join(f"{sum(r['c'][i] for r in R):,}".rjust(9) for i in range(len(CY)))
      + f"{min(r['mn'] for r in R)}".rjust(7) + f"{max(r['mx'] for r in R)}".rjust(7)
      + f"{sum(r['sum'] for r in R)/sum(r['nl'] for r in R):.1f}".rjust(8))

out = {
    'cross': cross_out, 'batch': batch, 'matrix': matrix,
    'lag': {'all': lag_all, 'prov': lag_prov, 'hosp': lag_hosp,
            'by_fy': {fy: stats([d for c, d in lag_cases if c['fy'] == fy])
                      for fy in sorted({c['fy'] for c, _ in lag_cases})}},
    'cut': CUT.strftime('%Y-%m-%d'),
    'claim_total': len(claim), 'claim_baht': round(sum(c['baht'] for c in claim)),
    'claim_svc69': sum(1 for c in claim if c['fy'] == '2569'),
    'claim_svc_old': sum(1 for c in claim if c['fy'] != '2569'),
    'claim_months': claim_months, 'claim_cases': claim_cases, 'claim_baht_m': claim_baht,
    'prov': prov_rows,
    'gap_months': gap_months, 'gap_disch': gap_disch, 'gap_claim': gap_claim,
}
p = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'claim_data.json')
with open(p, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
print(f"\n→ {p}")
