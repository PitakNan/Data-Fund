# -*- coding: utf-8 -*-
"""extract_ppfs.py — สร้าง _ppfs_data.json สำหรับหน้า PPFS ใน portal_v2.html

ที่มา : D:\\Hospital\\Rh1-SourceData-Archive\\ppfs\\PPFS_RGN1_REP_2569_<YYYYMMDD>.xlsx
        ชีต 'PPFS_ALL_RGN1_256x' = ข้อมูลดิบ "ทุกสังกัด" ทั้งเขต 1
        (แถวที่ 1 = ชื่อเรื่อง+วันที่ประมวลผล, แถวที่ 2 = หัวตารางไทย, แถวที่ 3 = ชื่อฟิลด์ SQL)

⚠️ ประวัติบั๊ก (แก้ 2026-08-06)
   ตัวสร้างเดิมเขียนสดในเซสชันเก่า ไม่ได้เก็บเป็นไฟล์ และมันตัดแถวที่ 'กิจกรรมหลัก' ว่างทิ้งทั้งหมด
   ทำให้ยอดในแดชบอร์ดขาดไป 12,465,300 บาท (2567: 174,000 / 2568: 9,907,200 / 2569: 2,384,100)
   และหน่วยบริการ 2 แห่งที่มีเฉพาะแถวดังกล่าวหายไปจากไฟล์ทั้งหน่วย
   (22865 สถานพยาบาลทัณฑสถานหญิงเชียงใหม่, 51368 เดอ มาลิณ คลินิกเวชกรรม)
   สคริปต์นี้เก็บแถวเหล่านั้นไว้ในกลุ่มท้ายสุด UNMAPPED แทนการทิ้ง — ไม่เดาว่าเป็นกองทุนใด

โครงสร้างผลลัพธ์ (ต้องคงเดิม — portal_v2.html อ่านตามนี้)
   meta.as_of         : ข้อความวันที่ประมวลผล เช่น '19 มิ.ย. 2569'
   years              : ['2567','2568','2569'] เรียงน้อย→มาก
   provs / types      : sorted()  |  type_grps คงลำดับ รพศ.→อื่นๆ
   type_grp_arr       : กลุ่มของ types[i] เรียง index ตรงกัน
   svcs / svc_short   : ชื่อกิจกรรมหลัก sorted() + UNMAPPED ต่อท้ายเป็นตัวสุดท้าย
   svc_yr[si*3+yi]    : [คน, ครั้ง, บาท] ระดับเขต
   hosp[].d[str(si)]  : 9 ตัวเลข = ปี0(คน,ครั้ง,บาท) ปี1(...) ปี2(...)
   hosp[].p / .t      : index ใน provs / types  |  .g = กลุ่มประเภท  |  .a = อำเภอ[:12]

การใช้งาน : python extract_ppfs.py   แล้วรัน build_home_cards.py ต่อ
"""
import json
import os
import re
import sys

import pandas as pd

SRC = r'D:\Hospital\Rh1-SourceData-Archive\ppfs\PPFS_RGN1_REP_2569_20260619.xlsx'
SHEET = 'PPFS_ALL_RGN1_256x'
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_ppfs_data.json')

# ป้ายกลุ่มสำหรับแถวที่ต้นทางไม่ได้กำกับกิจกรรมหลักมา — ห้ามเดาว่าเป็นกองทุนใด
UNMAPPED = '(ยังไม่ระบุกิจกรรมหลัก — รอยืนยันจาก สปสช.)'
UNMAPPED_SHORT = '(ยังไม่ระบุกลุ่ม)'

TYPE_GRPS = ['รพศ.', 'รพท.', 'รพช.', 'สอ./PCU', 'อปท.', 'เอกชน', 'อื่นๆ']

TYPE_GRP = {
    'คลินิก (เอกชน)': 'เอกชน',
    'มูลนิธิ': 'อื่นๆ',
    'ร้านขายยาแผนปัจจุบัน': 'อื่นๆ',
    'ศูนย์วิชาการต่าง ๆ (ในสังกัด สป สธ.)': 'อื่นๆ',
    'ศูนย์สุขภาพชุมชน(ในสังกัด สป.สธ.)': 'สอ./PCU',
    'สถานบริการสังกัดองค์กรปกครองส่วนท้องถิ่น(อปท.)': 'อปท.',
    'สถานบริการสาธารณสุขชุมชน (สสช.)': 'สอ./PCU',
    'สถานีอนามัย (สอ.)': 'สอ./PCU',
    'สังกัดกรมการแพทย์': 'อื่นๆ',
    'สังกัดกรมควบคุมโรค': 'อื่นๆ',
    'สังกัดกรมวิทยาศาสตร์การแพทย์': 'อื่นๆ',
    'สังกัดกรมสุขภาพจิต': 'อื่นๆ',
    'สังกัดกรมอนามัย': 'อื่นๆ',
    'สังกัดกระทรวงกลาโหม': 'อื่นๆ',
    'สังกัดกระทรวงการอุดมศึกษา วิทยาศาสตร์ วิจัยและนวัตกรรม': 'อื่นๆ',
    'สังกัดกระทรวงมหาดไทย / สำนักงานตำรวจแห่งชาติ': 'อื่นๆ',
    'สังกัดกระทรวงยุติธรรม': 'อื่นๆ',
    'สำนักงานสาธารณสุขอำเภอ (สสอ.)': 'อื่นๆ',
    'โรงพยาบาล (เอกชน)': 'เอกชน',
    'โรงพยาบาลชุมชน (รพช.) / โรงพยาบาลยุพราช (รพร.)': 'รพช.',
    'โรงพยาบาลทั่วไป (รพท.)': 'รพท.',
    'โรงพยาบาลศูนย์ (รพศ.)': 'รพศ.',
}

# ป้ายสั้นของ 26 กิจกรรมเดิม — ยกมาจาก _ppfs_data.json เดิมทั้งชุด
# เพื่อให้ป้ายบนแดชบอร์ดไม่เปลี่ยน (กฎ truncate เดิมขึ้นกับสระ/วรรณยุกต์ไทย ไม่ reproduce ตรง)
# กิจกรรมใหม่ที่ไม่มีในนี้จะใช้ short_label() สร้างให้
SHORT_MAP = {
    "1. บริการฝากครรภ์": "1. ฝากครรภ์",
    "10.บริการยุติการตั้งครรภ์ที่ไม่ปลอดภัย": "10. ยุติการตั้งครรภ์ที่ไม่",
    "11.บริการแว่นตาเด็ก": "11. แว่นตาเด็ก",
    "12.บริการคัดกรองมะเร็งปากมดลูก": "12. คัดกรองมะเร็งปากมดลูก",
    "13.บริการคัดกรองรอยโรคเสี่ยงมะเร็งและมะเร็งช่องปาด (Ca Oral Screening)": "13. คัดกรองรอยโรคเสี่ยงมะเ",
    "14.บริการตรวจคัดกรองและค้นหาวัณโรคในกลุ่มเสี่ยงสูง": "14. ตรวจคัดกรองและค้นหาวัณ",
    "15. บริการคัดกรองเบาหวานและไขมันในเลือด": "15. คัดกรองเบาหวานและไขมัน",
    "16. บริการตรวจคัดกรองยีนกลายพันธ์โรคมะเร็งเต้านม (BRCA1/BRCA2": "16. ตรวจคัดกรองยีนกลายพันธ",
    "17.บริการคัดกรองโลหิตจางจากการขาดธาตุเหล็ก": "17. คัดกรองโลหิตจางจากการข",
    "18. บริการยาเม็ดเสริมธาตุเหล็กและกรดโฟลิค": "18. ยาเม็ดเสริมธาตุเหล็กแล",
    "19.บริการเคลือบฟลูออไรด์": "19. เคลือบฟลูออไรด์",
    "2.บริการป้องกันและควบคุมโรคโลหิตจางธาลัสซีเมียในหญิงตั้งครรภ์": "2. ป้องกันและควบคุมโรคโลห",
    "20. บริการคัดกรองมะเร็งลำไส้ใหญ่และลำไส้ตรง (Fit Test)": "20. คัดกรองมะเร็งลำไส้ใหญ่",
    "21.บริการคัดกรองไวรัสตับอักเสบบี": "21. คัดกรองไวรัสตับอักเสบบ",
    "22.บริการคัดกรองไวรัสตับอักเสบซี": "22. คัดกรองไวรัสตับอักเสบซ",
    "23.บริการคัดกรองพยาธิใบไม้ตับด้วยการตรวจปัสสาวะ": "23. คัดกรองพยาธิใบไม้ตับด้",
    "24.บริการคัดกรองมะเร็งเต้านมด้วยเครื่องแมมแกรมและอัลตราซาวด์": "24. คัดกรองมะเร็งเต้านมด้ว",
    "25.บริการให้วัคซีนป้องกันโรค": "25. ให้วัคซีนป้องกันโรค",
    "3.1บริการคัดกรองธาลัสซีเมียในสามีหรือคู่ของหญิงตั้งครรภ์ที่เป็นชาย": "3.1. คัดกรองธาลัสซีเมียในสา",
    "3.2บริการคัดกรองโรคซิฟิลิส": "3.2. คัดกรองโรคซิฟิลิส",
    "4.บริการป้องกันและควบคุมกลุ่มอาการดาวน์ในหญิงตั้งครรภ์": "4. ป้องกันและควบคุมกลุ่มอ",
    "5. บริการป้องกันและควบคุมภาวะพร่องฮอร์โมไทรอยด์ (TSH)": "5. ป้องกันและควบคุมภาวะพร",
    "6. บริการตรวจคัดกรองผู้ป่วยโรคพันธุกรรมเมตาบอลิกด้วยเครื่อง  Tandem mass spectrometry (TMS) ": "6. ตรวจคัดกรองผู้ป่วยโรคพ",
    "7.บริการตรวจหลังคลอด": "7. ตรวจหลังคลอด",
    "8.บริการทดสอบการตั้งครรภ์": "8. ทดสอบการตั้งครรภ์",
    "9.บริการวางแผนครอบครัวและการป้องกันตั้งครรภ์ไม่พึงประสงค์": "9. วางแผนครอบครัวและการป้",
}

COLS = {'ปี งปม.(ส่งข้อมูล)': 'fy', 'จังหวัด': 'prov', 'รหัสหน่วยบริการ': 'hcode',
        'หน่วยบริการ': 'hname', 'อำเภอ': 'amp', 'ประเภท': 'htype',
        'กิจกรรมหลัก': 'svc', 'จำนวน(คน)': 'pid', 'จำนวน(ครั้ง)': 'vst',
        'การเบิกจ่าย(บาท)': 'amt'}


def short_label(name):
    """ป้ายสั้นสำหรับกิจกรรมใหม่ที่ยังไม่มีใน SHORT_MAP"""
    m = re.match(r'^([\d.]+?)\.?\s*บริการ\s*(.*)$', name)
    if not m:
        return name[:26]
    pre = m.group(1).rstrip('.') + '. '
    return pre + m.group(2)[:max(0, 26 - len(pre))]


def read_as_of():
    """อ่านวันที่ประมวลผลจากเซลล์ A1
    'ข้อมูลบริการทุกสิทธิ ประมวลผล ณ วันที่ 19 มิ.ย.2569' -> '19 มิ.ย. 2569'"""
    import openpyxl
    wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    try:
        a1 = str(wb[SHEET]['A1'].value or '')
    finally:
        wb.close()
    m = re.search(r'(\d{1,2})\s*([ก-๙.]+?)\s*(25\d\d)', a1)
    if m:
        return '%s %s %s' % (m.group(1), m.group(2), m.group(3))
    print('⚠️  อ่านวันที่จาก A1 ไม่ได้ (%r) — คงค่าเดิมในไฟล์' % a1)
    try:
        with open(OUT, encoding='utf-8') as f:
            return json.load(f)['meta']['as_of']
    except Exception:
        return ''


def main():
    if not os.path.exists(SRC):
        sys.exit('ไม่พบไฟล์ต้นฉบับ: %s' % SRC)

    df = pd.read_excel(SRC, sheet_name=SHEET, header=1, dtype=str)
    df = df.drop(index=0).reset_index(drop=True)   # แถวชื่อฟิลด์ SQL (FY_SEND, PROV1, ...)

    missing = [c for c in COLS if c not in df.columns]
    if missing:
        sys.exit('โครงสร้างคอลัมน์ต้นฉบับเปลี่ยน ขาด: %s' % missing)
    df = df[list(COLS)].rename(columns=COLS)

    for c in ('pid', 'vst', 'amt'):
        df[c] = pd.to_numeric(df[c]).fillna(0).astype('int64')

    # ⬇️ หัวใจของการแก้บั๊ก: เดิมแถวเหล่านี้ถูกตัดทิ้ง
    df['svc'] = df['svc'].fillna(UNMAPPED)
    n_unmapped = int((df['svc'] == UNMAPPED).sum())

    years = sorted(df['fy'].unique())
    provs = sorted(df['prov'].unique())
    types = sorted(df['htype'].unique())
    type_grp_arr = [TYPE_GRP.get(t, 'อื่นๆ') for t in types]
    unknown = [t for t in types if t not in TYPE_GRP]
    if unknown:
        print('⚠️  ประเภทหน่วยใหม่ที่ยังไม่มีในผัง TYPE_GRP (จัดเข้า "อื่นๆ"): %s' % unknown)

    real = sorted(s for s in df['svc'].unique() if s != UNMAPPED)
    svcs = real + ([UNMAPPED] if n_unmapped else [])
    svc_short = [SHORT_MAP.get(s) or short_label(s) for s in real]
    if n_unmapped:
        svc_short.append(UNMAPPED_SHORT)
    for s in real:
        if s not in SHORT_MAP:
            print('ℹ️  กิจกรรมหลักใหม่ "%s" -> ป้ายสั้น "%s"' % (s, short_label(s)))

    ny = len(years)
    yi = {y: i for i, y in enumerate(years)}
    si = {s: i for i, s in enumerate(svcs)}
    pi = {p: i for i, p in enumerate(provs)}
    ti = {t: i for i, t in enumerate(types)}

    # ---- svc_yr : ระดับเขต ----
    svc_yr = [[0, 0, 0] for _ in range(len(svcs) * ny)]
    for r in df.groupby(['svc', 'fy'], as_index=False)[['pid', 'vst', 'amt']].sum().itertuples(index=False):
        svc_yr[si[r.svc] * ny + yi[r.fy]] = [int(r.pid), int(r.vst), int(r.amt)]

    # ---- hosp : รายหน่วยบริการ ----
    meta = (df.groupby('hcode', as_index=False)
              .agg(hname=('hname', 'first'), prov=('prov', 'first'),
                   amp=('amp', 'first'), htype=('htype', 'first'))
              .sort_values('hcode'))
    buckets = {}
    for r in df.groupby(['hcode', 'svc', 'fy'], as_index=False)[['pid', 'vst', 'amt']].sum().itertuples(index=False):
        d = buckets.setdefault(r.hcode, {}).setdefault(str(si[r.svc]), [0] * (ny * 3))
        o = yi[r.fy] * 3
        d[o], d[o + 1], d[o + 2] = int(r.pid), int(r.vst), int(r.amt)

    hosp = [{'h': r.hcode, 'n': r.hname, 'p': pi[r.prov], 't': ti[r.htype],
             'g': TYPE_GRP.get(r.htype, 'อื่นๆ'), 'a': (r.amp or '')[:12],
             'd': buckets.get(r.hcode, {})}
            for r in meta.itertuples(index=False)]

    out = {'meta': {'as_of': read_as_of()}, 'years': years, 'provs': provs, 'types': types,
           'type_grps': TYPE_GRPS, 'type_grp_arr': type_grp_arr,
           'svcs': svcs, 'svc_short': svc_short, 'svc_yr': svc_yr, 'hosp': hosp}

    # ---- กระทบยอดก่อนเขียน — ทั้ง svc_yr และ hosp ต้องตรงกับข้อมูลดิบทุกปี ----
    ok = True
    print('กระทบยอดกับข้อมูลดิบ:')
    for y in years:
        r = df[df['fy'] == y]
        exp = [int(r['pid'].sum()), int(r['vst'].sum()), int(r['amt'].sum())]
        got_s = [sum(svc_yr[i * ny + yi[y]][m] for i in range(len(svcs))) for m in range(3)]
        got_h = [sum(v[yi[y] * 3 + m] for h in hosp for v in h['d'].values()) for m in range(3)]
        for lbl, got in (('svc_yr', got_s), ('hosp', got_h)):
            if got != exp:
                ok = False
            print('  %s %-7s คน %12s  ครั้ง %12s  บาท %15s   %s'
                  % (y, lbl, f'{got[0]:,}', f'{got[1]:,}', f'{got[2]:,}',
                     'OK' if got == exp else 'ไม่ตรง! ควรเป็น %s' % [f'{x:,}' for x in exp]))
    if not ok:
        sys.exit('❌ กระทบยอดไม่ตรง — ไม่เขียนไฟล์')

    with open(OUT, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))

    print()
    print('เขียน %s (%.0f KB)' % (OUT, os.path.getsize(OUT) / 1024))
    print('  ปี %s | จังหวัด %d | ประเภท %d | กิจกรรมหลัก %d | หน่วยบริการ %d'
          % (','.join(years), len(provs), len(types), len(svcs), len(hosp)))
    if n_unmapped:
        print('  แถวที่ต้นทางไม่กำกับกิจกรรมหลัก %d แถว -> กลุ่ม "%s" index %d (เดิมถูกตัดทิ้ง)'
              % (n_unmapped, UNMAPPED_SHORT, len(svcs) - 1))


if __name__ == '__main__':
    main()
