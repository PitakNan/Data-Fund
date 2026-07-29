# -*- coding: utf-8 -*-
"""
Extract TKA (Total Knee Arthroplasty) data from the 5-year retrospective CSV
export from สปสช. Produces JSON fragments that get hand-merged into the
HOSP_MONTH_PDX / PROV_DATA / TKA_VAL_MONTHLY / TKA_MONTHS constants embedded
in portal_v2.html.

Source file is Windows-874 (TIS-620) encoded. Convert to UTF-8 first:

  powershell -Command "$sr=New-Object System.IO.StreamReader('<src>.csv',[System.Text.Encoding]::GetEncoding(874)); $sw=New-Object System.IO.StreamWriter('<out>_utf8.csv',$false,[System.Text.Encoding]::UTF8); $sw.Write($sr.ReadToEnd()); $sr.Close(); $sw.Close()"
"""
import csv
import json
import re
import sys
from collections import defaultdict

SRC = r"C:\Users\LENOVO\AppData\Local\Temp\claude\D--\e7652e36-e3cc-43cb-8c3e-a3ede2bb42a6\scratchpad\tka_utf8.csv"

# column indices (positional — avoids Thai unicode-normalization key mismatches)
C_FY_SVC, C_FY_GB, C_PROV_HCODE, C_HNAME, C_HCODE, C_HTYPE, C_PID, C_TRANID, \
    C_AGE, C_AGEMAIN, C_AGESUB, C_SEX, C_ADMIT, C_DISCH, C_SENDDATE, \
    C_PDX, C_SDX, C_PROC, C_REVISION, C_ITEMCODE, C_HMAIN, C_PROVMAIN, \
    C_MEASNAME, C_MEASVAL = range(24)

TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
             'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.']

TKA_PROVS = ['เชียงราย', 'เชียงใหม่', 'ลำปาง', 'น่าน', 'ลำพูน', 'พะเยา', 'แพร่', 'แม่ฮ่องสอน']

# Hcode -> canonical name used by the EXISTING dashboard (PROV_DATA.hospitals / HOSP_PROV keys).
# Built explicitly per-Hcode rather than by string-cleanup of the CSV's หน่วยบริการ column: that
# column uses a different naming convention (space after 'รพ.', full spelled-out names) that in one
# case isn't even a simple whitespace difference (11453 'รพ. สมเด็จพระยุพราชปัว' in the CSV vs the
# dashboard's 'รพร.ปัว') — normalizing this exact set >>> a generic regex covering unknown formats.
# Any Hcode NOT listed here falls back to the CSV's own name (normalize_hname strips the space) —
# meaning genuinely new facilities in a future refresh degrade gracefully instead of erroring.
HCODE_NAME_MAP = {
    '10672': 'รพ.ลำปาง', '10674': 'รพ.เชียงรายประชานุเคราะห์', '10713': 'รพ.นครพิงค์',
    '10714': 'รพ.ลำพูน', '10715': 'รพ.แพร่', '10716': 'รพ.น่าน', '10717': 'รพ.พะเยา',
    '10718': 'รพ.เชียงคำ', '10719': 'รพ.ศรีสังวาลย์', '11119': 'รพ.จอมทอง', '11125': 'รพ.ฝาง',
    '11128': 'รพ.สันป่าตอง', '11130': 'รพ.สันทราย', '11169': 'รพ.สูงเม่น', '11177': 'รพ.เวียงสา',
    '11190': 'รพ.พาน', '11192': 'รพ.แม่จัน', '11194': 'รพ.แม่สาย', '11453': 'รพร.ปัว',
    '41509': 'รพ.ศูนย์การแพทย์มหาวิทยาลัยแม่ฟ้าหลวง',
    '13780': 'รพ.มหาราชนครเชียงใหม่ มช.',
    '14550': 'รพ.เชียงใหม่ใกล้หมอ',
    '14555': 'รพ.ศิริเวช ลำพูน',
    '41347': 'รพ.มหาวิทยาลัยพะเยา',
    '11512': 'รพ.ค่ายสุรศักดิ์มนตรี',
    '42186': 'รพ.เอกชน (Hcode 42186 ไม่พบในทำเนียบ)',
}
HCODE_HTYPE_FALLBACK = {
    '41509': 'รพ.มหาวิทยาลัย', '13780': 'รพ.มหาวิทยาลัย', '14550': 'รพ.เอกชน',
    '14555': 'รพ.เอกชน', '41347': 'รพ.มหาวิทยาลัย', '11512': 'รพ.ค่ายทหาร', '42186': 'รพ.เอกชน',
}

# code -> Thai name harvested from the existing embedded HOSP_MONTH_PDX blob
# (codes not covered here default to the code itself, matching prior behaviour)
with open(r"C:\Users\LENOVO\AppData\Local\Temp\claude\D--\e7652e36-e3cc-43cb-8c3e-a3ede2bb42a6\scratchpad\pdx_names.json",
          encoding='utf-8-sig') as f:
    PDX_NAMES = json.load(f)


def th_month_label(yyyymmdd_be):
    y = int(yyyymmdd_be[:4]) - 543  # BE -> CE
    m = int(yyyymmdd_be[4:6])
    return f"{TH_MONTHS[m-1]}{str((y+543) % 100).zfill(2)}"


def normalize_hname(name):
    """CSV names have a space after the 'รพ.' prefix ('รพ. จอมทอง'); the dashboard's existing
    hospital keys (PROV_DATA.hospitals, HOSP_PROV) never do ('รพ.จอมทอง') — collapse it so new
    per-hospital records key-match the existing province/hospital lookups instead of silently
    becoming orphan keys."""
    return re.sub(r'^(รพ\.|รพร\.|รพศ\.|รพท\.)\s+', r'\1', name)


def age_bucket(age_main):
    if age_main == 'น้อยกว่า 55 ปี':
        return 'lt55'
    if age_main == '55 - 59 ปี':
        return 'a5559'
    if age_main == '60 ปีขึ้นไป':
        return 'ge60'
    return None


def load_rows():
    with open(SRC, encoding='utf-8-sig', newline='') as f:
        r = csv.reader(f)
        header = next(r)
        assert len(header) == 24, f"unexpected column count: {len(header)}"
        rows = list(r)
    return rows


def pivot_cases(rows):
    """group the long/unpivoted rows into one record per case (3 measures -> 1 dict)"""
    cases = {}
    dropped_outside_region = 0
    for row in rows:
        if row[C_HMAIN] != '01':
            dropped_outside_region += 1
            continue
        # NOTE: one admission can have 2+ Item Code billing lines (e.g. 7005A main procedure +
        # 7402 add-on component) each with their own 3-measure rows — group by the case identity
        # only (excluding Item Code) and sum baht across lines, or every case gets double-counted.
        key = (row[C_PID], row[C_TRANID], row[C_HCODE], row[C_ADMIT])
        c = cases.get(key)
        if c is None:
            c = {
                'hcode': row[C_HCODE],
                'hname': HCODE_NAME_MAP.get(row[C_HCODE]) or normalize_hname(row[C_HNAME].strip()) or f'Hcode {row[C_HCODE]}',
                'htype': row[C_HTYPE].strip() or HCODE_HTYPE_FALLBACK.get(row[C_HCODE], ''),
                'prov': row[C_PROVMAIN],
                'fy': row[C_FY_SVC],
                'admit': row[C_ADMIT],
                'month': th_month_label(row[C_DISCH]),  # discharge date = standard hospital case-month attribution
                'age_bucket': age_bucket(row[C_AGEMAIN]),
                'sex': row[C_SEX],
                'is_revision': bool(row[C_REVISION].strip()),
                'pdx': row[C_PDX],
                'baht': 0.0,
            }
            cases[key] = c
        if row[C_MEASNAME] == 'ยอดชดเชย(บาท)':
            try:
                c['baht'] += float(row[C_MEASVAL])
            except ValueError:
                pass
    return list(cases.values()), dropped_outside_region


def build_hosp_month_pdx(cases):
    hmp = defaultdict(lambda: defaultdict(lambda: {
        'total': 0, 'lt55': 0, 'a5559': 0, 'ge60': 0, 'ge55': 0,
        '_pdx': defaultdict(lambda: {'cases': 0, 'lt55': 0, 'a5559': 0, 'ge60': 0, 'ge55': 0}),
    }))
    for c in cases:
        bucket = hmp[c['hname']][c['month']]
        bucket['total'] += 1
        ab = c['age_bucket']
        if ab:
            bucket[ab] += 1
            if ab in ('a5559', 'ge60'):
                bucket['ge55'] += 1
        p = bucket['_pdx'][c['pdx']]
        p['cases'] += 1
        if ab:
            p[ab] += 1
            if ab in ('a5559', 'ge60'):
                p['ge55'] += 1

    out = {}
    for hosp, months in hmp.items():
        out[hosp] = {}
        for mlabel, md in months.items():
            pdx_list = []
            for code, p in md['_pdx'].items():
                pdx_list.append({
                    'code': code, 'name': PDX_NAMES.get(code, code),
                    'cases': p['cases'], 'lt55': p['lt55'], 'a5559': p['a5559'],
                    'ge60': p['ge60'], 'ge55': p['ge55'],
                })
            pdx_list.sort(key=lambda x: -x['cases'])
            out[hosp][mlabel] = {
                'total': md['total'], 'lt55': md['lt55'], 'a5559': md['a5559'],
                'ge60': md['ge60'], 'ge55': md['ge55'], 'pdx': pdx_list,
            }
    return out


def month_sort_key(label, month_order):
    """label like 'ต.ค.68' -> comparable (calendar_year, month_idx). Thai FY: ต.ค. starts new FY
    but for plain chronological sort we just need real calendar order."""
    abbr = label[:-2]
    yy = int(label[-2:])
    midx = month_order.index(abbr)
    # calendar year continuity: Jan(yy) .. Dec(yy) is one BE calendar year
    return (yy, midx)


def build_monthly_series(cases, all_months_sorted):
    per_prov_month_cases = defaultdict(lambda: defaultdict(int))
    per_prov_month_baht = defaultdict(lambda: defaultdict(float))
    for c in cases:
        per_prov_month_cases[c['prov']][c['month']] += 1
        per_prov_month_baht[c['prov']][c['month']] += c['baht']
        per_prov_month_cases['ทั้งหมด'][c['month']] += 1
        per_prov_month_baht['ทั้งหมด'][c['month']] += c['baht']

    prov_data_monthly = {}
    val_monthly = {}
    for prov in TKA_PROVS + ['ทั้งหมด']:
        prov_data_monthly[prov] = [{'label': m, 'cases': per_prov_month_cases[prov].get(m, 0)} for m in all_months_sorted]
        val_monthly[prov] = [round(per_prov_month_baht[prov].get(m, 0)) for m in all_months_sorted]
    return prov_data_monthly, val_monthly


def build_prov_data_current_fy(cases, current_fy):
    """headline totals (total/lt55/.../hospitals) scoped to the latest fiscal year in the file,
    matching how the dashboard's quota/target hero section already works (current-year progress).
    The hospitals list is the UNION of every hospital seen across ALL 5 years (cases=0 allowed) so
    tkaNormMonthly's age-filtered trend lookup (which iterates d.hospitals against every month in
    TKA_MONTHS, not just the current-FY months) still finds facilities that were only active in
    earlier years."""

    def all_hospitals_for(prov_filter):
        hosp_type = {}
        for c in cases:
            if prov_filter is None or c['prov'] == prov_filter:
                hosp_type[c['hname']] = hosp_type.get(c['hname']) or c['htype']
        return hosp_type

    def summarize(fy_subset, prov_filter):
        total = len(fy_subset)
        lt55 = sum(1 for c in fy_subset if c['age_bucket'] == 'lt55')
        a5559 = sum(1 for c in fy_subset if c['age_bucket'] == 'a5559')
        ge60 = sum(1 for c in fy_subset if c['age_bucket'] == 'ge60')
        male = sum(1 for c in fy_subset if c['sex'] == '1')
        female = sum(1 for c in fy_subset if c['sex'] == '2')
        revision = sum(1 for c in fy_subset if c['is_revision'])
        hosp_cases = defaultdict(int)
        for c in fy_subset:
            hosp_cases[c['hname']] += 1
        hosp_type = all_hospitals_for(prov_filter)
        hospitals = [{'hosp': h, 'htype': hosp_type.get(h, ''), 'cases': hosp_cases.get(h, 0)}
                     for h in sorted(hosp_type, key=lambda h: -hosp_cases.get(h, 0))]
        return {
            'total': total, 'lt55': lt55, 'a5559': a5559, 'ge60': ge60, 'ge55': a5559 + ge60,
            'male': male, 'female': female, 'revision': revision, 'hospitals': hospitals,
        }

    fy_cases = [c for c in cases if c['fy'] == current_fy]
    out = {'ทั้งหมด': summarize(fy_cases, None)}
    for prov in TKA_PROVS:
        out[prov] = summarize([c for c in fy_cases if c['prov'] == prov], prov)
    return out


def build_quota_base(cases, basis_fy):
    """real per-province case counts for basis_fy — replaces the old ME-hardcoded fy68 dict"""
    counts = defaultdict(int)
    for c in cases:
        if c['fy'] == basis_fy and c['prov'] in TKA_PROVS:
            counts[c['prov']] += 1
    return {p: counts.get(p, 0) for p in TKA_PROVS}


def main():
    rows = load_rows()
    cases, dropped = pivot_cases(rows)
    print(f"rows={len(rows)} pivoted_cases={len(cases)} dropped_outside_region1={dropped}", file=sys.stderr)

    all_months = sorted(set(c['month'] for c in cases), key=lambda l: month_sort_key(l, TH_MONTHS))
    print(f"month range: {all_months[0]} .. {all_months[-1]} ({len(all_months)} months)", file=sys.stderr)

    hmp = build_hosp_month_pdx(cases)
    monthly, val_monthly = build_monthly_series(cases, all_months)

    fys_present = sorted(set(c['fy'] for c in cases))
    current_fy = fys_present[-1]
    basis_fy = fys_present[-2] if len(fys_present) > 1 else current_fy
    print(f"fiscal years present: {fys_present}; current_fy={current_fy}; quota basis_fy={basis_fy}", file=sys.stderr)

    prov_data = build_prov_data_current_fy(cases, current_fy)
    for prov, pd in prov_data.items():
        pd['monthly'] = monthly[prov]

    quota_base = build_quota_base(cases, basis_fy)

    tka_total = prov_data['ทั้งหมด']['total']
    print(f"TKA_TOTAL (real, FY{current_fy} to date) = {tka_total}", file=sys.stderr)
    print(f"quota_base (real FY{basis_fy}) = {quota_base}", file=sys.stderr)

    out_dir = r"C:\Users\LENOVO\AppData\Local\Temp\claude\D--\e7652e36-e3cc-43cb-8c3e-a3ede2bb42a6\scratchpad"
    with open(out_dir + r"\out_hosp_month_pdx.json", 'w', encoding='utf-8') as f:
        json.dump(hmp, f, ensure_ascii=False, separators=(',', ':'))
    with open(out_dir + r"\out_prov_data.json", 'w', encoding='utf-8') as f:
        json.dump(prov_data, f, ensure_ascii=False, separators=(',', ':'))
    with open(out_dir + r"\out_val_monthly.json", 'w', encoding='utf-8') as f:
        json.dump(val_monthly, f, ensure_ascii=False, separators=(',', ':'))
    with open(out_dir + r"\out_months.json", 'w', encoding='utf-8') as f:
        json.dump(all_months, f, ensure_ascii=False)
    with open(out_dir + r"\out_quota_base.json", 'w', encoding='utf-8') as f:
        json.dump({'basis_fy': basis_fy, 'current_fy': current_fy, 'quota': quota_base, 'tka_total': tka_total}, f, ensure_ascii=False, indent=1)

    print("done", file=sys.stderr)


if __name__ == '__main__':
    main()
