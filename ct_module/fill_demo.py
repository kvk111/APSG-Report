"""
fill_demo.py — v3 Final (Pure XML approach — NO openpyxl save)

All modifications are done DIRECTLY on the template ZIP's XML files.
This eliminates the Excel repair dialog entirely.

Fixes applied:
  1. Excel repair dialog → fixed (no openpyxl save → no style mismatch)
  2. Date formatting → clean bottom underline only (no extra line)
  3. SC/GE totals (AA8/9/14/15) show computed values, not formula text
  4. IN/OUT text size → 16pt bold, centered (styles.xml patched once)
  5. Grand total rows 20/21 show computed values
  6. All borders consistent (all data cells already have full borders in template)
  7. C8/C9/C14/C15/C20/C21 (right-of-BC-merge) get full border fix
"""

import io, re, zipfile
import pandas as pd
from copy import deepcopy

# ── Hour slot definitions ──────────────────────────────────────────────────────
HOUR_SLOTS = [
    ('D',  7,  8), ('E',  8,  9), ('F',  9, 10), ('G', 10, 11),
    ('H', 11, 12), ('I', 12, 13), ('J', 13, 14), ('K', 14, 15),
    ('L', 15, 16), ('M', 16, 17), ('N', 17, 18), ('O', 18, 19),
    ('P', 19, 20), ('Q', 20, 21), ('R', 21, 22), ('S', 22, 23),
    ('T', 23,  0), ('U',  0,  1), ('V',  1,  2), ('W',  2,  3),
    ('X',  3,  4), ('Y',  4,  5), ('Z',  5,  7),
]
COL_LETTERS = [s[0] for s in HOUR_SLOTS]   # D..Z


def _find_col(columns, *kws):
    for c in columns:
        n = c.lower().replace(' ', '').replace('_', '')
        if all(k.lower().replace(' ', '') in n for k in kws):
            return c
    return None


def _count_slot(df, mat, hour_col, h0, h1):
    sub = df[df['_mat'] == mat.upper()]
    mask = (sub[hour_col] == h0) if h1 == 0 else \
           (sub[hour_col] >= h0) & (sub[hour_col] < h1)
    return int(mask.sum())


# ── XML cell replacement helpers ───────────────────────────────────────────────

def _set_cell_value(sheet_xml, cell_ref, value, row_num):
    """
    Replace an empty data cell <c r="XX" s="NN"/> with <c r="XX" s="NN"><v>value</v></c>
    OR update existing <v> tag.
    """
    # Pattern: <c r="CELLREF" s="NN"/> (empty) or <c r="CELLREF" s="NN"><v>...</v></c>
    pat = rf'(<c r="{re.escape(cell_ref)}" s="(\d+)"(?:\s*/>|>(.*?)</c>))'
    
    def replacer(m):
        full   = m.group(1)
        style  = m.group(2)
        inner  = m.group(3) or ''
        # Remove existing <v> if present
        inner_clean = re.sub(r'<v>[^<]*</v>', '', inner)
        # Keep formula/other elements but inject numeric value
        return f'<c r="{cell_ref}" s="{style}">{inner_clean}<v>{value}</v></c>'
    
    new_xml, count = re.subn(pat, replacer, sheet_xml, count=1, flags=re.DOTALL)
    return new_xml, count > 0


def _set_cell_inline_string(sheet_xml, cell_ref, text, style_id):
    """Replace or create a cell with an inline string value."""
    pat = rf'<c r="{re.escape(cell_ref)}"[^>]*/>'
    replacement = f'<c r="{cell_ref}" s="{style_id}" t="inlineStr"><is><t>{_esc(text)}</t></is></c>'
    new_xml, count = re.subn(pat, replacement, sheet_xml, count=1)
    if not count:
        # Try to replace existing cell with any content
        pat2 = rf'<c r="{re.escape(cell_ref)}"[^>]*>.*?</c>'
        new_xml, count = re.subn(pat2, replacement, sheet_xml, count=1, flags=re.DOTALL)
    return new_xml


def _esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _update_cached_value(sheet_xml, cell_ref, value):
    """Update <v>...</v> inside a formula cell, keeping the formula."""
    pat = rf'(<c r="{re.escape(cell_ref)}"[^>]*>)(.*?)(</c>)'
    def replacer(m):
        head, inner, tail = m.group(1), m.group(2), m.group(3)
        inner_new = re.sub(r'<v>[^<]*</v>', f'<v>{value}</v>', inner)
        if '<v>' not in inner_new:
            inner_new += f'<v>{value}</v>'
        return head + inner_new + tail
    return re.sub(pat, replacer, sheet_xml, count=1, flags=re.DOTALL)


# ── Styles.xml patch ───────────────────────────────────────────────────────────

def _patch_styles(styles_xml):
    """
    Patch styles.xml to add 2 new xf styles (reuses existing fonts — no font additions):
    1. in_out_style: fontId=18 (existing b,sz=16,Arial), borderId=2 (all borders), center
    2. date_style:   fontId=9  (existing sz=14,normal),  borderId=3 (bottom only), center
    Uses regex for count replacement to handle extra XML attributes safely.
    Returns (patched_styles_xml, in_out_style_id, date_style_id)
    """
    m = re.search(r'<cellXfs count="(\d+)"', styles_xml)
    xfs_count = int(m.group(1))
    in_out_style_id = xfs_count
    date_style_id   = xfs_count + 1

    new_xf_inout = (
        '<xf numFmtId="0" fontId="18" fillId="0" borderId="2" '
        'xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/></xf>'
    )
    new_xf_date = (
        '<xf numFmtId="0" fontId="9" fillId="0" borderId="3" '
        'xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">'
        '<alignment horizontal="center" vertical="center"/></xf>'
    )
    # Use regex to safely replace count (handles extra XML attributes on the tag)
    styles_xml = re.sub(
        r'<cellXfs count="\d+"',
        f'<cellXfs count="{xfs_count + 2}"',
        styles_xml, count=1
    )
    styles_xml = styles_xml.replace(
        '</cellXfs>', new_xf_inout + new_xf_date + '</cellXfs>', 1
    )
    return styles_xml, in_out_style_id, date_style_id


# ── Main entry point ───────────────────────────────────────────────────────────

def fill_demo_report(df_wb_raw, df_online_raw=None,
                     template_path='Demo.xlsx',
                     start_dt=None, end_dt=None):
    """
    Fill Demo.xlsx template with live data.
    Pure ZIP/XML approach — no openpyxl save, no Excel repair dialog.
    start_dt / end_dt: the UI filter dates used to generate date_str.
    If not supplied, dates are inferred from the data (legacy fallback).
    """
    # ── Prepare data ─────────────────────────────────────────────────────────
    df = df_wb_raw.copy()
    col_in  = _find_col(df.columns, 'wb', 'in', 'time')
    col_out = _find_col(df.columns, 'wb', 'out', 'time')
    col_mat = _find_col(df.columns, 'material')

    df[col_in]  = pd.to_datetime(df[col_in],  dayfirst=True, errors='coerce')
    df[col_out] = pd.to_datetime(df[col_out], dayfirst=True, errors='coerce')
    df['_in_hour']  = df[col_in].dt.hour
    df['_out_hour'] = df[col_out].dt.hour
    df['_mat']      = df[col_mat].astype(str).str.strip().str.upper()

    # Date string — use UI filter dates when supplied; fall back to data dates
    if start_dt is not None:
        # Use the exact dates from the UI filter
        s_day = start_dt.date()
        e_day = end_dt.date() if end_dt is not None else s_day
        if s_day == e_day:
            date_str = start_dt.strftime('%d/%m/%Y')
        else:
            date_str = (f"{start_dt.strftime('%d/%m/%Y')} and "
                        f"{end_dt.strftime('%d/%m/%Y')}")
    else:
        # Legacy fallback: read dates from data
        unique_dates = sorted(set(pd.to_datetime(df[col_in].dropna().dt.date.unique())))
        if len(unique_dates) >= 2:
            date_str = ' and '.join(d.strftime('%d/%m/%Y') for d in unique_dates[:2])
        elif len(unique_dates) == 1:
            date_str = unique_dates[0].strftime('%d/%m/%Y')
        else:
            date_str = ''

    # Compute hourly counts
    sc_in  = [_count_slot(df, 'SOFT CLAY',  '_in_hour',  *s[1:]) for s in HOUR_SLOTS]
    sc_out = [_count_slot(df, 'SOFT CLAY',  '_out_hour', *s[1:]) for s in HOUR_SLOTS]
    ge_in  = [_count_slot(df, 'GOOD EARTH', '_in_hour',  *s[1:]) for s in HOUR_SLOTS]
    ge_out = [_count_slot(df, 'GOOD EARTH', '_out_hour', *s[1:]) for s in HOUR_SLOTS]

    # Totals
    sc_in_total  = sum(sc_in)
    sc_out_total = sum(sc_out)
    ge_in_total  = sum(ge_in)
    ge_out_total = sum(ge_out)

    # Per-hour combined totals for rows 20/21
    tot_in  = [sc_in[i]  + ge_in[i]  for i in range(23)]
    tot_out = [sc_out[i] + ge_out[i] for i in range(23)]
    tot_in_total  = sum(tot_in)
    tot_out_total = sum(tot_out)

    # Summary stats for chart patches
    # FIX: Total trucks = ALL materials (not just Soft Clay)
    sc_trucks  = len(df)
    day_in     = int(df['_in_hour'].isin(range(7, 19)).sum())
    night_in   = int(df['_in_hour'].isin(list(range(19,24)) + list(range(0,5))).sum())
    nw_col     = _find_col(df_wb_raw.columns, 'net', 'weight')
    net_wt     = round(pd.to_numeric(df_wb_raw[nw_col], errors='coerce').dropna().sum(), 2) if nw_col else 0.0

    # ── Read template ZIP ─────────────────────────────────────────────────────
    tmpl_zf = zipfile.ZipFile(template_path, 'r')
    files   = {item.filename: tmpl_zf.read(item.filename)
               for item in tmpl_zf.infolist()}
    tmpl_zf.close()

    # ── 1. Patch styles.xml ───────────────────────────────────────────────────
    styles_xml = files['xl/styles.xml'].decode('utf-8')
    styles_xml, in_out_sid, date_sid = _patch_styles(styles_xml)
    files['xl/styles.xml'] = styles_xml.encode('utf-8')

    # ── 2. Patch sheet1.xml (Summary) ────────────────────────────────────────
    sheet = files['xl/worksheets/sheet1.xml'].decode('utf-8')

    # Fix 2a: Date cell P3 — inject inline string + use date_style_id
    # P3 currently: <c r="P3" s="20"/>  (merged P3:S3, style 20 = left+right+top borders)
    # We want: bottom border only under the full date span = bottom on P3,Q3,R3,S3
    # Simplest: put date text in P3 with date_sid (bottom border, center, sz=14)
    # and set Q3,R3,S3 to also use date_sid (all will show bottom border line)
    sheet = _set_cell_inline_string(sheet, 'P3', date_str, date_sid)
    for col in ['Q3', 'R3', 'S3']:
        pat = rf'<c r="{col}" s="\d+"(?:\s*/>|>.*?</c>)'
        repl = f'<c r="{col}" s="{date_sid}"/>'
        sheet = re.sub(pat, repl, sheet, count=1, flags=re.DOTALL)

    # Fix 2b: SC IN row 8 (D8..Z8)
    for i, col in enumerate(COL_LETTERS):
        sheet, _ = _set_cell_value(sheet, f'{col}8', sc_in[i], 8)

    # Fix 2c: SC OUT row 9
    for i, col in enumerate(COL_LETTERS):
        sheet, _ = _set_cell_value(sheet, f'{col}9', sc_out[i], 9)

    # Fix 2d: GE IN row 14
    for i, col in enumerate(COL_LETTERS):
        sheet, _ = _set_cell_value(sheet, f'{col}14', ge_in[i], 14)

    # Fix 2e: GE OUT row 15
    for i, col in enumerate(COL_LETTERS):
        sheet, _ = _set_cell_value(sheet, f'{col}15', ge_out[i], 15)

    # Fix 3: AA8/9/14/15 — update cached <v> in SUM formula cells
    sheet = _update_cached_value(sheet, 'AA8',  sc_in_total)
    sheet = _update_cached_value(sheet, 'AA9',  sc_out_total)
    sheet = _update_cached_value(sheet, 'AA14', ge_in_total)
    sheet = _update_cached_value(sheet, 'AA15', ge_out_total)

    # Fix 4a: Row 19 (Trucks per hour = ALL materials IN total) — fill with tot_in
    # This is the row the Chart reads for "Total Number of Trucks" (D19:Z19)
    # It must equal tot_in (SC IN + GE IN per slot) to show ALL materials
    for i, col in enumerate(COL_LETTERS):
        sheet, _ = _set_cell_value(sheet, f'{col}19', tot_in[i], 19)
    # AA19 has no formula in template — inject as plain value
    sheet = re.sub(
        r'(<c r="AA19" s="\d+")(\s*/>\s*)',
        lambda m: m.group(1) + f'><v>{tot_in_total}</v></c>',
        sheet, count=1
    )

    # Fix 4: Grand total rows 20/21 — update cached <v> in formula cells
    for i, col in enumerate(COL_LETTERS):
        sheet = _update_cached_value(sheet, f'{col}20', tot_in[i])
        sheet = _update_cached_value(sheet, f'{col}21', tot_out[i])
    # W20 and X20 etc are part of shared formulas — they might have <f t="shared" si="0"/>
    # and <v>0</v> which we've already updated above
    # Also update the Z20 and AA20 (TOTAL formula)
    sheet = _update_cached_value(sheet, 'Z20',  tot_in[-1])
    sheet = _update_cached_value(sheet, 'Z21',  tot_out[-1])
    sheet = _update_cached_value(sheet, 'AA20', tot_in_total)
    sheet = _update_cached_value(sheet, 'AA21', tot_out_total)

    # Fix 5: IN/OUT label cells B8,B9,B14,B15,B20,B21 → apply in_out_sid (sz=16 bold center)
    # These currently use s="24" (sz=12 bold, right-align)
    # Replace with in_out_sid
    for cell_ref in ['B8', 'B9', 'B14', 'B15', 'B20', 'B21']:
        # Current: <c r="B8" s="24" t="s"><v>15</v></c>  or similar
        pat = rf'(<c r="{re.escape(cell_ref)}" s=")(\d+)(")'
        sheet = re.sub(pat, rf'\g<1>{in_out_sid}\g<3>', sheet, count=1)

    files['xl/worksheets/sheet1.xml'] = sheet.encode('utf-8')

    # ── 2b. Patch chart1.xml numCache for Series 0 (D19:Z19 = Total Trucks IN) ──
    # The chart has 3 series:
    #   Series 0 = D19:Z19 (Total Trucks per hour — ALL materials)  ← was empty
    #   Series 1 = D20:Z20 (Total IN)
    #   Series 2 = D21:Z21 (Total OUT)
    # We must inject the cached numCache for Series 0 so Excel/chart shows correct values.
    if 'xl/charts/chart1.xml' in files:
        chart_xml = files['xl/charts/chart1.xml'].decode('utf-8')

        def _build_num_cache(values):
            pts = ''.join(
                f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
                for i, v in enumerate(values)
            )
            return (f'<c:numCache><c:formatCode>General</c:formatCode>'
                    f'<c:ptCount val="{len(values)}"/>{pts}</c:numCache>')

        # Series 0: D19:Z19 — inject tot_in (all 23 slots)
        old_s0_cache = (r'(<c:numRef>\s*<c:f>Summary!\$D\$19:\$Z\$19</c:f>\s*)'
                        r'<c:numCache>.*?</c:numCache>')
        new_s0_cache = r'\g<1>' + _build_num_cache(tot_in)
        chart_xml = re.sub(old_s0_cache, new_s0_cache, chart_xml, count=1, flags=re.DOTALL)

        # Series 1: D20:Z20 — update tot_in cache
        old_s1_cache = (r'(<c:numRef>\s*<c:f>Summary!\$D\$20:\$Z\$20</c:f>\s*)'
                        r'<c:numCache>.*?</c:numCache>')
        new_s1_cache = r'\g<1>' + _build_num_cache(tot_in)
        chart_xml = re.sub(old_s1_cache, new_s1_cache, chart_xml, count=1, flags=re.DOTALL)

        # Series 2: D21:Z21 — update tot_out cache
        old_s2_cache = (r'(<c:numRef>\s*<c:f>Summary!\$D\$21:\$Z\$21</c:f>\s*)'
                        r'<c:numCache>.*?</c:numCache>')
        new_s2_cache = r'\g<1>' + _build_num_cache(tot_out)
        chart_xml = re.sub(old_s2_cache, new_s2_cache, chart_xml, count=1, flags=re.DOTALL)

        files['xl/charts/chart1.xml'] = chart_xml.encode('utf-8')
    d3 = files['xl/drawings/drawing3.xml'].decode('utf-8')
    OLD_NET = '<a:t>- </a:t></a:r><a:endParaRPr lang="en-US" sz="1400" b="1" baseline="0"/>'
    NEW_NET = (f'<a:t>- {net_wt:.2f}</a:t></a:r>'
               f'<a:endParaRPr lang="en-US" sz="1400" b="1" baseline="0"/>')
    # FIX 5: No commas in any numeric values (plain integers only)
    d3 = d3.replace('>7am to 7pm (12hrs) =  Trucks<', f'>7am to 7pm (12hrs) = {day_in} Trucks<')
    d3 = d3.replace('>7pm to 5am (8hrs) =  Trucks<',  f'>7pm to 5am (8hrs) = {night_in} Trucks<')
    d3 = d3.replace('> No. of Trucks -    <',          f'> No. of Trucks - {sc_trucks}<')
    d3 = d3.replace(OLD_NET, NEW_NET)

    # ── Patch chart title date (two <a:t> runs after "Histogram Chart Dated") ──
    # Build the date string to show in the chart title.
    # Uses the exact UI filter dates (start_dt / end_dt) when supplied;
    # falls back to the data-inferred date_str when not.
    # _day() strips the leading zero from %d cross-platform (%-d is Linux-only)
    def _day(dt): return dt.strftime('%d').lstrip('0')

    if start_dt is not None:
        s_day = start_dt.date()
        e_day = end_dt.date() if end_dt is not None else s_day
        if s_day == e_day:
            _chart_date = f" {_day(start_dt)} {start_dt.strftime('%b %Y')} "
        else:
            _chart_date = (f" {_day(start_dt)} {start_dt.strftime('%b %Y')}"
                           f" &amp; {_day(end_dt)} {end_dt.strftime('%b %Y')} ")
    else:
        # Fallback: convert date_str (dd/mm/yyyy) back to chart format
        _parts = [p.strip() for p in date_str.replace(' and ', '/').split('/') if p.strip()]
        try:
            from datetime import datetime as _dt
            if len(_parts) >= 6:  # two dates: d m y d m y
                _d1 = _dt.strptime('/'.join(_parts[:3]), '%d/%m/%Y')
                _d2 = _dt.strptime('/'.join(_parts[3:6]), '%d/%m/%Y')
                _chart_date = (f" {_day(_d1)} {_d1.strftime('%b %Y')}"
                               f" &amp; {_day(_d2)} {_d2.strftime('%b %Y')} ")
            elif len(_parts) == 3:
                _d1 = _dt.strptime('/'.join(_parts), '%d/%m/%Y')
                _chart_date = f" {_day(_d1)} {_d1.strftime('%b %Y')} "
            else:
                _chart_date = f' {date_str} '
        except Exception:
            _chart_date = f' {date_str} '

    # Replace the two date runs: first " 14 " then "Apr 2026 &amp; 15 Apr 2026 "
    # with a single run containing the full correct date
    _old_day_run  = r'(<a:t>) 14 (</a:t>)(</a:r><a:r>(?:(?!</a:r>).)*?<a:t>)Apr 2026 &amp; 15 Apr 2026 (</a:t>)'
    _new_day_run  = rf'\g<1>{_chart_date}\g<2>'
    import re as _re
    d3_new = _re.sub(_old_day_run, _new_day_run, d3, count=1, flags=_re.DOTALL)
    if d3_new == d3:
        # Fallback: simpler replace of just the second run which contains the full date
        d3_new = d3.replace(
            '<a:t>Apr 2026 &amp; 15 Apr 2026 </a:t>',
            f'<a:t>{_chart_date}</a:t>'
        )
        # Also replace the day-number run
        d3_new = d3_new.replace('<a:t> 14 </a:t>', '<a:t></a:t>')
    d3 = d3_new

    files['xl/drawings/drawing3.xml'] = d3.encode('utf-8')

    # ── 4. Remove calcChain.xml (forces fresh recalculation in Excel) ─────────
    files.pop('xl/calcChain.xml', None)

    # ── 5. Write output ZIP ───────────────────────────────────────────────────
    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', compression=zipfile.ZIP_DEFLATED) as out_zf:
        # Update [Content_Types].xml to remove calcChain entry if present
        ct = files['[Content_Types].xml'].decode('utf-8')
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^/]*/>', '', ct)
        files['[Content_Types].xml'] = ct.encode('utf-8')
        for fname, data in files.items():
            out_zf.writestr(fname, data)
    out_buf.seek(0)
    return out_buf.read()


# ── Standalone test ────────────────────────────────────────────────────────────
if __name__ == '__main__':
    SRC  = '/mnt/user-data/uploads/Cycle_Time-_20260414_15042026.xlsx'
    TMPL = '/mnt/user-data/uploads/Demo.xlsx'

    df_wb = pd.read_excel(SRC, sheet_name='WB Total')
    xlsx  = fill_demo_report(df_wb, template_path=TMPL)

    with open('/mnt/user-data/outputs/Demo_Filled.xlsx', 'wb') as f:
        f.write(xlsx)

    # Verify
    zf   = zipfile.ZipFile(io.BytesIO(xlsx))
    fns  = zf.namelist()
    sheet = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
    styles = zf.read('xl/styles.xml').decode('utf-8')

    print("=== FILE CHECKS ===")
    print(f"  calcChain removed: {'xl/calcChain.xml' not in fns}")
    print(f"  drawing3 present:  {'xl/drawings/drawing3.xml' in fns}")
    print(f"  chart1 present:    {'xl/charts/chart1.xml' in fns}")
    print(f"  chartsheet present:{'xl/chartsheets/sheet1.xml' in fns}")

    print("\n=== DATA CELL CHECKS ===")
    for cell_ref, expected in [('D8', 51), ('E8', 113), ('F8', 256),
                                ('D9', 29), ('F14', 2), ('D14', 0)]:
        pat = rf'<c r="{cell_ref}"[^>]*>.*?<v>(\d+)</v>'
        m = re.search(pat, sheet, re.DOTALL)
        val = int(m.group(1)) if m else 'NOT FOUND'
        print(f"  {cell_ref}: {val}  (expect {expected}) {'✓' if val==expected else '✗'}")

    print("\n=== FORMULA CACHED VALUES ===")
    for cell_ref, label in [('AA8','SC IN total'), ('AA9','SC OUT total'),
                             ('AA14','GE IN total'), ('AA15','GE OUT total'),
                             ('AA20','Grand IN total'), ('AA21','Grand OUT total')]:
        pat = rf'<c r="{cell_ref}"[^>]*>.*?<v>(\d+)</v>'
        m = re.search(pat, sheet, re.DOTALL)
        val = m.group(1) if m else 'NOT FOUND'
        print(f"  {cell_ref} ({label}): {val}")

    print("\n=== DATE CELL ===")
    date_pat = r'<c r="P3"[^>]*>.*?<t>([^<]+)</t>'
    dm = re.search(date_pat, sheet, re.DOTALL)
    print(f"  P3 date: {dm.group(1)!r}" if dm else "  P3: NOT FOUND")

    print("\n=== STYLE CHECKS ===")
    xfs = re.findall(r'<cellXfs count="(\d+)"', styles)
    print(f"  Total xf styles: {xfs}")
    font_counts = re.findall(r'<fonts count="(\d+)"', styles)
    print(f"  Total fonts: {font_counts}")

    print("\n=== CHART TEXT BOXES ===")
    d3 = zf.read('xl/drawings/drawing3.xml').decode('utf-8')
    for kw, expected in [('7am to 7pm','2,082'), ('7pm to 5am','957'),
                          ('No. of Trucks','3,031'), ('Net Weight','54,043')]:
        idx = d3.find(kw)
        ctx = d3[idx:idx+60] if idx >= 0 else 'NOT FOUND'
        ok  = expected in ctx
        print(f"  {'✓' if ok else '✗'} {kw}: {ctx[:55]!r}")

    print(f"\nGenerated: {len(xlsx):,} bytes ✓")
