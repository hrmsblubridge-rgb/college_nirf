"""
Generate ALL 3 Excel files using xlsxwriter (WPS Office compatible).
1. college_rankings.xlsx - NIRF Rankings Original (Rank | Name | City | State)
2. college_rankings_with_shortnames.xlsx - NIRF with Short Names
3. indian_colleges_sorted.xlsx - All India Colleges
"""
import sys, json, re
sys.path.insert(0, '/app')
sys.path.insert(0, '/app/backend')
import xlsxwriter

# Import data lists from existing scripts
from scripts.create_college_excel import ranked, band_101_150, band_151_200, band_201_300
from scripts.create_college_shortnames import ranked as ranked_sn, \
    band_101_150 as band_sn_101, band_151_200 as band_sn_151, band_201_300 as band_sn_201


def clean_str(s):
    if not s: return ''
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', str(s)).strip()


def make_formats(wb):
    """Create shared format dict for a workbook."""
    f = {}
    base = {'font_name': 'Calibri', 'valign': 'vcenter', 'border': 1, 'border_color': '#D1D5DB'}

    f['title'] = wb.add_format({'font_name': 'Calibri', 'bold': True, 'font_size': 15,
        'font_color': 'white', 'bg_color': '#0F4C81', 'align': 'center', 'valign': 'vcenter'})
    f['summary'] = wb.add_format({'font_name': 'Calibri', 'italic': True, 'font_size': 10,
        'font_color': '#374151', 'bg_color': '#DBEAFE', 'align': 'center', 'valign': 'vcenter'})
    f['hdr'] = wb.add_format({'font_name': 'Calibri', 'bold': True, 'font_size': 12,
        'font_color': 'white', 'bg_color': '#1A73E8', 'align': 'center', 'valign': 'vcenter',
        'text_wrap': True, 'border': 1, 'border_color': 'white'})
    f['band_sep'] = wb.add_format({'font_name': 'Calibri', 'bold': True, 'font_size': 10,
        'font_color': 'white', 'bg_color': '#374151', 'align': 'center', 'valign': 'vcenter'})

    # Data formats (ranked vs unranked, alt rows)
    for prefix, bg, bg_alt in [('r', '#EBF3FD', '#FFFFFF'), ('u', '#F9FAFB', '#FFFFFF')]:
        for alt_name, bg_color in [('', bg), ('_alt', bg_alt)]:
            key = f'{prefix}{alt_name}'
            f[f'{key}_rank'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
                'font_color': '#1A73E8' if prefix == 'r' else '#9CA3AF',
                'bg_color': bg_color, 'align': 'center'})
            f[f'{key}_data'] = wb.add_format({**base, 'font_size': 10,
                'font_color': '#111827', 'bg_color': bg_color})
            f[f'{key}_short'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
                'font_color': '#1557B0', 'bg_color': bg_color})
            f[f'{key}_short_empty'] = wb.add_format({**base, 'font_size': 10,
                'font_color': '#C0C0C0', 'bg_color': bg_color})

    # All-India specific
    f['ai_data'] = wb.add_format({**base, 'font_size': 10})
    f['ai_data_alt'] = wb.add_format({**base, 'font_size': 10, 'bg_color': '#F0F7FF'})
    f['ai_rank'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
        'font_color': '#1557B0', 'align': 'center'})
    f['ai_rank_alt'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
        'font_color': '#1557B0', 'align': 'center', 'bg_color': '#F0F7FF'})
    f['ai_band'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
        'font_color': '#E67E22', 'align': 'center'})
    f['ai_band_alt'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
        'font_color': '#E67E22', 'align': 'center', 'bg_color': '#F0F7FF'})
    f['ai_nl'] = wb.add_format({**base, 'font_size': 10, 'italic': True,
        'font_color': '#DC2626', 'align': 'center'})
    f['ai_nl_alt'] = wb.add_format({**base, 'font_size': 10, 'italic': True,
        'font_color': '#DC2626', 'align': 'center', 'bg_color': '#F0F7FF'})
    f['ai_short'] = wb.add_format({**base, 'font_size': 10, 'bold': True, 'font_color': '#1557B0'})
    f['ai_short_alt'] = wb.add_format({**base, 'font_size': 10, 'bold': True,
        'font_color': '#1557B0', 'bg_color': '#F0F7FF'})

    return f


def get_fmts(f, ranked, alt):
    prefix = 'r' if ranked else 'u'
    suffix = '_alt' if alt else ''
    return f[f'{prefix}{suffix}_rank'], f[f'{prefix}{suffix}_data'], f[f'{prefix}{suffix}_short'], f[f'{prefix}{suffix}_short_empty']


# ══════════════════════════════════════════════════════════════════════════
# FILE 1: NIRF Rankings (Original) - Rank | Name | City | State
# ══════════════════════════════════════════════════════════════════════════
def gen_rankings_original():
    path = '/app/college_rankings.xlsx'
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet('College Rankings')
    f = make_formats(wb)

    ws.set_column(0, 0, 10)   # Rank
    ws.set_column(1, 1, 65)   # Name
    ws.set_column(2, 2, 30)   # City
    ws.set_column(3, 3, 25)   # State

    total = 100 + len(band_101_150) + len(band_151_200) + len(band_201_300)

    # Title
    ws.merge_range(0, 0, 0, 3, "India Rankings 2025 - Engineering Colleges (NIRF)", f['title'])
    ws.set_row(0, 38)
    # Summary
    ws.merge_range(1, 0, 1, 3,
        f"Total Ranked Colleges: 100   |   Total Colleges: {total}   |   Source: NIRF 2025", f['summary'])
    ws.set_row(1, 22)
    # Headers
    for ci, h in enumerate(['Rank', 'College Name', 'City', 'State']):
        ws.write(2, ci, h, f['hdr'])
    ws.set_row(2, 28)

    row = 3

    def write_band_sep(r, label):
        ws.merge_range(r, 0, r, 3, label, f['band_sep'])
        ws.set_row(r, 20)

    def write_data(r, rank_val, name, city, state, is_ranked, alt):
        rf, df, _, _ = get_fmts(f, is_ranked, alt)
        ws.write(r, 0, clean_str(str(rank_val)) if rank_val else '', rf)
        ws.write(r, 1, clean_str(name), df)
        ws.write(r, 2, clean_str(city), df)
        ws.write(r, 3, clean_str(state), df)
        ws.set_row(r, 18)

    # Top 100
    write_band_sep(row, "RANKED COLLEGES - Top 100 (Specific NIRF Rank)")
    row += 1
    for i, (rank, name, city, state) in enumerate(ranked):
        write_data(row, rank, name, city, state, True, i % 2 == 1)
        row += 1

    # Bands
    for label, band, rank_str in [
        ("RANK BAND: 101-150", band_101_150, "101-150"),
        ("RANK BAND: 151-200", band_151_200, "151-200"),
        ("RANK BAND: 201-300", band_201_300, "201-300"),
    ]:
        write_band_sep(row, label)
        row += 1
        for i, (name, city, state) in enumerate(band):
            write_data(row, rank_str, name, city, state, False, i % 2 == 1)
            row += 1

    ws.freeze_panes(3, 0)
    wb.close()
    print(f"[1/3] Saved: {path} ({total} colleges)")


# ══════════════════════════════════════════════════════════════════════════
# FILE 2: NIRF Rankings with Short Names
# ══════════════════════════════════════════════════════════════════════════
def gen_rankings_shortnames():
    path = '/app/college_rankings_with_shortnames.xlsx'
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet('Rankings with Short Names')
    f = make_formats(wb)

    ws.set_column(0, 0, 8)    # Rank
    ws.set_column(1, 1, 62)   # Name
    ws.set_column(2, 2, 28)   # Short Name
    ws.set_column(3, 3, 28)   # City
    ws.set_column(4, 4, 22)   # State

    total = 100 + len(band_sn_101) + len(band_sn_151) + len(band_sn_201)

    ws.merge_range(0, 0, 0, 4,
        "India Rankings 2025 - Engineering Colleges with Short Names (NIRF)", f['title'])
    ws.set_row(0, 36)
    ws.merge_range(1, 0, 1, 4,
        f"Total Ranked: 100   |   Total Colleges: {total}   |   Short Name = Globally recognized name only   |   Source: NIRF 2025",
        f['summary'])
    ws.set_row(1, 20)

    for ci, h in enumerate(['Rank', 'College Name', 'Short Name', 'City', 'State']):
        ws.write(2, ci, h, f['hdr'])
    ws.set_row(2, 28)

    row = 3

    def write_band_sep(r, label):
        ws.merge_range(r, 0, r, 4, label, f['band_sep'])
        ws.set_row(r, 19)

    def write_data(r, rank_val, name, short, city, state, is_ranked, alt):
        rf, df, sf, sef = get_fmts(f, is_ranked, alt)
        ws.write(r, 0, clean_str(str(rank_val)) if rank_val else '', rf)
        ws.write(r, 1, clean_str(name), df)
        ws.write(r, 2, clean_str(short), sf if short else sef)
        ws.write(r, 3, clean_str(city), df)
        ws.write(r, 4, clean_str(state), df)
        ws.set_row(r, 18)

    # Top 100
    write_band_sep(row, "RANKED COLLEGES - Top 100 (Specific NIRF Rank 1-100)")
    row += 1
    for i, (rank, name, short, city, state) in enumerate(ranked_sn):
        write_data(row, rank, name, short, city, state, True, i % 2 == 1)
        row += 1

    for label, band, rank_str in [
        ("RANK BAND: 101-150", band_sn_101, "101-150"),
        ("RANK BAND: 151-200", band_sn_151, "151-200"),
        ("RANK BAND: 201-300", band_sn_201, "201-300"),
    ]:
        write_band_sep(row, label)
        row += 1
        for i, (name, short, city, state) in enumerate(band):
            write_data(row, rank_str, name, short, city, state, False, i % 2 == 1)
            row += 1

    ws.freeze_panes(3, 0)
    wb.close()
    print(f"[2/3] Saved: {path} ({total} colleges)")


# ══════════════════════════════════════════════════════════════════════════
# FILE 3: All India Colleges
# ══════════════════════════════════════════════════════════════════════════
def gen_all_india():
    path = '/app/frontend/public/indian_colleges_sorted.xlsx'
    json_path = '/tmp/all_india_colleges.json'

    with open(json_path) as jf:
        data = json.load(jf)

    wb = xlsxwriter.Workbook(path, {'strings_to_numbers': False})
    ws = wb.add_worksheet('All India Colleges')
    f = make_formats(wb)

    ws.set_column(0, 0, 12)   # Rank
    ws.set_column(1, 1, 65)   # Name
    ws.set_column(2, 2, 25)   # Short Name
    ws.set_column(3, 3, 20)   # City
    ws.set_column(4, 4, 20)   # State

    for ci, h in enumerate(['Rank', 'College Name', 'Short Name', 'City', 'State']):
        ws.write(0, ci, h, f['hdr'])
    ws.set_row(0, 30)

    for idx, d in enumerate(data):
        row = idx + 1
        rank = clean_str(d.get('rank', 'NL'))
        name = clean_str(d.get('college_name', ''))
        short = clean_str(d.get('short_names', ''))
        city = clean_str(d.get('city', ''))
        state = clean_str(d.get('state', ''))
        alt = row % 2 == 0

        # Rank
        if rank == 'NL':
            ws.write(row, 0, rank, f['ai_nl_alt'] if alt else f['ai_nl'])
        elif rank.startswith(("101-", "151-", "201-")):
            ws.write(row, 0, rank, f['ai_band_alt'] if alt else f['ai_band'])
        else:
            ws.write(row, 0, rank, f['ai_rank_alt'] if alt else f['ai_rank'])

        # Name
        ws.write(row, 1, name, f['ai_data_alt'] if alt else f['ai_data'])
        # Short Name
        if short:
            ws.write(row, 2, short, f['ai_short_alt'] if alt else f['ai_short'])
        else:
            ws.write(row, 2, '', f['ai_data_alt'] if alt else f['ai_data'])
        # City
        ws.write(row, 3, city, f['ai_data_alt'] if alt else f['ai_data'])
        # State
        ws.write(row, 4, state, f['ai_data_alt'] if alt else f['ai_data'])

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(data), 4)
    wb.close()
    print(f"[3/3] Saved: {path} ({len(data)} colleges)")


if __name__ == '__main__':
    gen_rankings_original()
    gen_rankings_shortnames()
    gen_all_india()
    print("\nAll 3 Excel files generated successfully!")
