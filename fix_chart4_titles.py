# Fix chart title typo: '2/3th sub-harmonics' -> '1/3th sub-harmonics' (Chart 4)
# Target sheets: 'summary P1', 'summary P2', 'Compare P1', 'Compare P2'

import zipfile, os, glob, shutil, re, xml.etree.ElementTree as ET

SEARCH_DIR  = r'C:/Users/seijis/Desktop/04_MMA_Data_Pmax_Harmonics'
TARGET_SHEETS = {'summary P1', 'summary P2', 'Compare P1', 'Compare P2'}

# '2/3th' is stored in a single <a:t> element
FIND    = '2/3th'
REPLACE = '1/3th'


def get_wb_rels(zf):
    with zf.open('xl/_rels/workbook.xml.rels') as f:
        tree = ET.parse(f)
    return {r.get('Id'): r.get('Target') for r in tree.getroot()}


def get_sheet_drawings(zf, sheet_target):
    parts = sheet_target.rsplit('/', 1)
    rels_path = f'xl/{parts[0]}/_rels/{parts[1]}.rels' if len(parts) == 2 else f'xl/_rels/{parts[0]}.rels'
    drawings = []
    try:
        with zf.open(rels_path) as f:
            tree = ET.parse(f)
        for rel in tree.getroot():
            if 'drawing' in rel.get('Type', '').lower():
                t = rel.get('Target', '')
                drawings.append(t[3:] if t.startswith('../') else t)
    except KeyError:
        pass
    return drawings


def get_drawing_charts(zf, drawing_path):
    parts = drawing_path.rsplit('/', 1)
    rels_path = f'xl/{parts[0]}/_rels/{parts[1]}.rels'
    charts = []
    try:
        with zf.open(rels_path) as f:
            tree = ET.parse(f)
        for rel in tree.getroot():
            if 'chart' in rel.get('Type', '').lower():
                t = rel.get('Target', '')
                charts.append('xl/' + t[3:] if t.startswith('../') else 'xl/' + t)
    except KeyError:
        pass
    return charts


def fix_chart_title(xml_bytes):
    xml_str = xml_bytes.decode('utf-8')
    c_title_pat = re.compile(r'(<c:title\b[^>]*>)(.*?)(</c:title>)', re.DOTALL)
    at_pat = re.compile(r'(<a:t>)(.*?)(</a:t>)', re.DOTALL)
    changes = []

    def replace_in_title(m):
        content = m.group(2)
        all_texts = ''.join(re.findall(r'<a:t>(.*?)</a:t>', content, re.DOTALL))
        # Only change if title contains '2/3th sub-harmonic' (not just any '2/3th')
        if '2/3th sub-harmonic' not in all_texts:
            return m.group(0)

        def replace_at(am):
            text = am.group(2)
            if FIND in text:
                new_text = text.replace(FIND, REPLACE)
                changes.append(f'"{all_texts[:60].strip()}"')
                return am.group(1) + new_text + am.group(3)
            return am.group(0)

        new_content = at_pat.sub(replace_at, content)
        return m.group(1) + new_content + m.group(3)

    new_xml = c_title_pat.sub(replace_in_title, xml_str)
    return new_xml.encode('utf-8'), changes


def process_file(xlsx_path):
    all_changes = []
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zf:
            names = zf.namelist()
            with zf.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
            ns  = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            rns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
            wb_rels = get_wb_rels(zf)

            target_charts = set()
            for sheet in tree.getroot().findall(f'.//{ns}sheet'):
                if sheet.get('name') not in TARGET_SHEETS:
                    continue
                rid = sheet.get(rns + 'id')
                sheet_target = wb_rels.get(rid, '')
                for drw in get_sheet_drawings(zf, sheet_target):
                    for cp in get_drawing_charts(zf, drw):
                        target_charts.add((sheet.get('name'), cp))

            if not target_charts:
                return []

            file_contents = {n: zf.open(n).read() for n in names}

    except Exception as e:
        print(f'  ERROR reading {xlsx_path}: {e}')
        return []

    modified = False
    for sheet_name, chart_path in target_charts:
        key = next((k for k in file_contents if k.replace('\\', '/') == chart_path.replace('\\', '/')), None)
        if key is None:
            continue
        new_bytes, changes = fix_chart_title(file_contents[key])
        if changes:
            file_contents[key] = new_bytes
            modified = True
            for c in changes:
                all_changes.append(f'[{sheet_name}] {os.path.basename(chart_path)}: {c}')

    if modified:
        tmp = xlsx_path + '.tmp'
        try:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                for name, data in file_contents.items():
                    zf_out.writestr(name, data)
            shutil.move(tmp, xlsx_path)
        except Exception as e:
            if os.path.exists(tmp):
                os.remove(tmp)
            print(f'  ERROR writing {xlsx_path}: {e}')
            return []

    return all_changes


def main():
    xlsx_files = glob.glob(os.path.join(SEARCH_DIR, '**', '*.xlsx'), recursive=True)
    print(f'Found {len(xlsx_files)} xlsx files\n')

    total = 0
    files_changed = 0
    for fp in sorted(xlsx_files):
        changes = process_file(fp)
        if changes:
            print(f'MODIFIED: {os.path.relpath(fp, SEARCH_DIR)}')
            for c in changes:
                print(f'  {c}')
            print()
            total += len(changes)
            files_changed += 1

    print(f'Done. {files_changed} files modified, {total} chart titles changed.')


if __name__ == '__main__':
    main()
