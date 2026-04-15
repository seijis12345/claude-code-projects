# Fix chart titles by visual position:
# Chart7 (7th chart): '2/3 sub-harmonic' -> '1/3 sub-harmonic'
# Chart9 (9th chart): '2/3 sub-harmonic' -> '1/3 sub-harmonic'
# Target sheets: 'summary P1', 'summary P2', 'Compare P1', 'Compare P2'

import zipfile, os, glob, shutil, re, xml.etree.ElementTree as ET

SEARCH_DIR    = r'C:/Users/seijis/Desktop/04_MMA_Data_Pmax_Harmonics'
TARGET_SHEETS = {'summary P1', 'summary P2', 'Compare P1', 'Compare P2'}
NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

FIND    = '2/3 sub-harmonic'
REPLACE = '1/3 sub-harmonic'


def get_wb_rels(zf):
    with zf.open('xl/_rels/workbook.xml.rels') as f:
        return {r.get('Id'): r.get('Target') for r in ET.parse(f).getroot()}


def get_sheet_drawing(zf, sheet_target):
    parts = sheet_target.rsplit('/', 1)
    rels_path = f'xl/{parts[0]}/_rels/{parts[1]}.rels' if len(parts) == 2 else f'xl/_rels/{parts[0]}.rels'
    try:
        with zf.open(rels_path) as f:
            for rel in ET.parse(f).getroot():
                if 'drawing' in rel.get('Type', '').lower():
                    t = rel.get('Target', '')
                    return t[3:] if t.startswith('../') else t
    except KeyError:
        pass
    return None


def get_charts_in_visual_order(zf, drawing_path):
    drw_key  = 'xl/' + drawing_path
    rels_key = f'xl/{drawing_path.rsplit("/",1)[0]}/_rels/{drawing_path.rsplit("/",1)[1]}.rels'
    with zf.open(rels_key) as f:
        rid_to_path = {}
        for rel in ET.parse(f).getroot():
            if 'chart' in rel.get('Type', '').lower():
                t = rel.get('Target', '')
                rid_to_path[rel.get('Id')] = 'xl/' + t[3:] if t.startswith('../') else 'xl/' + t
    with zf.open(drw_key) as f:
        dtree = ET.fromstring(f.read())
    anchors = []
    for anchor in dtree:
        if 'Anchor' not in anchor.tag.split('}')[-1]:
            continue
        from_el = anchor.find(f'{{{NS_XDR}}}from')
        row = int(from_el.find(f'{{{NS_XDR}}}row').text) if from_el is not None else 999
        col = int(from_el.find(f'{{{NS_XDR}}}col').text) if from_el is not None else 999
        for el in anchor.iter():
            rid = el.get(f'{{{NS_R}}}id')
            if rid and rid in rid_to_path:
                anchors.append((row, col, rid_to_path[rid]))
                break
    anchors.sort(key=lambda x: (x[0], x[1]))
    return [p for _, _, p in anchors]


def apply_title_fix(xml_bytes, find_str, replace_str):
    xml_str = xml_bytes.decode('utf-8')
    c_title_pat = re.compile(r'(<c:title\b[^>]*>)(.*?)(</c:title>)', re.DOTALL)
    at_pat = re.compile(r'<a:t>(.*?)</a:t>', re.DOTALL)
    changes = []

    def replace_in_title(m):
        content = m.group(2)
        all_text = ''.join(at_pat.findall(content))
        if find_str not in all_text:
            return m.group(0)

        # Case 1: find_str in a single <a:t>
        single_pat = re.compile(
            r'(<a:t>)(.*?)(' + re.escape(find_str) + r')(.*?)(</a:t>)', re.DOTALL)
        new_content, n = single_pat.subn(
            lambda am: am.group(1) + am.group(2) + replace_str + am.group(4) + am.group(5),
            content)
        if n:
            changes.append(f'"{all_text[:65].strip()}"')
            return m.group(1) + new_content + m.group(3)

        # Case 2: find_str split across two <a:t> elements
        for sp in range(1, len(find_str)):
            pa = re.escape(find_str[:sp])
            pb = re.escape(find_str[sp:])
            split_pat = re.compile(
                r'(<a:t>)(.*?)(' + pa + r')(</a:t>)((?:(?!</a:t>).)*?)(<a:t>)(' + pb + r')',
                re.DOTALL)
            new_content, n = split_pat.subn(
                lambda am: (am.group(1) + am.group(2) + replace_str[:sp] +
                            am.group(4) + am.group(5) + am.group(6) + replace_str[sp:]),
                content)
            if n:
                changes.append(f'"{all_text[:65].strip()}"')
                return m.group(1) + new_content + m.group(3)

        return m.group(0)

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

            # chart_path -> visual chart number (7 or 9)
            chart_fixes = {}
            for sheet in tree.getroot().findall(f'.//{ns}sheet'):
                if sheet.get('name') not in TARGET_SHEETS:
                    continue
                rid = sheet.get(rns + 'id')
                sheet_target = wb_rels.get(rid, '')
                drw = get_sheet_drawing(zf, sheet_target)
                if not drw:
                    continue
                ordered = get_charts_in_visual_order(zf, drw)
                if len(ordered) >= 7:
                    chart_fixes[ordered[6]] = 7   # Chart 7 (0-indexed: 6)
                if len(ordered) >= 9:
                    chart_fixes[ordered[8]] = 9   # Chart 9 (0-indexed: 8)

            if not chart_fixes:
                return []

            file_contents = {n: zf.open(n).read() for n in names}

    except Exception as e:
        print(f'  ERROR reading {xlsx_path}: {e}')
        return []

    modified = False
    for chart_path, chart_num in chart_fixes.items():
        key = next((k for k in file_contents
                    if k.replace('\\', '/') == chart_path.replace('\\', '/')), None)
        if key is None:
            continue
        new_bytes, changes = apply_title_fix(file_contents[key], FIND, REPLACE)
        if changes:
            file_contents[key] = new_bytes
            modified = True
            for c in changes:
                all_changes.append(f'Chart{chart_num} [{os.path.basename(chart_path)}]: {c} -> "{REPLACE}"')

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
    xlsx_files = [f for f in glob.glob(os.path.join(SEARCH_DIR, '**', '*.xlsx'), recursive=True)
                  if '~$' not in f]
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
