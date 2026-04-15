# Fix chart titles by visual position on sheet:
# Chart3 (3rd chart): '2/3th sub-harmonics' -> '1/3th sub-harmonics'
# Chart4 (4th chart): '1/3th sub-harmonics' -> '2/3th sub-harmonics'
# Target sheets: 'summary P1', 'summary P2', 'Compare P1', 'Compare P2'

import zipfile, os, glob, shutil, re, xml.etree.ElementTree as ET

SEARCH_DIR    = r'C:/Users/seijis/Desktop/04_MMA_Data_Pmax_Harmonics'
TARGET_SHEETS = {'summary P1', 'summary P2', 'Compare P1', 'Compare P2'}

NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


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
    """Return list of chart xml paths sorted by visual position (row, col)."""
    drw_key  = 'xl/' + drawing_path
    rels_key = f'xl/{drawing_path.rsplit("/",1)[0]}/_rels/{drawing_path.rsplit("/",1)[1]}.rels'
    with zf.open(rels_key) as f:
        rtree = ET.parse(f)
    rid_to_path = {}
    for rel in rtree.getroot():
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
    return [path for _, _, path in anchors]


def get_chart_main_title(xml_bytes):
    xml_str = xml_bytes.decode('utf-8')
    titles = re.findall(r'<c:title\b[^>]*>.*?</c:title>', xml_str, re.DOTALL)
    if not titles:
        return ''
    return ''.join(re.findall(r'<a:t>(.*?)</a:t>', titles[0], re.DOTALL))


def apply_title_fix(xml_bytes, find_str, replace_str):
    """
    Replace find_str->replace_str within c:title blocks.
    Handles both single-element and split-element cases.
    """
    xml_str = xml_bytes.decode('utf-8')
    c_title_pat = re.compile(r'(<c:title\b[^>]*>)(.*?)(</c:title>)', re.DOTALL)
    at_pat = re.compile(r'<a:t>(.*?)</a:t>', re.DOTALL)
    changes = []

    def replace_in_title(m):
        content = m.group(2)
        all_text = ''.join(at_pat.findall(content))
        if find_str not in all_text:
            return m.group(0)

        # Case 1: find_str is entirely within one <a:t> element
        single_pat = re.compile(r'(<a:t>)(.*?)(' + re.escape(find_str) + r')(.*?)(</a:t>)', re.DOTALL)
        new_content, n = single_pat.subn(
            lambda am: am.group(1) + am.group(2) + replace_str + am.group(4) + am.group(5),
            content
        )
        if n:
            changes.append(f'"{all_text[:60].strip()}"')
            return m.group(1) + new_content + m.group(3)

        # Case 2: find_str spans multiple <a:t> elements (split at arbitrary point)
        # find_str = A + B where A is tail of one a:t and B is start of next
        for split_pos in range(1, len(find_str)):
            part_a = re.escape(find_str[:split_pos])
            part_b = re.escape(find_str[split_pos:])
            split_pat = re.compile(
                r'(<a:t>)(.*?)(' + part_a + r')(</a:t>)((?:(?!</a:t>).)*?)(<a:t>)(' + part_b + r')',
                re.DOTALL
            )
            new_content, n = split_pat.subn(
                lambda am: (am.group(1) + am.group(2) + replace_str[:split_pos] +
                            am.group(4) + am.group(5) + am.group(6) + replace_str[split_pos:]),
                content
            )
            if n:
                changes.append(f'"{all_text[:60].strip()}"')
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

            # Map: chart_path -> (sheet_name, visual_position_1indexed)
            chart_fixes = {}  # chart_path -> (find, replace)
            for sheet in tree.getroot().findall(f'.//{ns}sheet'):
                if sheet.get('name') not in TARGET_SHEETS:
                    continue
                rid = sheet.get(rns + 'id')
                sheet_target = wb_rels.get(rid, '')
                drw = get_sheet_drawing(zf, sheet_target)
                if not drw:
                    continue
                ordered = get_charts_in_visual_order(zf, drw)
                if len(ordered) >= 3:
                    chart_fixes[ordered[2]] = ('2/3th sub-harmonics', '1/3th sub-harmonics')  # Chart3
                if len(ordered) >= 4:
                    chart_fixes[ordered[3]] = ('1/3th sub-harmonics', '2/3th sub-harmonics')  # Chart4

            if not chart_fixes:
                return []

            file_contents = {n: zf.open(n).read() for n in names}

    except Exception as e:
        print(f'  ERROR reading {xlsx_path}: {e}')
        return []

    modified = False
    for chart_path, (find_str, replace_str) in chart_fixes.items():
        key = next((k for k in file_contents if k.replace('\\', '/') == chart_path.replace('\\', '/')), None)
        if key is None:
            continue
        new_bytes, changes = apply_title_fix(file_contents[key], find_str, replace_str)
        if changes:
            file_contents[key] = new_bytes
            modified = True
            chart_num = 3 if replace_str == '1/3th sub-harmonics' else 4
            for c in changes:
                all_changes.append(f'Chart{chart_num} [{os.path.basename(chart_path)}]: {c} -> "{replace_str}"')

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
    xlsx_files = [f for f in xlsx_files if '~$' not in f]
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
