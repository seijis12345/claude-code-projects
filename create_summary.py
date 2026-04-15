"""
5桁数字で始まる Excel ファイルに Summary P1 / P2 シートを追加し
5桁数字の直後に " Summary_" を挿入したファイル名で保存するスクリプト

  例) 10259 after swap tripler.xls  → 10259 Summary_after swap tripler.xls
      10259_XVDAC.xlsx              → 10259 Summary_XVDAC.xlsx
      10259abc.xlsx                 → 10259 Summary_abc.xlsx

【使い方】
  1. DATA_FILE に対象ファイルのパスを指定して実行
  2. Reference.xlsx は同じフォルダに置くか REF_FILE で指定
"""

import win32com.client
import os
import re

# ════════════════════════════════════════════════════════════════
#  ▼▼▼  ここだけ変更してください  ▼▼▼
DATA_FILE = r'C:\Users\seijis\Desktop\40049 2nd stage.xlsx'
# ▲▲▲  ここだけ変更してください  ▲▲▲
# ════════════════════════════════════════════════════════════════

# Reference.xlsx はスクリプトと同じフォルダ or 直接指定
REF_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        'Reference.xlsx')

# ── 出力ファイル名を生成 ─────────────────────────────────────────
# 5桁数字の後の区切り文字（スペース・_ ・- 等）を1文字だけ除去して "Summary_" を挿入
base = os.path.basename(DATA_FILE)
m = re.match(r'^(\d{5})[ _\-]?(.+)$', base, re.IGNORECASE)
if not m:
    raise ValueError(f"ファイル名が想定形式（5桁数字で始まる）と異なります: {base}")

prefix, rest = m.group(1), m.group(2)
out_name = f"{prefix} Summary_{rest}"
OUT_FILE = os.path.join(os.path.dirname(os.path.abspath(DATA_FILE)), out_name)

print(f"入力ファイル : {DATA_FILE}")
print(f"Referenceファイル: {REF_FILE}")
print(f"出力ファイル : {OUT_FILE}")

# ── Excel COM 起動 ───────────────────────────────────────────────
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible       = False
excel.DisplayAlerts = False

try:
    wb_ref  = excel.Workbooks.Open(os.path.abspath(REF_FILE))
    wb_data = excel.Workbooks.Open(os.path.abspath(DATA_FILE))

    # ── シート名マッピング（Harmonics_N の番号で対応）────────────────
    ref_harmonics  = {s.Name.split('.')[1]: s.Name
                      for s in wb_ref.Sheets if 'Harmonics' in s.Name}
    data_harmonics = {s.Name.split('.')[1]: s.Name
                      for s in wb_data.Sheets if 'Harmonics' in s.Name}

    name_map = {ref_harmonics[k]: data_harmonics[k]
                for k in ref_harmonics if k in data_harmonics}

    print("\nシート名マッピング:")
    for k, v in name_map.items():
        print(f"  {k} → {v}")

    # ── Summary P1 / P2 を先頭にコピー ───────────────────────────
    first_sheet = wb_data.Sheets(1)

    wb_ref.Sheets('summary P2').Copy(Before=first_sheet)
    new_p2 = excel.ActiveSheet
    new_p2.Name = 'summary P2'

    wb_ref.Sheets('summary P1').Copy(Before=new_p2)
    new_p1 = excel.ActiveSheet
    new_p1.Name = 'summary P1'

    # ── AS2:DN2 のシート名を更新 ──────────────────────────────────
    def update_sheet_names(sheet):
        for col in range(45, 119):          # AS(45) 〜 DN(118)
            cell = sheet.Cells(2, col)
            val  = cell.Value
            if val in name_map:
                cell.Value = name_map[val]

    update_sheet_names(new_p1)
    update_sheet_names(new_p2)
    print("AS2:DN2 のシート名を更新しました。")

    # ── 別名で保存 ────────────────────────────────────────────────
    wb_data.SaveAs(os.path.abspath(OUT_FILE),
                   FileFormat=51)           # 51 = xlOpenXMLWorkbook (.xlsx)
    print(f"\n保存完了: {OUT_FILE}")

finally:
    try:
        wb_ref.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        wb_data.Close(SaveChanges=False)
    except Exception:
        pass
    excel.Quit()

print("完了しました。")
