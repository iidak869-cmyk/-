"""
業務改善指示書 Excel生成スクリプト

Usage:
    python generate_kaizen.py --date "2026年7月29日" \
        --incident "..." --risk "..." --cause "..." \
        --improvement "..." --action "..." \
        --output "業務改善指示書_20260729.xlsx"
"""

import argparse
import sys
from datetime import date
import openpyxl
from openpyxl.styles import (
    Font, Fill, PatternFill, Alignment, Border, Side, Color
)
from openpyxl.utils import get_column_letter


ORANGE_FILL = PatternFill(start_color="FBE4D5", end_color="FBE4D5", fill_type="solid")
GRAY_FILL   = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
THIN_BORDER = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000'),
)

SECTION_LABELS = {
    1: "① 認知されたミス・事例（起きた具体的事象を記載）・・・発生日、発生内容、発見経緯、誰が関与したか、顧客・従業員への影響有無など事実ベースで記載してください。",
    2: "②想定リスク・・・この問題から考えられる影響範囲や波及リスクを具体的に記載してください。顧客・従業員への信頼毀損、損失金額、契約不履行、業務遅延などに加え、同様の案件が存在しているか、見落としが他にも潜在していないかも想定のうえ記載してください。",
    3: "③ 問題点（要因分析）・・・このミスの背景にある仕組み・体制・習慣・情報共有の不備など、根本的な要因まで想定して記載してください。「人的ミス」「確認不足」だけで片づけず、再発の温床になりうる構造にも目を向けてください。",
    4: "④ 改善策・・・問題が起きた事象に対する対処だけでなく、運用手順や仕組み、マニュアル、教育内容の見直しなど実務レベルで持続的に改善できる範囲まで記載してください。短期対応で終わらず、現場内で再発防止が図れるよう検討してください。",
    5: "⑤ 対策・行動・・・仕組みや方針そのものを問い直す必要がある、全社的または部門横断的な課題への対応方針を記載してください。経営層や他部門の調整を要する難易度の高い対策も含め、長期的に再発を防ぐ視点で検討してください。",
}


def set_cell(ws, cell_ref, value, bold=False, size=9, wrap=True,
             h_align='left', v_align='top', fill=None, font_color='000000',
             border=True):
    cell = ws[cell_ref]
    cell.value = value
    cell.font = Font(name='メイリオ', bold=bold, size=size, color=font_color)
    cell.alignment = Alignment(wrap_text=wrap, horizontal=h_align, vertical=v_align)
    if fill:
        cell.fill = fill
    if border:
        cell.border = THIN_BORDER


def create_kaizen_excel(creation_date, incident, risk, cause, improvement, action, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Column widths — spread across J columns (A to J)
    col_widths = {'A': 8.86, 'B': 10, 'C': 10, 'D': 6, 'E': 10,
                  'F': 10, 'G': 10, 'H': 10, 'I': 10, 'J': 10}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Row 1: タイトル
    ws.merge_cells('A1:J1')
    ws['A1'].value = '業務改善指示書'
    ws['A1'].font = Font(name='メイリオ', bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 24

    # Row 2: メタ情報
    ws.row_dimensions[2].height = 18
    ws.merge_cells('B2:C2')
    ws.merge_cells('F2:G2')
    ws.merge_cells('I2:J2')

    set_cell(ws, 'A2', '作成日', h_align='center', v_align='center', border=True)
    set_cell(ws, 'B2', creation_date, h_align='center', v_align='center', border=True)
    set_cell(ws, 'D2', '作成者', h_align='center', v_align='center', border=True)
    set_cell(ws, 'E2', 'デザイン課', h_align='center', v_align='center', fill=GRAY_FILL, border=True)
    set_cell(ws, 'F2', '飯田圭祐', h_align='left', v_align='center', fill=GRAY_FILL, border=True)
    set_cell(ws, 'H2', '上長サイン', h_align='center', v_align='center', border=True)
    # Apply border to merged cells manually
    for col in ['C2', 'G2', 'I2', 'J2']:
        ws[col].border = THIN_BORDER

    # Section 1: ① 認知されたミス・事例
    ws.merge_cells('A3:J3')
    ws.row_dimensions[3].height = 31.5
    set_cell(ws, 'A3', SECTION_LABELS[1], bold=True, size=9,
             h_align='left', v_align='center', fill=ORANGE_FILL, border=True)

    ws.merge_cells('A4:J10')
    ws.row_dimensions[4].height = 120
    set_cell(ws, 'A4', incident, bold=False, size=9,
             h_align='left', v_align='top', border=True)

    # Section 2: ② 想定リスク
    ws.merge_cells('A11:J12')
    ws.row_dimensions[11].height = 12.75
    ws.row_dimensions[12].height = 34.5
    set_cell(ws, 'A11', SECTION_LABELS[2], bold=True, size=8,
             h_align='left', v_align='center', fill=ORANGE_FILL, border=True)

    ws.merge_cells('A13:J16')
    ws.row_dimensions[13].height = 18
    ws.row_dimensions[14].height = 18
    ws.row_dimensions[15].height = 18
    ws.row_dimensions[16].height = 60
    set_cell(ws, 'A13', risk, bold=False, size=9,
             h_align='left', v_align='top', border=True)

    # Section 3: ③ 問題点
    ws.merge_cells('A17:J17')
    ws.row_dimensions[17].height = 30
    set_cell(ws, 'A17', SECTION_LABELS[3], bold=True, size=9,
             h_align='left', v_align='center', fill=ORANGE_FILL, border=True)

    ws.merge_cells('A18:J21')
    ws.row_dimensions[18].height = 18
    ws.row_dimensions[19].height = 18
    ws.row_dimensions[20].height = 18
    ws.row_dimensions[21].height = 60
    set_cell(ws, 'A18', cause, bold=False, size=9,
             h_align='left', v_align='top', border=True)

    # Section 4: ④ 改善策 (red text)
    ws.merge_cells('A22:J22')
    ws.row_dimensions[22].height = 48
    set_cell(ws, 'A22', SECTION_LABELS[4], bold=True, size=9,
             h_align='left', v_align='center', fill=ORANGE_FILL,
             font_color='FF0000', border=True)

    ws.merge_cells('A23:J25')
    ws.row_dimensions[23].height = 18
    ws.row_dimensions[24].height = 18
    ws.row_dimensions[25].height = 60
    set_cell(ws, 'A23', improvement, bold=False, size=9,
             h_align='left', v_align='top', border=True)

    # Section 5: ⑤ 対策・行動 (red text)
    ws.merge_cells('A26:J26')
    ws.row_dimensions[26].height = 59.25
    set_cell(ws, 'A26', SECTION_LABELS[5], bold=True, size=9,
             h_align='left', v_align='center', fill=ORANGE_FILL,
             font_color='FF0000', border=True)

    ws.merge_cells('A27:J30')
    ws.row_dimensions[27].height = 32.25
    ws.row_dimensions[28].height = 18
    ws.row_dimensions[29].height = 18
    ws.row_dimensions[30].height = 18
    set_cell(ws, 'A27', action, bold=False, size=9,
             h_align='left', v_align='top', border=True)

    # Print area and page setup
    ws.print_area = 'A1:J30'
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    wb.save(output_path)
    print(f"Saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='業務改善指示書 Excel生成')
    parser.add_argument('--date', required=True, help='作成日 (例: 2026年7月29日)')
    parser.add_argument('--incident', required=True, help='①発生事象')
    parser.add_argument('--risk', required=True, help='②想定リスク')
    parser.add_argument('--cause', required=True, help='③問題点・要因')
    parser.add_argument('--improvement', required=True, help='④改善策')
    parser.add_argument('--action', required=True, help='⑤対策・行動')
    parser.add_argument('--output', required=True, help='出力ファイルパス')
    args = parser.parse_args()

    create_kaizen_excel(
        creation_date=args.date,
        incident=args.incident,
        risk=args.risk,
        cause=args.cause,
        improvement=args.improvement,
        action=args.action,
        output_path=args.output,
    )


if __name__ == '__main__':
    main()
