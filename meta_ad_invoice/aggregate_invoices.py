"""
Meta広告 請求書集計スクリプト（テスト環境）

処理フロー:
  1. 指定フォルダ内のPDF請求書を1件ずつ読み取り、PDFごとにExcelファイルを作成する
  2. 作成したExcelファイルを一時保存フォルダへ保存する
  3. 一時保存フォルダのExcelファイルを1件ずつ確認し、集計用Excel
     （G:\\共有ドライブ\\RPA運用-技術課連携\\〇月請求分.xlsx）へ以下の項目を転記する
       - キャンペーン名 / 日時 / 消化金額 / 参照番号 / クレジットカード番号 / アカウント名
  4. データが無い、または必要項目を正しく取得できなかったExcelファイルは
     「白紙」フォルダへ移動する

実行環境:
  ユーザーのWindows端末（C:\\、G:\\共有ドライブへアクセスできる端末）で実行する想定。

事前準備:
  pip install -r requirements.txt

実行例:
  python aggregate_invoices.py ^
      --source "C:\\Users\\iida869\\Downloads\\202605_CS運用・TRANSCEND＆ValueCreationRoom＆株式会社Growth canvas＆Hackeer＆ピラティス" ^
      --temp-dir "C:\\Users\\iida869\\Downloads\\_一時保存" ^
      --billing-month 6

  --rpa-dir と --blank-dir は既定値のままでよければ省略可（README参照）。

注意:
  Meta広告請求書PDFのレイアウトは請求方式・言語設定・時期によって表記ゆれがある。
  実際のPDFで転記結果を確認し、必要に応じて HEADER_PATTERNS / LINE_ITEM_PATTERN を調整すること。
  --debug 指定でPDFから抽出した生テキストを一時保存フォルダに書き出せるので、
  パターン調整時はそちらを参照する。
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import dataclass, fields
from pathlib import Path

import pdfplumber
from openpyxl import Workbook, load_workbook

# 集計用Excel・per-PDF Excelの列順
FIELD_ORDER = ["キャンペーン名", "日時", "消化金額", "参照番号", "クレジットカード番号", "アカウント名"]
AGGREGATE_HEADERS = ["元PDFファイル名", *FIELD_ORDER]

# PDFテキストから請求書ヘッダー情報（1PDFにつき1つ）を拾う正規表現。
# 表記ゆれに備えて複数パターンを許容し、最初にマッチしたものを採用する。
HEADER_PATTERNS: dict[str, list[str]] = {
    "参照番号": [
        r"参照番号[:\s：]*([A-Za-z0-9\-]{6,})",
        r"Reference\s*(?:Number|No\.?)[:\s]*([A-Za-z0-9\-]{6,})",
    ],
    "クレジットカード番号": [
        r"(?:クレジットカード番号|カード番号|お支払い方法)[:\s：]*([0-9Xx\*・\- ]{4,})",
        r"(?:Payment method|Card number)[:\s]*([0-9Xx\*\- ]{4,})",
    ],
    "アカウント名": [
        r"(?:広告アカウント名|アカウント名)[:\s：]*(.+)",
        r"Account name[:\s]*(.+)",
    ],
    "日時": [
        r"(?:請求日|明細期間|請求期間)[:\s：]*(.+)",
        r"(?:Invoice date|Billing period)[:\s]*(.+)",
    ],
}

# 明細（キャンペーンごとの行）を拾う正規表現。1行に「キャンペーン名 ... 金額円」が
# 並ぶ形式を想定。表形式のPDFでは extract_table() を優先し、これはフォールバック。
# 「円」を必須にすることで、参照番号やカード番号などの末尾が数字の行を誤検出しないようにする。
LINE_ITEM_PATTERN = re.compile(
    r"^(?P<campaign>.+?)\s+(?P<amount>[\d,]+\s*円)\s*$"
)
TABLE_HEADER_PATTERN = re.compile(r"キャンペーン.*(?:消化金額|金額)")


@dataclass
class InvoiceRow:
    キャンペーン名: str = ""
    日時: str = ""
    消化金額: str = ""
    参照番号: str = ""
    クレジットカード番号: str = ""
    アカウント名: str = ""

    def is_complete(self) -> bool:
        return all(getattr(self, f.name).strip() for f in fields(self))

    def as_row(self) -> list[str]:
        return [getattr(self, name) for name in FIELD_ORDER]


def extract_header_fields(text: str) -> dict[str, str]:
    result = {"参照番号": "", "クレジットカード番号": "", "アカウント名": "", "日時": ""}
    for key, patterns in HEADER_PATTERNS.items():
        for pattern in patterns:
            m = re.search(pattern, text)
            if m:
                result[key] = m.group(1).strip()
                break
    return result


def extract_line_items(pdf: "pdfplumber.PDF") -> list[tuple[str, str]]:
    """(キャンペーン名, 消化金額) の一覧を、表があれば表から、無ければテキストから抽出する。"""
    items: list[tuple[str, str]] = []

    for page in pdf.pages:
        for table in page.extract_tables() or []:
            for row in table:
                cells = [c.strip() for c in row if c]
                if len(cells) < 2:
                    continue
                # 「キャンペーン」列と「金額」列らしきものを両端から推測する
                name_cell = cells[0]
                amount_cell = next(
                    (c for c in reversed(cells) if re.search(r"[\d,]+\s*円?$", c)),
                    None,
                )
                if name_cell and amount_cell and not name_cell.startswith("キャンペーン名"):
                    items.append((name_cell, amount_cell))

    if items:
        return items

    # 表が取れなかった場合はテキスト行から拾う。
    # 「キャンペーン名 消化金額」のような見出し行より後ろだけを明細候補として扱う。
    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.splitlines()
        header_idx = next((i for i, line in enumerate(lines) if TABLE_HEADER_PATTERN.search(line)), None)
        candidate_lines = lines[header_idx + 1:] if header_idx is not None else lines
        for line in candidate_lines:
            m = LINE_ITEM_PATTERN.match(line.strip())
            if m:
                items.append((m.group("campaign").strip(), m.group("amount").strip()))

    return items


def extract_invoice_rows(pdf_path: Path) -> list[InvoiceRow]:
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        header = extract_header_fields(full_text)
        line_items = extract_line_items(pdf)

    if not line_items:
        # 明細が拾えなくてもヘッダー情報だけの1行として返す（あとで不備判定される）
        return [InvoiceRow(**header)]

    rows = []
    for campaign, amount in line_items:
        rows.append(InvoiceRow(
            キャンペーン名=campaign,
            消化金額=amount,
            日時=header["日時"],
            参照番号=header["参照番号"],
            クレジットカード番号=header["クレジットカード番号"],
            アカウント名=header["アカウント名"],
        ))
    return rows


def write_pdf_excel(rows: list[InvoiceRow], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "請求書データ"
    ws.append(FIELD_ORDER)
    for row in rows:
        ws.append(row.as_row())
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def read_pdf_excel(xlsx_path: Path) -> list[InvoiceRow]:
    wb = load_workbook(xlsx_path)
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    rows = []
    for excel_row in ws.iter_rows(min_row=2, values_only=True):
        if excel_row is None or all(v is None for v in excel_row):
            continue
        record = dict(zip(header, excel_row))
        rows.append(InvoiceRow(**{k: (record.get(k) or "") for k in FIELD_ORDER}))
    return rows


def load_or_create_aggregate(path: Path) -> Workbook:
    if path.exists():
        return load_workbook(path)
    wb = Workbook()
    ws = wb.active
    ws.title = "請求集計"
    ws.append(AGGREGATE_HEADERS)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return wb


def dedup_key(row: InvoiceRow) -> tuple[str, str, str]:
    """1件の請求書に複数キャンペーンの明細が含まれるため、参照番号だけでなく
    キャンペーン名・消化金額も合わせて重複判定キーとする。"""
    return (row.参照番号, row.キャンペーン名, row.消化金額)


def existing_dedup_keys(ws) -> set[tuple[str, str, str]]:
    idx = {name: AGGREGATE_HEADERS.index(name) for name in ("参照番号", "キャンペーン名", "消化金額")}
    return {
        (row[idx["参照番号"]], row[idx["キャンペーン名"]], row[idx["消化金額"]])
        for row in ws.iter_rows(min_row=2, values_only=True)
        if row and row[idx["参照番号"]]
    }


def step1_convert_pdfs_to_excel(source_dir: Path, temp_dir: Path, debug: bool) -> list[Path]:
    pdf_paths = sorted(source_dir.glob("*.pdf"))
    if not pdf_paths:
        print(f"[警告] PDFが見つかりません: {source_dir}", file=sys.stderr)

    created: list[Path] = []
    for pdf_path in pdf_paths:
        try:
            rows = extract_invoice_rows(pdf_path)
        except Exception as e:
            print(f"[エラー] PDF読み取り失敗: {pdf_path.name} ({e})", file=sys.stderr)
            rows = [InvoiceRow()]

        out_path = temp_dir / f"{pdf_path.stem}.xlsx"
        write_pdf_excel(rows, out_path)
        created.append(out_path)
        print(f"作成: {out_path.name}")

        if debug:
            with pdfplumber.open(pdf_path) as pdf:
                raw_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            debug_path = temp_dir / f"{pdf_path.stem}.debug.txt"
            debug_path.write_text(raw_text, encoding="utf-8")

    return created


def step3_transcribe_and_sort(
    excel_paths: list[Path],
    aggregate_path: Path,
    blank_dir: Path,
) -> tuple[int, int]:
    wb = load_or_create_aggregate(aggregate_path)
    ws = wb.active
    known_keys = existing_dedup_keys(ws)

    blank_dir.mkdir(parents=True, exist_ok=True)
    ok_count = 0
    ng_count = 0

    for xlsx_path in excel_paths:
        rows = read_pdf_excel(xlsx_path)
        if not rows or not all(r.is_complete() for r in rows):
            dest = blank_dir / xlsx_path.name
            shutil.move(str(xlsx_path), str(dest))
            print(f"[白紙移動] 必要項目が不足: {xlsx_path.name}")
            ng_count += 1
            continue

        for row in rows:
            key = dedup_key(row)
            if key in known_keys:
                continue  # 再実行時の重複転記を防止
            ws.append([xlsx_path.name, *row.as_row()])
            known_keys.add(key)

        ok_count += 1
        print(f"[転記完了] {xlsx_path.name}")

    wb.save(aggregate_path)
    return ok_count, ng_count


def build_aggregate_path(rpa_dir: Path, billing_month: int) -> Path:
    return rpa_dir / f"{billing_month}月請求分.xlsx"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", required=True, type=Path, help="PDF請求書が入っているフォルダ")
    parser.add_argument("--temp-dir", required=True, type=Path, help="PDFから作成したExcelの一時保存先")
    parser.add_argument(
        "--rpa-dir",
        type=Path,
        default=Path(r"G:\共有ドライブ\RPA運用-技術課連携"),
        help="集計用Excel・白紙フォルダが置かれている共有ドライブ上のフォルダ",
    )
    parser.add_argument("--billing-month", required=True, type=int, help="対象請求月（例: 6）")
    parser.add_argument(
        "--blank-dir",
        type=Path,
        default=None,
        help="不備ファイルの移動先（省略時は <rpa-dir>/白紙）",
    )
    parser.add_argument("--debug", action="store_true", help="PDFの抽出テキストを一時保存先に書き出す")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    if not args.source.is_dir():
        print(f"[エラー] 指定フォルダが存在しません: {args.source}", file=sys.stderr)
        return 1

    args.temp_dir.mkdir(parents=True, exist_ok=True)
    blank_dir = args.blank_dir or (args.rpa_dir / "白紙")
    aggregate_path = build_aggregate_path(args.rpa_dir, args.billing_month)

    print(f"1-2. PDF → Excel変換（保存先: {args.temp_dir}）")
    excel_paths = step1_convert_pdfs_to_excel(args.source, args.temp_dir, args.debug)

    print(f"3-4. 集計用Excelへ転記（{aggregate_path}）／不備ファイルは {blank_dir} へ移動")
    ok_count, ng_count = step3_transcribe_and_sort(excel_paths, aggregate_path, blank_dir)

    print(f"完了: 転記{ok_count}件 / 白紙移動{ng_count}件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
