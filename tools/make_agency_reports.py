"""販売代理店20社の販売報告（Issue #15）のサンプルを作る。仕様は docs/sample-data.md §2.1・§6。8.3 の題材。

代理店から毎月届く Excel の販売報告の、2026年9月分（本文の「今日」は2026年10月）。
中身は売上明細（tools/make_sales.py）の「2026年9月・販売経路＝代理店」の行で、代理店ごとに1ファイルにする。
会社が配った様式どおりのファイルに、次のずれと重複を混ぜる（どの代理店に何を入れるかは下の定数で決める）。

- 列の順番が違う（COLUMN_ORDERS）
- 見出しの行が1行下にずれている（HEADER_SHIFTED。上に一言書き足してある）
- 部署名を「大阪」のように略して書く（DEPARTMENT_ALIASES。§2.1 の「代理店の報告での書き方」）
- 再送：同じ報告をもう一度送ってきた（RESENT。「_再送」のファイル。中身は元と同じ）
- 訂正：元の報告に漏れがあり、全件を書き直して送ってきた（CORRECTED。「_訂正」のファイル）
- 同じファイルの中で1行を2回貼っている（IN_FILE_DUPLICATE）

9月に代理店経由の売上がない代理店（栞門産業株式会社）は、見出しだけの空の報告を送ってくる。これは意図したもの
（「売上なし」の報告も届く。統合しても行は増えない）。

統合と重複除去の正解（売上明細の該当行と一致する）を samples/answers/ に出す。統合のルール（merge_reports）：

- 訂正版がある代理店は、元のファイル（と、その再送）を丸ごと捨てて訂正版を採る（実務の差し替え。訂正版の注記
  「こちらに差し替えをお願いします」に合わせる）。重複した行だけを消すのではない
- 残ったファイルをつなぎ、部署名を正式名に戻して、同じ行を1行にする（再送・同じファイルの中の重複）
- 除いた行は「除いた重複行」シートに、元のファイル名と理由（訂正版に差し替え／再送／同じ行の重複）を付けて載せる。
  統合後の行＋除いた行＝統合前の行（件数・金額とも）

乱数は固定シードなので、何度実行しても同じ内容になる。一時フォルダに保存して個人情報を取り除き、残っていないことを
確かめてから samples/ に移す（tools/staging.py）。

    uv run python tools/make_agency_reports.py
"""
import datetime as dt
import random
import sys
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import make_sales
import sample_master as master
import staging
from xlsx_meta import scrub_personal_info, verify_no_personal_info

YEAR, MONTH = 2026, 9
SAMPLES = Path(__file__).resolve().parent.parent / "samples"
OUTPUT_DIR = SAMPLES / f"代理店報告_{YEAR}年{MONTH}月"
ANSWER = SAMPLES / "answers" / f"代理店報告_{YEAR}年{MONTH}月_統合.xlsx"
SEED = 20260930

SHEET_NAME = "販売報告"
TITLE = f"販売報告書（{master.COMPANY_NAME} 御中）"
# 会社が配った様式：1〜3行目に表題・代理店名・対象月、4行目は空き、5行目が見出し、6行目から明細
STANDARD_COLUMNS = ("販売日", "伝票番号", "部署名", "得意先名", "商品コード", "商品名", "数量", "単価", "金額")
STANDARD_HEADER_ROW = 5
SHIFT_NOTE = "いつもお世話になっております。9月分の販売実績をご報告します。"
NOTES = {"再送": "※再送（先日お送りした報告と同じものです）",
         "訂正": "※訂正版（先日の報告に漏れがありました。こちらに差し替えをお願いします）"}

_AGENCY = {i: a.name for i, a in enumerate(master.AGENCIES)}
# 列の順番が違う代理店
COLUMN_ORDERS = {
    _AGENCY[5]: ("販売日", "伝票番号", "得意先名", "部署名", "商品名", "商品コード", "数量", "単価", "金額"),
    _AGENCY[11]: ("伝票番号", "販売日", "部署名", "得意先名", "商品コード", "商品名", "金額", "数量", "単価"),
    _AGENCY[19]: ("得意先名", "販売日", "伝票番号", "商品コード", "商品名", "数量", "単価", "金額", "部署名"),
}
# 見出しの行が1行下にずれている代理店
HEADER_SHIFTED = frozenset({_AGENCY[1], _AGENCY[9], _AGENCY[14]})
# 部署名を略して書く代理店と、その書き方（§2.1）
DEPARTMENT_ALIASES = {
    _AGENCY[0]: "東京", _AGENCY[6]: "東北", _AGENCY[8]: "大阪", _AGENCY[12]: "中部", _AGENCY[19]: "海外",
}
# 同じ報告を再送してきた代理店
RESENT = frozenset({_AGENCY[8], _AGENCY[15]})
# 元の報告に漏れがあり、訂正版（全件）を送ってきた代理店と、元の報告で漏れていた行数。
# 訂正版は様式どおりに書き直してある（部署名も正式名）。茜舟商会は元の報告が略称なので、
# 部署名をそろえてからでないと、元の報告と訂正版の行が重複に見えない
CORRECTED = {_AGENCY[0]: 2, _AGENCY[14]: 1}
# 同じファイルの中で1行を2回貼っている代理店
IN_FILE_DUPLICATE = frozenset({_AGENCY[19]})

_ALIAS_TO_DEPARTMENT = {alias: d.name for d in master.SALES_DEPARTMENTS for alias in d.agency_report_names}


def _check_constants() -> None:
    """定数の代理店名・略称・列名がマスタと様式に合っていることを確かめる。違えば ValueError。"""
    agency_dept = {a.name: a.department for a in master.AGENCIES}
    named = [*COLUMN_ORDERS, *HEADER_SHIFTED, *DEPARTMENT_ALIASES, *RESENT, *CORRECTED, *IN_FILE_DUPLICATE]
    unknown = [n for n in named if n not in agency_dept]
    if unknown:
        raise ValueError(f"マスタにない代理店: {unknown}")
    for agency, alias in DEPARTMENT_ALIASES.items():
        if _ALIAS_TO_DEPARTMENT.get(alias) != agency_dept[agency]:
            raise ValueError(f"{agency}: 略称 {alias!r} は担当部署 {agency_dept[agency]} の書き方ではない")
    for agency, columns in COLUMN_ORDERS.items():
        if sorted(columns) != sorted(STANDARD_COLUMNS):
            raise ValueError(f"{agency}: 列の名前が様式と違う: {columns}")


_check_constants()  # 起動時に確かめる


class ReportLine(NamedTuple):
    """報告の明細1行（様式の列の順番）。department は報告に書かれたとおり（略称のことがある）。"""
    date: dt.date
    slip_no: str
    department: str
    customer: str
    product_code: str
    product_name: str
    quantity: int
    unit_price: int
    amount: int

    def by_column(self) -> dict[str, object]:
        return dict(zip(STANDARD_COLUMNS, self))


class Report(NamedTuple):
    """1つの報告ファイル。"""
    file_name: str
    agency: str
    kind: str                  # 通常 / 再送 / 訂正
    columns: tuple[str, ...]   # 見出しの並び
    header_row: int            # 見出しの行番号（1始まり）
    lines: tuple[ReportLine, ...]
    note: str | None           # 見出しより上に書き足した一言


class MergedRow(NamedTuple):
    """統合後の1行。部署名は販売管理システム上の名前に戻してある。"""
    agency: str
    date: dt.date
    slip_no: str
    department: str
    customer: str
    product_code: str
    product_name: str
    quantity: int
    unit_price: int
    amount: int


ANSWER_HEADER = ("代理店名", *STANDARD_COLUMNS)
REMOVED_HEADER = ("元のファイル", "理由", *ANSWER_HEADER)

# 除いた理由
REASON_REPLACED = "訂正版に差し替え"
REASON_RESENT = "再送"
REASON_DUPLICATE = "同じ行の重複"
REASON_OTHER_FILE = "別のファイルと重複"


class RemovedRow(NamedTuple):
    """統合で除いた1行と、その行があったファイル・除いた理由。"""
    file_name: str
    reason: str
    row: MergedRow


def _file_name(agency: str, kind: str) -> str:
    suffix = "" if kind == "通常" else f"_{kind}"
    return f"{agency}_{YEAR}年{MONTH}月{suffix}.xlsx"


def _line(row: make_sales.SalesRow, department: str) -> ReportLine:
    return ReportLine(row.date, row.slip_no, department, row.customer, row.product_code,
                      row.product_name, row.quantity, row.unit_price, row.amount)


def _check_row_counts(rows_by_agency: dict[str, list]) -> None:
    """訂正版・同じ行の重複を仕込む代理店に、必要な行数があることを確かめる。足りなければ ValueError。

    訂正版の代理店は、元の報告から CORRECTED の行数を抜いても1行以上残る必要がある。
    同じ行を2回貼る代理店は、1行以上の売上が要る。
    """
    problems = []
    for agency, omitted in CORRECTED.items():
        n = len(rows_by_agency.get(agency, []))
        if n <= omitted:
            problems.append(f"{agency}: 訂正版で漏れていた行 {omitted} 行に対して、{YEAR}年{MONTH}月の売上が {n} 行しかない")
    for agency in IN_FILE_DUPLICATE:
        n = len(rows_by_agency.get(agency, []))
        if n < 1:
            problems.append(f"{agency}: 同じ行を2回貼る代理店なのに、{YEAR}年{MONTH}月の売上が {n} 行")
    if problems:
        raise ValueError("代理店の売上の行数が足りない（定数か売上明細のシードを見直す）: " + "; ".join(problems))


def build_reports(sales_rows: list[make_sales.SalesRow], seed: int = SEED) -> list[Report]:
    """報告ファイルの中身を作る。代理店の順に、通常の報告のあとに再送・訂正版を置く。"""
    rng = random.Random(seed)
    month_rows = [r for r in sales_rows if r.channel == "代理店" and (r.date.year, r.date.month) == (YEAR, MONTH)]
    rows_by_agency = {a.name: [r for r in month_rows if r.agency == a.name] for a in master.AGENCIES}
    _check_row_counts(rows_by_agency)
    reports: list[Report] = []
    for agency in master.AGENCIES:
        rows = rows_by_agency[agency.name]
        dept = DEPARTMENT_ALIASES.get(agency.name, agency.department)
        lines = [_line(r, dept) for r in rows]
        if agency.name in CORRECTED:
            omitted = set(rng.sample(range(len(lines)), CORRECTED[agency.name]))
            lines = [line for i, line in enumerate(lines) if i not in omitted]
        if agency.name in IN_FILE_DUPLICATE:
            i = rng.randrange(len(lines))
            lines.insert(i + 1, lines[i])  # 同じ行を続けて2回貼った
        columns = COLUMN_ORDERS.get(agency.name, STANDARD_COLUMNS)
        shifted = agency.name in HEADER_SHIFTED
        header_row = STANDARD_HEADER_ROW + 1 if shifted else STANDARD_HEADER_ROW
        regular = Report(_file_name(agency.name, "通常"), agency.name, "通常", columns, header_row,
                         tuple(lines), SHIFT_NOTE if shifted else None)
        reports.append(regular)
        if agency.name in RESENT:
            reports.append(regular._replace(file_name=_file_name(agency.name, "再送"), kind="再送"))
        if agency.name in CORRECTED:
            reports.append(Report(_file_name(agency.name, "訂正"), agency.name, "訂正", STANDARD_COLUMNS,
                                  STANDARD_HEADER_ROW, tuple(_line(r, agency.department) for r in rows), None))
    return reports


def _merged_row(report: Report, line: ReportLine) -> MergedRow:
    department = _ALIAS_TO_DEPARTMENT.get(line.department, line.department)
    return MergedRow(report.agency, *line._replace(department=department))


def merge_reports(reports: list[Report]) -> tuple[list[MergedRow], list[RemovedRow]]:
    """報告を統合する。訂正版がある代理店は元のファイルを丸ごと差し替え、残りは同じ行を1行にする。

    1. 訂正版がある代理店の、訂正版以外のファイル（元の報告・その再送）の行は、全部「訂正版に差し替え」で除く
    2. 残ったファイルをファイル名の順に読み、部署名を正式名に戻して、同じ行は先に出てきたものを残す。
       同じファイルの中で2回目なら「同じ行の重複」、再送のファイルなら「再送」で除く
    戻り値は（残した行を販売日・伝票番号・商品コードの順に並べたもの, 除いた行をファイル名の順に並べたもの）。
    """
    corrected = {r.agency for r in reports if r.kind == "訂正"}
    seen: dict[MergedRow, str] = {}  # 行 → 最初に出てきたファイル名
    kept: list[MergedRow] = []
    removed: list[RemovedRow] = []
    for report in sorted(reports, key=lambda r: r.file_name):
        for line in report.lines:
            row = _merged_row(report, line)
            if report.agency in corrected and report.kind != "訂正":
                removed.append(RemovedRow(report.file_name, REASON_REPLACED, row))
            elif row in seen:
                if seen[row] == report.file_name:
                    reason = REASON_DUPLICATE
                elif report.kind == "再送":
                    reason = REASON_RESENT
                else:
                    reason = REASON_OTHER_FILE
                removed.append(RemovedRow(report.file_name, reason, row))
            else:
                seen[row] = report.file_name
                kept.append(row)
    kept.sort(key=lambda r: (r.date, r.slip_no, r.product_code))
    return kept, removed


def _new_book(app: xw.App) -> xw.Book:
    book = app.books.add()
    while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
        book.sheets[-1].delete()
    return book


def _format_table(sheet: xw.Sheet, columns: tuple[str, ...], header_row: int, n_rows: int) -> None:
    last_col = len(columns)
    header = sheet.range((header_row, 1), (header_row, last_col))
    header.font.bold = True
    header.color = (221, 235, 247)
    if n_rows:
        for i, name in enumerate(columns, start=1):
            cells = sheet.range((header_row + 1, i), (header_row + n_rows, i))
            if name == "販売日":
                cells.number_format = "yyyy/mm/dd"
            elif name in ("数量", "単価", "金額"):
                cells.number_format = "#,##0"
    sheet.range((header_row, 1), (header_row + max(n_rows, 1), last_col)).columns.autofit()


def _save(book: xw.Book, path: Path) -> None:
    book.api.RemovePersonalInformation = True
    book.save(str(path))
    book.close()
    scrub_personal_info(path)
    verify_no_personal_info(path)


def write_report(app: xw.App, report: Report, path: Path) -> None:
    book = _new_book(app)
    sheet = book.sheets[0]
    sheet.name = SHEET_NAME
    sheet.range("A1").value = TITLE
    sheet.range("A1").font.bold = True
    sheet.range("B3").number_format = "@"  # 「2026年9月」を Excel が日付に変えないよう文字列にする
    sheet.range("A2").value = [["代理店名", report.agency], ["対象月", f"{YEAR}年{MONTH}月"]]
    if report.kind in NOTES:
        sheet.range("D2").value = NOTES[report.kind]
    if report.note:
        sheet.range("A4").value = report.note
    values = [list(report.columns)]
    values += [[line.by_column()[c] for c in report.columns] for line in report.lines]
    sheet.range((report.header_row, 1)).value = values
    _format_table(sheet, report.columns, report.header_row, len(report.lines))
    _save(book, path)


def write_answer(app: xw.App, kept: list[MergedRow], removed: list[RemovedRow], path: Path) -> None:
    book = _new_book(app)
    sheet = book.sheets[0]
    sheet.name = "統合"
    sheet.range("A1").value = [list(ANSWER_HEADER)] + [list(r) for r in kept]
    _format_table(sheet, ("代理店名", *STANDARD_COLUMNS), 1, len(kept))
    removed_sheet = book.sheets.add("除いた重複行", after=sheet)
    removed_sheet.range("A1").value = [list(REMOVED_HEADER)] + [[r.file_name, r.reason, *r.row] for r in removed]
    _format_table(removed_sheet, REMOVED_HEADER, 1, len(removed))
    sheet.activate()
    _save(book, path)


def print_summary(reports: list[Report], kept: list[MergedRow], removed: list[RemovedRow]) -> None:
    before_rows = sum(len(r.lines) for r in reports)
    before_total = sum(line.amount for r in reports for line in r.lines)
    after_total = sum(r.amount for r in kept)
    removed_total = sum(r.row.amount for r in removed)
    print(f"ファイル数 {len(reports)}（代理店 {len(master.AGENCIES)} 社）")
    print(f"統合前 {before_rows} 行 {before_total:,} 円")
    print(f"統合後 {len(kept)} 行 {after_total:,} 円")
    print(f"除いた行 {len(removed)} 行 {removed_total:,} 円")
    for reason in (REASON_REPLACED, REASON_RESENT, REASON_DUPLICATE, REASON_OTHER_FILE):
        rows = [r for r in removed if r.reason == reason]
        if rows:
            print(f"  {reason}\t{len(rows)} 行\t{sum(r.row.amount for r in rows):,} 円")
    balanced = (len(kept) + len(removed), after_total + removed_total) == (before_rows, before_total)
    print(f"検算：統合後＋除いた行＝統合前 … {'一致' if balanced else '不一致'}"
          f"（{len(kept)}＋{len(removed)}＝{len(kept) + len(removed)} 行、"
          f"{after_total:,}＋{removed_total:,}＝{after_total + removed_total:,} 円）")
    if not balanced:
        raise RuntimeError("統合後＋除いた行が統合前と合わない")
    print("\nファイル\t行数\t金額\t見出し行\t列の順番\t部署名")
    for r in reports:
        order = "様式どおり" if r.columns == STANDARD_COLUMNS else "違う"
        depts = "・".join(sorted({line.department for line in r.lines})) or "-"
        print(f"{r.file_name}\t{len(r.lines)}\t{sum(line.amount for line in r.lines):,}\t"
              f"{r.header_row}\t{order}\t{depts}")


def main() -> int:
    reports = build_reports(make_sales.build_rows())
    kept, removed = merge_reports(reports)
    stale = {p.name for p in OUTPUT_DIR.glob("*.xlsx")} - {r.file_name for r in reports}
    if stale:
        raise RuntimeError(f"{OUTPUT_DIR} に今回作らないファイルがある。消してから実行する: {sorted(stale)}")
    with staging.staging_dir() as stage:
        (stage / "reports").mkdir()
        moves = {stage / "reports" / r.file_name: OUTPUT_DIR / r.file_name for r in reports}
        staged_answer = stage / ANSWER.name
        with xw.App(visible=False, add_book=False) as app:
            for report, path in zip(reports, moves, strict=True):
                write_report(app, report, path)
            write_answer(app, kept, removed, staged_answer)
        moves[staged_answer] = ANSWER
        staging.publish_all(moves)  # 全ファイルの個人情報の検査（_save）が通ってから移す
    print(f"saved {len(reports)} files to {OUTPUT_DIR} and {ANSWER} (personal info removed)\n")
    print_summary(reports, kept, removed)
    return 0


if __name__ == "__main__":
    sys.exit(main())
