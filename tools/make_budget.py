"""部署別の年度予算表（Issue #14）のサンプルブックと、変換後の正解を作る。仕様は docs/sample-data.md §2.1・§7。

8.2「結合セルや複数行ヘッダーの解消」の題材。
- 見た目の表（samples/）：2025年度の売上予算を、人間が見やすいように作った表。
  2段ヘッダー（上期・下期 × 月）、部（東日本営業部・西日本営業部）ごとの小計行、結合セル、罫線・塗りつぶし・太字。
  小計・合計・上期計・下期計・年間合計は SUM の数式。単位は千円。
- 正解（samples/answers/）：同じ予算を「部署・年月・予算額（千円）」の縦持ちにした、データとしての表。
  部署は §2.1 の「予算表での書き方」（「西日本営業部 大阪支店」型）。結合セルの部と課・拠点をつなげた名前になる。

予算額は売上明細（tools/make_sales.py）の2025年度の実績を部署別・月別に集計し、それに上乗せして10万円単位に丸めて作る。
広島営業所だけは前年並みの見込みで予算を立てた（実績の落ち込みを見込んでいない）ので、予算達成率が9割を切る。
乱数は固定シードなので、何度実行しても同じ内容になる。一時フォルダに保存して個人情報を取り除き、残っていないことを
確かめてから samples/ に移す（tools/staging.py）。

    uv run python tools/make_budget.py

生成ルールやシードを変えたら、ブックを作り直し、tests/test_make_budget.py のゴールデン値を出力に合わせて更新する。
"""
import datetime as dt
import random
import sys
from collections import defaultdict
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import make_sales
import sample_master as master
import staging
from xlsx_meta import scrub_personal_info, verify_no_personal_info

SAMPLES = Path(__file__).resolve().parent.parent / "samples"
OUTPUT = SAMPLES / "部署別売上予算_2025年度.xlsx"
ANSWER_OUTPUT = SAMPLES / "answers" / "部署別売上予算_2025年度_縦持ち.xlsx"
SHEET_NAME = "2025年度予算"
ANSWER_SHEET_NAME = "予算"
ANSWER_HEADER = ["部署", "年月", "予算額（千円）"]
SEED = 20250401

FISCAL_YEAR = 2025
FY_MONTHS = [dt.date(FISCAL_YEAR, m, 1) for m in range(4, 13)] + [dt.date(FISCAL_YEAR + 1, m, 1) for m in range(1, 4)]
MONTH_LABELS = [f"{m.month}月" for m in FY_MONTHS]

# 実績への上乗せ（平均と標準偏差）。予算は実績よりやや高めに立てる
BUDGET_UPLIFT = 1.04
BUDGET_NOISE = 0.04
# 広島営業所は落ち込みを見込まずに予算を立てた（make_sales.HIROSHIMA_FACTOR[2025] の逆数ぶん高い）
HIROSHIMA_PLAN = 1 / make_sales.HIROSHIMA_FACTOR[FISCAL_YEAR]
HIROSHIMA_NAME = "広島営業所"
# 予算は10万円単位（千円単位で100の倍数）
ROUND_YEN = 100_000

TITLE = f"{FISCAL_YEAR}年度 部署別売上予算"
SUBTITLE = f"{master.COMPANY_NAME} 営業本部"
UNIT_NOTE = "（単位：千円）"
FOOTNOTE = "※ 2025年3月の経営会議で承認。10月の単価改定（ポンプ・バルブ3%値上げ）を織り込み済み。"

# 列：A＝部、B＝課・拠点、C〜H＝4〜9月、I＝上期計、J〜O＝10〜3月、P＝下期計、Q＝年間合計
H1_MONTH_COLS = "CDEFGH"
H2_MONTH_COLS = "JKLMNO"
H1_TOTAL_COL, H2_TOTAL_COL, YEAR_TOTAL_COL = "I", "P", "Q"
LAST_COL = YEAR_TOTAL_COL
VALUE_COLS = "CDEFGHIJKLMNOPQ"
HEADER_ROWS = (4, 5)
FIRST_DATA_ROW = 6

# 色（人間向けの見た目）
HEADER_FILL = (47, 84, 150)
HEADER_FONT = (255, 255, 255)
SUBTOTAL_FILL = (255, 242, 204)
TOTAL_FILL = (217, 225, 242)
HALF_TOTAL_FILL = (242, 242, 242)

# Excel の定数（xlwings の api で使う）
XL_CENTER = -4108
XL_RIGHT = -4152
XL_CONTINUOUS = 1
XL_DOUBLE = -4119
XL_THIN = 2
XL_EDGE_TOP = 8


class Budget(NamedTuple):
    department: str  # 予算表での書き方（§2.1）
    month: dt.date   # 月の1日
    amount: int      # 予算額（千円）


class LayoutRow(NamedTuple):
    """見た目の表の1行。kind は unit（部署）・subtotal（部の小計）・total（営業本部の合計）。"""
    row: int
    kind: str
    group: str               # A列（部。合計行は「営業本部 合計」）
    label: str | None        # B列（課・拠点、または「小計」。部の下に課・拠点がない行は None）
    department: str | None   # unit の行だけ。予算表での書き方


def month_column(month: dt.date) -> str:
    i = FY_MONTHS.index(month)
    return H1_MONTH_COLS[i] if i < 6 else H2_MONTH_COLS[i - 6]


def layout() -> list[LayoutRow]:
    """見た目の表の行構成。部の下に課・拠点が2つ以上ある部だけ小計行を置き、最後に営業本部の合計行を置く。"""
    units_by_dept: dict[str, list[master.OrgUnit]] = defaultdict(list)
    for d in master.SALES_DEPARTMENTS:
        unit = next(u for u in master.ORG_UNITS if u.code == d.answer_code)
        units_by_dept[unit.department].append(unit)
    rows: list[LayoutRow] = []
    r = FIRST_DATA_ROW
    for dept, units in units_by_dept.items():
        for u in units:
            rows.append(LayoutRow(r, "unit", dept, u.section, u.name))
            r += 1
        if len(units) > 1:
            rows.append(LayoutRow(r, "subtotal", dept, "小計", None))
            r += 1
    rows.append(LayoutRow(r, "total", "営業本部 合計", None, None))
    return rows


def merged_ranges() -> list[str]:
    """結合セルの範囲。表題・2段ヘッダー・部（縦に結合）・課のない行と合計行（横に結合）。"""
    ranges = [f"A1:{LAST_COL}1", "A4:B5", f"C4:{H1_TOTAL_COL}4", f"J4:{H2_TOTAL_COL}4", f"{YEAR_TOTAL_COL}4:{YEAR_TOTAL_COL}5"]
    rows = layout()
    groups: dict[str, list[LayoutRow]] = defaultdict(list)
    for r in rows:
        groups[r.group].append(r)
    for group_rows in groups.values():
        first, last = group_rows[0].row, group_rows[-1].row
        if len(group_rows) > 1:
            ranges.append(f"A{first}:A{last}")
        elif group_rows[0].label is None:
            ranges.append(f"A{first}:B{first}")
    return ranges


def _actual_by_department_month(rows: list[make_sales.SalesRow]) -> dict[tuple[str, dt.date], int]:
    sums: dict[tuple[str, dt.date], int] = defaultdict(int)
    first, last = FY_MONTHS[0], FY_MONTHS[-1]
    for r in rows:
        month = r.date.replace(day=1)
        if first <= month <= last:
            sums[(r.department, month)] += r.amount
    return sums


def build_budget(seed: int = SEED) -> list[Budget]:
    """2025年度の予算（部署順・月順）。同じシードなら同じ内容になる。"""
    rng = random.Random(seed)
    actual = _actual_by_department_month(make_sales.build_rows())
    budget = []
    for d in master.SALES_DEPARTMENTS:
        plan = HIROSHIMA_PLAN if d.name == HIROSHIMA_NAME else 1.0
        for month in FY_MONTHS:
            yen = actual[(d.name, month)] * plan * rng.gauss(BUDGET_UPLIFT, BUDGET_NOISE)
            amount = round(yen / ROUND_YEN) * (ROUND_YEN // 1000)
            if amount <= 0:
                raise ValueError(f"{d.budget_name} {month:%Y/%m}: 予算額が0以下になった（実績 {actual[(d.name, month)]:,}円）")
            budget.append(Budget(d.budget_name, month, amount))
    return budget


def _row_formulas(row: int) -> dict[str, str]:
    return {
        H1_TOTAL_COL: f"=SUM({H1_MONTH_COLS[0]}{row}:{H1_MONTH_COLS[-1]}{row})",
        H2_TOTAL_COL: f"=SUM({H2_MONTH_COLS[0]}{row}:{H2_MONTH_COLS[-1]}{row})",
        YEAR_TOTAL_COL: f"={H1_TOTAL_COL}{row}+{H2_TOTAL_COL}{row}",
    }


def _sum_formula(col: str, rows: list[int]) -> str:
    if rows == list(range(rows[0], rows[-1] + 1)):
        return f"=SUM({col}{rows[0]}:{col}{rows[-1]})"
    return "=" + "+".join(f"{col}{r}" for r in rows)


def _new_book(app: xw.App, name: str) -> tuple[xw.Book, xw.Sheet]:
    book = app.books.add()
    while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
        book.sheets[-1].delete()
    sheet = book.sheets[0]
    sheet.name = name
    return book, sheet


def _save(book: xw.Book, path: Path) -> None:
    book.api.RemovePersonalInformation = True
    book.save(str(path))
    book.close()


def _write_table(sheet: xw.Sheet, budget: list[Budget]) -> None:
    rows = layout()
    amounts = {(b.department, b.month): b.amount for b in budget}
    last_row = rows[-1].row

    sheet.range("A1").value = TITLE
    sheet.range("A2").value = SUBTITLE
    sheet.range(f"{YEAR_TOTAL_COL}2").value = UNIT_NOTE
    sheet.range("A4").value = "部署"
    sheet.range("C4").value = "上期"
    sheet.range("J4").value = "下期"
    sheet.range(f"{YEAR_TOTAL_COL}4").value = "年間合計"
    sheet.range("C5").value = [MONTH_LABELS[:6] + ["上期計"] + MONTH_LABELS[6:] + ["下期計"]]

    previous_group = None
    for r in rows:
        if r.group != previous_group:  # 結合する範囲の左上のセルにだけ書く
            sheet.range(f"A{r.row}").value = r.group
            previous_group = r.group
        if r.label is not None:
            sheet.range(f"B{r.row}").value = r.label
        if r.kind == "unit":
            sheet.range(f"C{r.row}").value = [[amounts[(r.department, m)] for m in FY_MONTHS[:6]]]
            sheet.range(f"J{r.row}").value = [[amounts[(r.department, m)] for m in FY_MONTHS[6:]]]
        else:
            if r.kind == "subtotal":
                targets = [u.row for u in rows if u.kind == "unit" and u.group == r.group]
            else:  # 合計：小計行と、小計のない部の行
                grouped = {u.group for u in rows if u.kind == "subtotal"}
                targets = [u.row for u in rows
                           if u.kind == "subtotal" or (u.kind == "unit" and u.group not in grouped)]
            for col in H1_MONTH_COLS + H2_MONTH_COLS:
                sheet.range(f"{col}{r.row}").formula = _sum_formula(col, targets)
        for col, formula in _row_formulas(r.row).items():
            sheet.range(f"{col}{r.row}").formula = formula

    sheet.range(f"A{last_row + 2}").value = FOOTNOTE

    # --- 見た目 ---
    for ref in merged_ranges():
        sheet.range(ref).merge()
    title = sheet.range(f"A1:{LAST_COL}1")
    title.font.bold = True
    title.font.size = 14
    title.api.HorizontalAlignment = XL_CENTER
    sheet.range(f"{YEAR_TOTAL_COL}2").api.HorizontalAlignment = XL_RIGHT

    header = sheet.range(f"A4:{LAST_COL}5")
    header.color = HEADER_FILL
    header.font.color = HEADER_FONT
    header.font.bold = True
    header.api.HorizontalAlignment = XL_CENTER
    header.api.VerticalAlignment = XL_CENTER

    body = sheet.range(f"A{FIRST_DATA_ROW}:{LAST_COL}{last_row}")
    body.number_format = "#,##0"
    sheet.range(f"A{FIRST_DATA_ROW}:A{last_row}").api.VerticalAlignment = XL_CENTER
    for col in (H1_TOTAL_COL, H2_TOTAL_COL):
        sheet.range(f"{col}{FIRST_DATA_ROW}:{col}{last_row}").color = HALF_TOTAL_FILL
    for r in rows:
        line = sheet.range(f"A{r.row}:{LAST_COL}{r.row}")
        if r.kind == "subtotal":
            sheet.range(f"B{r.row}:{LAST_COL}{r.row}").color = SUBTOTAL_FILL
            line.font.bold = True
            sheet.range(f"B{r.row}").api.HorizontalAlignment = XL_CENTER
        elif r.kind == "total":
            line.color = TOTAL_FILL
            line.font.bold = True
            sheet.range(f"A{r.row}").api.HorizontalAlignment = XL_CENTER
    sheet.range(f"{YEAR_TOTAL_COL}{FIRST_DATA_ROW}:{YEAR_TOTAL_COL}{last_row}").font.bold = True

    borders = sheet.range(f"A4:{LAST_COL}{last_row}").api.Borders
    borders.LineStyle = XL_CONTINUOUS
    borders.Weight = XL_THIN
    total_top = sheet.range(f"A{last_row}:{LAST_COL}{last_row}").api.Borders(XL_EDGE_TOP)
    total_top.LineStyle = XL_DOUBLE

    sheet.range("A:B").column_width = 14
    sheet.range(f"C:{LAST_COL}").column_width = 9
    sheet.range(f"{YEAR_TOTAL_COL}:{YEAR_TOTAL_COL}").column_width = 11
    sheet.book.api.Windows(1).DisplayGridlines = False


def _write_answer(sheet: xw.Sheet, budget: list[Budget]) -> None:
    sheet.range("A1").value = [ANSWER_HEADER] + [[b.department, b.month, b.amount] for b in budget]
    last = len(budget) + 1
    sheet.range("A1:C1").font.bold = True
    sheet.range(f"B2:B{last}").number_format = "yyyy/mm"
    sheet.range(f"C2:C{last}").number_format = "#,##0"
    sheet.range(f"A1:C{last}").columns.autofit()


def write_workbooks(budget: list[Budget], table_path: Path, answer_path: Path) -> None:
    with xw.App(visible=False, add_book=False) as app:
        book, sheet = _new_book(app, SHEET_NAME)
        _write_table(sheet, budget)
        app.calculate()
        _save(book, table_path)
        book, sheet = _new_book(app, ANSWER_SHEET_NAME)
        _write_answer(sheet, budget)
        _save(book, answer_path)


def print_summary(budget: list[Budget]) -> None:
    total = sum(b.amount for b in budget)
    print(f"予算 {len(budget)} 行　年間合計 {total:,} 千円")
    for name in dict.fromkeys(b.department for b in budget):
        h1 = sum(b.amount for b in budget if b.department == name and b.month in FY_MONTHS[:6])
        year = sum(b.amount for b in budget if b.department == name)
        print(f"{name}\t上期 {h1:,}\t下期 {year - h1:,}\t年間 {year:,}")


def main() -> int:
    budget = build_budget()
    with staging.staging_dir() as stage:
        moves = {stage / OUTPUT.name: OUTPUT, stage / ANSWER_OUTPUT.name: ANSWER_OUTPUT}
        write_workbooks(budget, *moves)
        for path in moves:
            scrub_personal_info(path)
            verify_no_personal_info(path)
        staging.publish_all(moves)
    for path in moves.values():
        print(f"saved {path} (personal info removed)")
    print_summary(budget)
    return 0


if __name__ == "__main__":
    sys.exit(main())
