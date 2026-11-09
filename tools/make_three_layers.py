"""「見えている値・数式・書式の三層」を示すブック（Issue #18）を作る。仕様は docs/sample-data.md §7。

5.2・5.3 と第5章のコラム（Range.Value / Formula / NumberFormat / Text）の題材。
売上明細（make_sales.build_rows）から 2026年9月の部署別売上を集計し、小さな集計表にする。

- B1：「2026年9月」と見えるが、中身は日付のシリアル値（表示形式 yyyy"年"m"月"）
- B列・D列：「12,764,608円」と見えるが、中身は数値 12764608（「円」とカンマは表示形式）
- C列：税込額の数式 =B4*1.1
- E列：前年同月比の数式 =B4/D4-1（表示形式 0.0%）。前年同月比が −5% を下回るセルに条件付き書式で色が付く（広島営業所だけ）
- 合計行：=SUM(B4:B11) など

ブックを一時フォルダに保存して個人情報を取り除いたあと、もう一度 Excel で開いて各セルの Text・Value2・Formula・
NumberFormat・条件付き書式を反映した色を読み出し、原稿用の正解の一覧（Markdown）を書き出す。両方そろってから
samples/ に移す（tools/staging.py）。

    uv run python tools/make_three_layers.py
"""
import datetime as dt
import sys
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import make_sales
import sample_master as master
import staging
from xlsx_meta import scrub_personal_info, verify_no_personal_info

ROOT = Path(__file__).resolve().parent.parent
OUTPUT = ROOT / "samples" / "部署別売上_2026年9月.xlsx"
ANSWER = ROOT / "samples" / "answers" / "部署別売上_2026年9月_セルの中身.md"
SHEET_NAME = "集計"

YEAR, MONTH = 2026, 9
MONTH_CELL = "B1"
HEADER = ("部署名", "売上（税抜）", "売上（税込）", "前年同月", "前年同月比")
HEADER_ROW = 3
FIRST_ROW = HEADER_ROW + 1

YEN_FORMAT = '#,##0"円"'
PERCENT_FORMAT = "0.0%"
MONTH_FORMAT = 'yyyy"年"m"月"'
# 前年同月比がこれを下回ると色を付ける。0 にすると、2026年9月は広島営業所のほかに
# 本社営業課（−0.1%）・仙台営業所・名古屋支店（ともに −1.5%）にも色が付き、「広島だけが落ちている」が見えなくなる
HIGHLIGHT_BELOW = -0.05
HEADER_COLOR = (221, 235, 247)  # make_sales の見出しと同じ薄い青
HIGHLIGHT_FILL = (255, 199, 206)  # Excel の既定の「明るい赤の背景」
HIGHLIGHT_FONT = (156, 0, 6)  # 同「濃い赤の文字」

# Excel の定数
XL_CELL_VALUE = 1
XL_LESS = 6
XL_COLOR_INDEX_NONE = -4142


class DepartmentSales(NamedTuple):
    department: str
    this_year: int  # 2026年9月の金額合計（税抜）
    last_year: int  # 2025年9月の金額合計（税抜）

    @property
    def year_over_year(self) -> float | None:
        """前年同月比（増減率）。前年同月の売上が0なら None（ブックの数式は #DIV/0! になる）。"""
        return self.this_year / self.last_year - 1 if self.last_year else None


def summary_lines(table: list[DepartmentSales]) -> list[str]:
    """標準出力に出す表（見出しと部署）。タブ区切り。前年同月比が出せない部署は「－」。"""
    lines = ["部署名\t2026年9月\t2025年9月\t前年同月比"]
    for t in table:
        ratio = "－" if t.year_over_year is None else f"{t.year_over_year:+.1%}"
        lines.append(f"{t.department}\t{t.this_year:,}\t{t.last_year:,}\t{ratio}")
    return lines


class CellInfo(NamedTuple):
    """Excel で開いたときのセルの中身。"""
    text: str
    value2: object
    formula: str | None  # 数式がなければ None
    number_format: str
    highlight: tuple[int, int, int] | None  # 条件付き書式などで表示上付いている塗りつぶし色。なければ None


def department_sales(rows: list[make_sales.SalesRow]) -> list[DepartmentSales]:
    """部署ごとの 2026年9月と前年同月の金額合計（返品行を含む）。並びはマスタの売上部署の順。"""
    sums: dict[tuple[str, int], int] = {}
    for r in rows:
        if r.date.month == MONTH and r.date.year in (YEAR, YEAR - 1):
            key = (r.department, r.date.year)
            sums[key] = sums.get(key, 0) + r.amount
    return [DepartmentSales(d, sums.get((d, YEAR), 0), sums.get((d, YEAR - 1), 0))
            for d in master.SALES_DEPARTMENT_NAMES]


def _rgb_to_excel(rgb: tuple[int, int, int]) -> int:
    r, g, b = rgb
    return r + g * 256 + b * 65536


def _excel_to_rgb(color: int) -> tuple[int, int, int]:
    return color % 256, color // 256 % 256, color // 65536


def write_workbook(table: list[DepartmentSales], path: Path) -> None:
    last = FIRST_ROW + len(table) - 1
    total = last + 1
    with xw.App(visible=False, add_book=False) as app:
        book = app.books.add()
        while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
            book.sheets[-1].delete()
        sheet = book.sheets[0]
        sheet.name = SHEET_NAME

        sheet.range("A1").value = "対象月"
        sheet.range(MONTH_CELL).value = dt.datetime(YEAR, MONTH, 1)
        sheet.range(MONTH_CELL).number_format = MONTH_FORMAT

        sheet.range(f"A{HEADER_ROW}").value = list(HEADER)
        header = sheet.range(f"A{HEADER_ROW}:E{HEADER_ROW}")
        header.font.bold = True
        header.color = HEADER_COLOR

        for i, t in enumerate(table):
            r = FIRST_ROW + i
            sheet.range(f"A{r}").value = [t.department, t.this_year, f"=B{r}*1.1", t.last_year, f"=B{r}/D{r}-1"]
        sheet.range(f"A{total}").value = [
            "合計", f"=SUM(B{FIRST_ROW}:B{last})", f"=SUM(C{FIRST_ROW}:C{last})",
            f"=SUM(D{FIRST_ROW}:D{last})", f"=B{total}/D{total}-1",
        ]
        sheet.range(f"A{total}:E{total}").font.bold = True
        sheet.range(f"B{FIRST_ROW}:D{total}").number_format = YEN_FORMAT
        sheet.range(f"E{FIRST_ROW}:E{total}").number_format = PERCENT_FORMAT

        rule = sheet.range(f"E{FIRST_ROW}:E{last}").api.FormatConditions.Add(
            Type=XL_CELL_VALUE, Operator=XL_LESS, Formula1=f"={HIGHLIGHT_BELOW}")
        rule.Interior.Color = _rgb_to_excel(HIGHLIGHT_FILL)
        rule.Font.Color = _rgb_to_excel(HIGHLIGHT_FONT)

        sheet.range(f"A1:E{total}").columns.autofit()
        book.api.RemovePersonalInformation = True
        book.save(str(path))
        book.close()


def read_cells(path: Path) -> dict[str, CellInfo]:
    """ブックを Excel で開き、値の入っているセルの中身を読み出す。キーはセル番地（A1 形式）。"""
    cells = {}
    with xw.App(visible=False, add_book=False) as app:
        book = app.books.open(str(path), read_only=True)
        try:
            sheet = book.sheets[SHEET_NAME]
            for cell in sheet.used_range:
                api = cell.api
                if api.Value2 is None:
                    continue
                interior = api.DisplayFormat.Interior
                highlight = None if interior.ColorIndex == XL_COLOR_INDEX_NONE else _excel_to_rgb(int(interior.Color))
                cells[cell.get_address(False, False)] = CellInfo(
                    text=api.Text,
                    value2=api.Value2,
                    formula=api.Formula if api.HasFormula else None,
                    number_format=api.NumberFormat,
                    highlight=highlight,
                )
        finally:
            book.close()
    return cells


def _format_value(value: object) -> str:
    if isinstance(value, float):
        return f"{value:.15g}"  # Excel が扱う有効桁数（15桁）で書く。1.1 倍の計算誤差の末尾は出さない
    return str(value)


def _md(text: str) -> str:
    return text.replace("|", r"\|")


def _color_name(cell: CellInfo, header: bool) -> str:
    if cell.highlight is None or header:
        return "（なし）"
    if cell.highlight == HIGHLIGHT_FILL:
        return "明るい赤（条件付き書式）"
    raise ValueError(f"想定外の塗りつぶし色: {cell.highlight}")


def write_answer(cells: dict[str, CellInfo], path: Path) -> None:
    hiroshima = FIRST_ROW + master.SALES_DEPARTMENT_NAMES.index("広島営業所")
    lines = [
        "# 部署別売上_2026年9月.xlsx のセルの中身（正解の一覧）",
        "",
        "`tools/make_three_layers.py` が生成する。手で編集しない。",
        "ブックを Excel（Microsoft 365・日本語版）で開いて、各セルの `Range.Text`・`Range.Value2`・`Range.Formula`・"
        "`Range.NumberFormat` と、条件付き書式を反映した塗りつぶし色（`Range.DisplayFormat.Interior`）を読み出したもの。",
        "",
        f"- シート：「{SHEET_NAME}」の1枚。値は売上明細（`samples/売上明細_2024年10月-2026年9月.xlsx`）の"
        "2026年9月と2025年9月の金額を部署別に合計したもの（返品行を含む・税抜）",
        f"- {MONTH_CELL}：「2026年9月」と見えるが、中身は 2026/9/1 のシリアル値。"
        "VBA の `Range.Value` は日付型（2026/09/01）、xlwings の `range.value` は `datetime` で返る",
        f"- B列・D列：「円」とカンマは表示形式が付けている。中身はただの数値",
        f"- C列：税込額の数式（`=B{FIRST_ROW}*1.1` など）。中身の値は計算結果で、小数を含むことがある（表示は四捨五入）",
        f"- E列：前年同月比の数式。前年同月比が {HIGHLIGHT_BELOW:.0%} を下回るセルに条件付き書式で色が付く"
        f"（E{hiroshima}＝広島営業所だけ）。色はセル自体の書式ではなく、ルールが表示のときに付けている",
        "",
        "| セル | 見えている値（Text） | 中身の値（Value2） | 数式（Formula） | 表示形式（NumberFormat） | 条件付き書式の色 |",
        "|---|---|---|---|---|---|",
    ]
    for ref, cell in cells.items():
        header = int(ref[1:]) == HEADER_ROW
        formula = f"`{_md(cell.formula)}`" if cell.formula else "（なし）"
        lines.append(
            f"| {ref} | {_md(cell.text)} | {_md(_format_value(cell.value2))} | {formula} | "
            f"`{_md(cell.number_format)}` | {_color_name(cell, header)} |"
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    table = department_sales(make_sales.build_rows())
    with staging.staging_dir() as stage:
        book, answer = stage / OUTPUT.name, stage / ANSWER.name
        write_workbook(table, book)
        scrub_personal_info(book)
        verify_no_personal_info(book)
        write_answer(read_cells(book), answer)
        staging.publish_all({book: OUTPUT, answer: ANSWER})
    print(f"saved {OUTPUT} (personal info removed)")
    print(f"saved {ANSWER}\n")
    print("\n".join(summary_lines(table)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
