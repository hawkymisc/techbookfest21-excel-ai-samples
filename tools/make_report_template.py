"""会社指定のレポート様式（Issue #16）のサンプルを作る。仕様は docs/sample-data.md §7、使うのは 9.3。

株式会社ミナセ流体工業の「月次営業報告書」。次の3つを作る。

- samples/月次営業報告書_テンプレート.xlsx：入力欄が空の様式
- samples/answers/月次営業報告書_2026年9月.xlsx：2026年9月分を流し込んだ正解
- samples/answers/月次営業報告書_2026年9月.pdf：正解を Excel の ExportAsFixedFormat で PDF にしたもの

様式は「黄色の入力欄」と「それ以外（ロゴ・見出し・部署名・注記・計算式）」に分かれている。
9.3 では、AI に様式を触らせず入力欄だけを埋めさせ、PDF にする。正解の値は売上明細（make_sales.build_rows()）から集計する。
正解はテンプレートのファイルを開いて入力欄だけを埋めて作るので、様式がテンプレートと同じになる。

    uv run python tools/make_report_template.py

PDF の作成日時は実行のたびに変わる（中身は同じ）。3つとも一時フォルダに作り、xlsx の個人情報の除去と検査、PDF の
検査を通ってから samples/ に移す（tools/staging.py）。PDF は作成者・保存先パスのほか、実行ユーザー名（ホーム
ディレクトリ名・環境変数 USERNAME・Office のユーザー名）が入っていないことを確かめる。

Excel は COM で印刷設定（PageSetup）を1つでも触ると、既定のプリンターの設定（DEVMODE。プリンター名を含む）を
xl/printerSettings/printerSettings1.bin としてブックに書き込む。PrintCommunication を切っても出る。
プリンター名は作業者の環境の情報なので、保存後にこのパートと参照だけを取り除く（用紙・向き・1ページに収める設定は
シートの XML の pageSetup 要素にあり、取り除いても残る）。ワークシートとグラフシートの両方を対象にし、取り除いたあとに
シートの r:id がすべて rels にあること（壊れたブックにしていないこと）を確かめる。
"""
import datetime as dt
import os
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import zlib
from collections import defaultdict
from collections.abc import Iterable
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import make_sales
import sample_master as master
import staging
from xlsx_meta import PATH_PATTERN, scrub_personal_info, verify_no_personal_info

try:
    import winreg  # Office のユーザー名をレジストリから読む（Windows だけ）
except ImportError:
    winreg = None  # Windows 以外。Office のユーザー名は Excel の Application.UserName だけで確かめる

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "samples" / "月次営業報告書_テンプレート.xlsx"
ANSWER = ROOT / "samples" / "answers" / "月次営業報告書_2026年9月.xlsx"
ANSWER_PDF = ANSWER.with_suffix(".pdf")

SHEET_NAME = "月次営業報告書"
REPORT_MONTH = dt.date(2026, 9, 1)
# 正解の作成日。本文の「今日」は 2026年10月（docs/sample-data.md §1）なので、10月の最初の月曜にした
REPORT_DATE = dt.date(2026, 10, 5)

# --- 様式のレイアウト ---
DEPT_FIRST_ROW = 10
DEPT_ROWS = range(DEPT_FIRST_ROW, DEPT_FIRST_ROW + len(master.SALES_DEPARTMENT_NAMES))
TOTAL_ROW = DEPT_ROWS.stop
PRINT_AREA = "$A$1:$F$28"
# 入力欄（名前を付けておく。AI や読者がセル番地でなく名前で探せる）
INPUT_NAMES = {
    "報告年月": "$B$6",
    "作成日": "$E$6",
    "当月売上": f"$B${DEPT_ROWS.start}:$B${DEPT_ROWS.stop - 1}",
    "前月売上": f"$C${DEPT_ROWS.start}:$C${DEPT_ROWS.stop - 1}",
    "前年同月売上": f"$E${DEPT_ROWS.start}:$E${DEPT_ROWS.stop - 1}",
    "特記事項": "$A$21:$F$24",
}
COLUMN_WIDTHS = {"A": 16, "B": 15, "C": 15, "D": 10, "E": 15, "F": 10}
INPUT_FILL = (255, 242, 204)        # 入力欄：薄い黄色
INPUT_FILL_ARGB = "FFFFF2CC"        # 上の色を styles.xml で表したもの（テストで使う）
FORMULA_FILL = (242, 242, 242)      # 計算式の欄：薄い灰色
HEADER_FILL = (217, 225, 242)       # 表の見出し：薄い青
LOGO_NAVY = (31, 56, 100)
LOGO_LIGHT = (155, 194, 230)
NOTES = (
    "※ 金額は販売管理システムの売上明細の税抜金額の合計。返品（数量がマイナスの行）を差し引いた額。",
    "※ 前月比＝当月売上÷前月売上−1、前年同月比＝当月売上÷前年同月売上−1（いずれも増減率）。全社合計と増減率は計算式で自動計算する。",
    "※ 黄色の欄だけを入力する。ロゴ・見出し・部署名・注記・計算式は変更しない。",
)

# Excel の定数
XL_PAPER_A4 = 9
XL_PORTRAIT = 1
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_CENTER = -4108
XL_LEFT = -4131
XL_RIGHT = -4152
XL_TOP = -4160
XL_CONTINUOUS = 1
XL_DOUBLE = -4119
XL_THIN = 2
XL_EDGE_TOP = 8
MSO_SHAPE_OVAL = 9
MSO_SHAPE_WAVE = 103
MSO_TEXT_ORIENTATION_HORIZONTAL = 1


class DepartmentSales(NamedTuple):
    """報告書の1行。金額は税抜・返品差し引き後の合計。"""
    department: str
    current: int
    previous_month: int
    previous_year: int

    @property
    def month_over_month(self) -> float | None:
        return self.current / self.previous_month - 1 if self.previous_month else None

    @property
    def year_over_year(self) -> float | None:
        return self.current / self.previous_year - 1 if self.previous_year else None


def _month_before(month: dt.date) -> dt.date:
    return dt.date(month.year - 1, 12, 1) if month.month == 1 else dt.date(month.year, month.month - 1, 1)


def summarize(rows: Iterable[make_sales.SalesRow], year: int, month: int) -> list[DepartmentSales]:
    """売上明細の行から、部署別の当月・前月・前年同月の売上を集計する。部署はマスタの順で全部出す。"""
    current = dt.date(year, month, 1)
    previous_month = _month_before(current)
    previous_year = dt.date(year - 1, month, 1)
    sums: dict[tuple[str, dt.date], int] = defaultdict(int)
    for r in rows:
        sums[(r.department, dt.date(r.date.year, r.date.month, 1))] += r.amount
    return [
        DepartmentSales(d, sums[(d, current)], sums[(d, previous_month)], sums[(d, previous_year)])
        for d in master.SALES_DEPARTMENT_NAMES
    ]


def company_total(summary: Iterable[DepartmentSales]) -> DepartmentSales:
    summary = list(summary)
    return DepartmentSales(
        "全社",
        sum(s.current for s in summary),
        sum(s.previous_month for s in summary),
        sum(s.previous_year for s in summary),
    )


def format_ratio(ratio: float | None) -> str:
    """増減率の表示。比較月の売上が0で増減率が出せないときは「－」（様式の計算式が空欄を出すのに合わせる）。"""
    return "－" if ratio is None else f"{ratio:.1%}"


def summary_lines(summary: Iterable[DepartmentSales]) -> list[str]:
    """標準出力に出す表（見出し・部署・全社合計）。タブ区切り。"""
    summary = list(summary)
    lines = ["部署\t当月売上\t前月売上\t前月比\t前年同月売上\t前年同月比"]
    for s in [*summary, company_total(summary)]:
        lines.append(f"{s.department}\t{s.current:,}\t{s.previous_month:,}\t{format_ratio(s.month_over_month)}"
                     f"\t{s.previous_year:,}\t{format_ratio(s.year_over_year)}")
    return lines


def remarks(summary: Iterable[DepartmentSales]) -> str:
    """特記事項の欄に書く文。前年同月を下回った部署と、全社の前年同月比。"""
    summary = list(summary)
    below = [s for s in summary if s.year_over_year is not None and s.year_over_year < 0]
    total = company_total(summary)
    if below:
        first = "前年同月を下回った部署：" + "、".join(f"{s.department}（{s.year_over_year:.1%}）" for s in below) + "。"
    else:
        first = "前年同月を下回った部署はない。"
    if total.year_over_year is None:
        return first
    return first + f"全社の前年同月比は {total.year_over_year:.1%}。"


# --- 様式を作る ---

def _draw_logo(sheet: xw.Sheet) -> None:
    """架空の社名ロゴ。丸の中に波を置いたマークと、社名の文字を図形でまとめる（画像は使わない）。"""
    shapes = sheet.api.Shapes
    mark = shapes.AddShape(MSO_SHAPE_OVAL, 6, 6, 40, 40)
    mark.Name = "ロゴ_マーク"
    mark.Fill.ForeColor.RGB = xw.utils.rgb_to_int(LOGO_NAVY)
    mark.Line.Visible = False
    wave = shapes.AddShape(MSO_SHAPE_WAVE, 12, 24, 28, 10)
    wave.Name = "ロゴ_波"
    wave.Fill.ForeColor.RGB = xw.utils.rgb_to_int(LOGO_LIGHT)
    wave.Line.Visible = False
    text = shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, 50, 10, 150, 32)
    text.Name = "ロゴ_社名"
    text.Fill.Visible = False
    text.Line.Visible = False
    frame = text.TextFrame2
    frame.MarginLeft = 0
    frame.WordWrap = False
    frame.TextRange.Text = "ミナセ流体工業"
    frame.TextRange.Font.Size = 16
    frame.TextRange.Font.Bold = True
    frame.TextRange.Font.Fill.ForeColor.RGB = xw.utils.rgb_to_int(LOGO_NAVY)
    group = shapes.Range(["ロゴ_マーク", "ロゴ_波", "ロゴ_社名"]).Group()
    group.Name = "ロゴ"


def _box(rng: xw.Range) -> None:
    for edge in (7, 8, 9, 10, 11, 12):  # 外枠4辺と内側の縦横
        border = rng.api.Borders(edge)
        border.LineStyle = XL_CONTINUOUS
        border.Weight = XL_THIN


def _build_form(book: xw.Book) -> None:
    while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
        book.sheets[-1].delete()
    sheet = book.sheets[0]
    sheet.name = SHEET_NAME
    for col, width in COLUMN_WIDTHS.items():
        sheet.range(f"{col}1").column_width = width
    sheet.api.Cells.Font.Size = 10

    # ヘッダー：ロゴ・発行元・表題
    for r in (1, 2, 3):
        sheet.range(f"A{r}").row_height = 18
    _draw_logo(sheet)
    sheet.range("F2").value = master.COMPANY_NAME
    sheet.range("F3").value = "営業本部"
    sheet.range("F2:F3").api.HorizontalAlignment = XL_RIGHT
    title = sheet.range("A4:F4")
    title.merge()
    title.value = "月次営業報告書"
    title.api.HorizontalAlignment = XL_CENTER
    title.font.size = 18
    title.font.bold = True
    sheet.range("A4").row_height = 32

    # 報告年月・作成日
    sheet.range("A6").value = "報告年月"
    sheet.range("D6").value = "作成日"
    sheet.range("B6").number_format = 'yyyy"年"m"月"'
    sheet.range("E6").number_format = "yyyy/mm/dd"
    sheet.range("A6").font.bold = True
    sheet.range("D6").font.bold = True
    sheet.range("D6").api.HorizontalAlignment = XL_RIGHT
    for ref in ("B6", "E6"):
        _box(sheet.range(ref))
        sheet.range(ref).api.HorizontalAlignment = XL_CENTER

    # 1. 部署別売上
    sheet.range("A8").value = "1. 部署別売上（税抜・円）"
    sheet.range("A8").font.bold = True
    header = sheet.range("A9:F9")
    header.value = ["部署", "当月売上", "前月売上", "前月比", "前年同月売上", "前年同月比"]
    header.color = HEADER_FILL
    header.font.bold = True
    header.api.HorizontalAlignment = XL_CENTER
    first, last = DEPT_ROWS.start, DEPT_ROWS.stop - 1
    sheet.range(f"A{first}").options(transpose=True).value = list(master.SALES_DEPARTMENT_NAMES)
    sheet.range(f"A{TOTAL_ROW}").value = "全社合計"
    for col in "BCE":
        sheet.range(f"{col}{TOTAL_ROW}").formula = (
            f'=IF(COUNT({col}{first}:{col}{last})=0,"",SUM({col}{first}:{col}{last}))'
        )
    for r in [*DEPT_ROWS, TOTAL_ROW]:
        sheet.range(f"D{r}").formula = f'=IF(OR(B{r}="",N(C{r})=0),"",B{r}/C{r}-1)'
        sheet.range(f"F{r}").formula = f'=IF(OR(B{r}="",N(E{r})=0),"",B{r}/E{r}-1)'
    sheet.range(f"B{first}:C{TOTAL_ROW}").number_format = "#,##0"
    sheet.range(f"E{first}:E{TOTAL_ROW}").number_format = "#,##0"
    sheet.range(f"D{first}:D{TOTAL_ROW}").number_format = "0.0%"
    sheet.range(f"F{first}:F{TOTAL_ROW}").number_format = "0.0%"
    sheet.range(f"D{first}:D{TOTAL_ROW}").color = FORMULA_FILL
    sheet.range(f"F{first}:F{TOTAL_ROW}").color = FORMULA_FILL
    sheet.range(f"B{TOTAL_ROW}:C{TOTAL_ROW}").color = FORMULA_FILL
    sheet.range(f"E{TOTAL_ROW}").color = FORMULA_FILL
    sheet.range(f"A{TOTAL_ROW}:F{TOTAL_ROW}").font.bold = True
    _box(sheet.range(f"A9:F{TOTAL_ROW}"))
    total_top = sheet.range(f"A{TOTAL_ROW}:F{TOTAL_ROW}").api.Borders(XL_EDGE_TOP)
    total_top.LineStyle = XL_DOUBLE

    # 2. 特記事項
    sheet.range("A20").value = "2. 特記事項"
    sheet.range("A20").font.bold = True
    note_box = sheet.range("A21:F24")
    note_box.merge()
    note_box.api.WrapText = True
    note_box.api.VerticalAlignment = XL_TOP
    note_box.api.HorizontalAlignment = XL_LEFT
    _box(note_box)

    # 注記
    sheet.range("A26").options(transpose=True).value = list(NOTES)
    sheet.range("A26:A28").font.size = 8

    # 入力欄（黄色）と名前
    for name, ref in INPUT_NAMES.items():
        sheet.range(ref).color = INPUT_FILL
        book.names.add(name, f"='{SHEET_NAME}'!{ref}")

    # 印刷設定：A4 縦、1ページに収める
    setup = sheet.api.PageSetup
    setup.PrintArea = PRINT_AREA
    setup.PaperSize = XL_PAPER_A4
    setup.Orientation = XL_PORTRAIT
    setup.Zoom = False
    setup.FitToPagesWide = 1
    setup.FitToPagesTall = 1
    setup.CenterHorizontally = True
    setup.LeftMargin = book.app.api.CentimetersToPoints(1.8)
    setup.RightMargin = book.app.api.CentimetersToPoints(1.8)
    setup.TopMargin = book.app.api.CentimetersToPoints(2.0)
    setup.BottomMargin = book.app.api.CentimetersToPoints(2.0)
    sheet.range("A1").select()


def _fill_answer(book: xw.Book, summary: list[DepartmentSales]) -> None:
    sheet = book.sheets[SHEET_NAME]
    sheet.range("報告年月").value = REPORT_MONTH
    sheet.range("作成日").value = REPORT_DATE
    sheet.range("当月売上").options(transpose=True).value = [s.current for s in summary]
    sheet.range("前月売上").options(transpose=True).value = [s.previous_month for s in summary]
    sheet.range("前年同月売上").options(transpose=True).value = [s.previous_year for s in summary]
    sheet.range("特記事項")[0].value = remarks(summary)


def _finish(path: Path) -> None:
    """保存した xlsx から、プリンター設定と個人情報を取り除いて確かめる。"""
    drop_printer_settings(path)
    scrub_personal_info(path)
    verify_no_personal_info(path)


def write_template(path: Path) -> None:
    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        book = app.books.add()
        _build_form(book)
        book.api.RemovePersonalInformation = True
        book.save(str(path))
        book.close()
    _finish(path)


def write_answer(summary: list[DepartmentSales], template: Path, answer: Path, pdf: Path) -> None:
    """配るのと同じテンプレートのファイルを開いて入力欄だけを埋め、別名で保存して PDF にする。"""
    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        office_user = app.api.UserName  # Office のユーザー名（PDF に入っていないことを確かめる）
        book = app.books.open(str(template))
        _fill_answer(book, summary)
        book.api.RemovePersonalInformation = True
        book.save(str(answer))
        book.api.ExportAsFixedFormat(
            Type=XL_TYPE_PDF, Filename=str(pdf), Quality=XL_QUALITY_STANDARD,
            IncludeDocProperties=False, IgnorePrintAreas=False, OpenAfterPublish=False,
        )
        book.close()
    _finish(answer)
    verify_pdf_no_personal_info(pdf, default_user_names(extra=[office_user]))


# --- プリンター設定のパートを取り除く ---

PRINTER_PART = re.compile(r"^xl/printerSettings/[^/]+$")
PRINTER_CONTENT_TYPE = re.compile(
    r'<(?:Default|Override)\b[^>]*ContentType="application/vnd\.openxmlformats-officedocument\.spreadsheetml\.printerSettings"[^>]*/>'
)
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# Transitional（.../officeDocument/2006/relationships/printerSettings）と Strict（purl.oclc.org の URL）の両方に合う
PRINTER_REL_TYPE_SUFFIX = "/relationships/printerSettings"
# ワークシートとグラフシートの本体と、その rels
SHEET_PART = re.compile(r"^xl/(?:worksheets|chartsheets)/[^/]+\.xml$")
RELATIONSHIP_ELEMENT = re.compile(r"<Relationship\b[^>]*?(?:/>|>\s*</Relationship\s*>)")
PAGE_SETUP_TAG = re.compile(r"<pageSetup\b[^>]*>")


def _rels_name(sheet_part: str) -> str:
    folder, name = sheet_part.rsplit("/", 1)
    return f"{folder}/_rels/{name}.rels"


def _printer_rel_ids(rels_xml: bytes) -> set[str]:
    """rels の中で、Type がプリンター設定の Relationship の Id。属性の順には依らない（XML として読む）。"""
    root = ET.fromstring(rels_xml)
    return {r.get("Id") for r in root.iter(f"{{{PACKAGE_REL_NS}}}Relationship")
            if (r.get("Type") or "").endswith(PRINTER_REL_TYPE_SUFFIX)}


def _rel_ids(rels_xml: bytes) -> set[str]:
    return {r.get("Id") for r in ET.fromstring(rels_xml).iter(f"{{{PACKAGE_REL_NS}}}Relationship")}


def _drop_relationships(text: str, ids: set[str]) -> str:
    """rels の文字列から、Id が ids の Relationship 要素だけを取り除く（ほかの部分は書き換えない）。"""
    def repl(m: re.Match) -> str:
        return "" if ET.fromstring(m.group(0)).get("Id") in ids else m.group(0)
    return RELATIONSHIP_ELEMENT.sub(repl, text)


def _drop_page_setup_rel_ids(text: str, ids: set[str]) -> str:
    """シートの pageSetup 要素から、ids を指す r:id 属性を取り除く。接頭辞は r: とは限らないので宣言から探す。"""
    prefixes = re.findall(r'xmlns:([A-Za-z_][\w.-]*)\s*=\s*"' + re.escape(OFFICE_REL_NS) + '"', text)

    def repl(m: re.Match) -> str:
        tag = m.group(0)
        for prefix in prefixes:
            tag = re.sub(rf'\s+{re.escape(prefix)}:id\s*=\s*"([^"]*)"',
                         lambda a: "" if a.group(1) in ids else a.group(0), tag)
        return tag

    return PAGE_SETUP_TAG.sub(repl, text)


def check_sheet_relationships(parts: dict[str, bytes], label: str) -> None:
    """全シート（ワークシート・グラフシート）の r:id が、そのシートの rels にあることを確かめる。なければ RuntimeError。"""
    for name, data in parts.items():
        if not SHEET_PART.match(name):
            continue
        used = {value for element in ET.fromstring(data).iter()
                for key, value in element.attrib.items() if key.startswith(f"{{{OFFICE_REL_NS}}}")}
        if not used:
            continue
        rels = parts.get(_rels_name(name))
        missing = used - (_rel_ids(rels) if rels is not None else set())
        if missing:
            raise RuntimeError(f"{label}: {name} の r:id {sorted(missing)} が {_rels_name(name)} にない")


def drop_printer_settings(path: Path) -> None:
    """xlsx からプリンター設定のパート（xl/printerSettings/*）と、それを指す参照を取り除く。

    対象はワークシート（xl/worksheets）とグラフシート（xl/chartsheets）。rels は XML として読み、Type が
    プリンター設定の Relationship を属性の順に関係なく見つける。取り除いたあとに全シートの r:id が rels にあることを
    確かめ、なければ RuntimeError（元のファイルは変えない）。
    """
    path = Path(path)
    with zipfile.ZipFile(path) as src:
        entries = [(info, src.read(info.filename)) for info in src.infolist()]
    parts = {info.filename: data for info, data in entries}
    printer_ids = {name: _printer_rel_ids(parts[_rels_name(name)])
                   for name in parts if SHEET_PART.match(name) and _rels_name(name) in parts}

    out: dict[str, bytes] = {}
    for info, data in entries:
        name = info.filename
        if PRINTER_PART.match(name):
            continue
        if name == "[Content_Types].xml":
            data = PRINTER_CONTENT_TYPE.sub("", data.decode("utf-8")).encode("utf-8")
        elif printer_ids.get(name):
            data = _drop_page_setup_rel_ids(data.decode("utf-8"), printer_ids[name]).encode("utf-8")
        else:
            sheet = next((s for s in printer_ids if _rels_name(s) == name), None)
            if sheet and printer_ids[sheet]:
                data = _drop_relationships(data.decode("utf-8"), printer_ids[sheet]).encode("utf-8")
                left = _printer_rel_ids(data)
                if left:
                    raise RuntimeError(f"{path.name}: {name} のプリンター設定 {sorted(left)} を取り除けなかった")
        out[name] = data
    check_sheet_relationships(out, path.name)

    fd, tmp_name = tempfile.mkstemp(dir=path.parent, prefix=f".{path.stem}.", suffix=".tmp")
    tmp = Path(tmp_name)
    try:
        with os.fdopen(fd, "wb") as f, zipfile.ZipFile(f, "w", zipfile.ZIP_DEFLATED) as dst:
            for info, _ in entries:
                if info.filename in out:
                    dst.writestr(info, out[info.filename])
        os.replace(tmp, path)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


# --- PDF の個人情報 ---

PDF_PERSON_KEYS = re.compile(rb"/(Author|Creator|Producer|Title|Subject|Keywords)\s*(\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f]*>)")


def _pdf_string(raw: bytes) -> str:
    if raw.startswith(b"<"):
        data = bytes.fromhex(raw[1:-1].decode("ascii"))
    else:
        data = re.sub(rb"\\(.)", rb"\1", raw[1:-1])
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be", errors="replace")
    return data.decode("latin-1")


def pdf_info(path: Path) -> dict[str, str]:
    """PDF の文書情報（作成者・作成アプリなど）を読む。値は文字列にしたもの。"""
    return {m.group(1).decode(): _pdf_string(m.group(2)) for m in PDF_PERSON_KEYS.finditer(Path(path).read_bytes())}


PDF_STREAM = re.compile(
    rb"\bobj\s*(?P<dict><<(?:(?!endobj).)*?>>)\s*stream\r?\n(?P<body>.*?)\r?\nendstream", re.DOTALL
)


def _pdf_streams(data: bytes) -> list[tuple[bytes, bytes]]:
    """PDF のストリームを（辞書, 中身）の組で返す。FlateDecode は展開する（展開できないものは生のまま）。"""
    streams = []
    for m in PDF_STREAM.finditer(data):
        body = m.group("body")
        if b"/FlateDecode" in m.group("dict"):
            try:
                body = zlib.decompress(body)
            except zlib.error:
                # 画像など、展開しても読めない形式のストリームがある。生のまま検査する
                print(f"note: could not inflate a PDF stream ({len(body)} bytes); checking raw bytes", file=sys.stderr)
        streams.append((m.group("dict"), body))
    return streams


OFFICE_USER_INFO_KEY = r"Software\Microsoft\Office\Common\UserInfo"


def office_user_name_from_registry() -> str | None:
    """Office のユーザー名（「Microsoft Office のユーザー名」の設定）をレジストリから読む。取れなければ None。"""
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, OFFICE_USER_INFO_KEY) as key:
            value, _ = winreg.QueryValueEx(key, "UserName")
    except FileNotFoundError:
        # キーか値がない＝Office のユーザー名を設定していない。探す名前が1つ減るだけなので続ける
        return None
    return str(value) or None


def default_user_names(extra: Iterable[str | None] = ()) -> list[str]:
    """個人情報の検査で探す実行ユーザーの名前：ホームディレクトリ名・環境変数 USERNAME・Office のユーザー名
    （レジストリ）と、extra（Excel の Application.UserName など）。空のものと重複は除く。"""
    names: list[str] = []
    for name in (Path.home().name, os.environ.get("USERNAME"), office_user_name_from_registry(), *extra):
        if name and name not in names:
            names.append(name)
    return names


def verify_pdf_no_personal_info(path: Path, user_names: Iterable[str] | None = None) -> None:
    """PDF に作成者名・実行ユーザー名・保存先パスが入っていれば RuntimeError を送出する。

    user_names を省略すると default_user_names()（ホームディレクトリ名・USERNAME・Office のユーザー名）を探す。
    実行ユーザー名はストリームの中（展開したもの）も含めて探す。保存先パスは、ストリームの外（文書情報など）と
    XMP メタデータのストリームだけを見る（フォントや描画のストリームはバイナリで、誤検出が多いため）。
    """
    path = Path(path)
    if user_names is None:
        user_names = default_user_names()
    data = path.read_bytes()
    info = pdf_info(path)
    if info.get("Author"):
        raise RuntimeError(f"{path.name}: 作成者（/Author）が入っている: {info['Author']!r}")
    for key, value in info.items():
        m = PATH_PATTERN.search(value)
        if m:
            raise RuntimeError(f"{path.name}: /{key} に保存先パスらしい文字列がある（{m.group(0)!r}）")

    streams = _pdf_streams(data)
    haystacks = [data.lower()] + [body.lower() for _, body in streams]
    for user in (n for n in user_names if len(n) >= 3):
        encodings = (user.encode("utf-8"), user.encode("utf-16-be"), user.encode("utf-16-le"),
                     user.encode("utf-16-be").hex().encode())
        if any(e.lower() in h for e in encodings for h in haystacks):
            raise RuntimeError(f"{path.name}: 実行ユーザー名が含まれている")

    outside_streams = PDF_STREAM.sub(b"", data)
    metadata = [body for d, body in streams if b"/Metadata" in d]
    for text in [outside_streams, *metadata]:
        m = PATH_PATTERN.search(text.decode("latin-1"))
        if m:
            raise RuntimeError(f"{path.name}: 保存先パスらしい文字列がある（{m.group(0)!r}）")


def main() -> int:
    summary = summarize(make_sales.build_rows(), REPORT_MONTH.year, REPORT_MONTH.month)
    with staging.staging_dir() as stage:
        moves = {stage / TEMPLATE.name: TEMPLATE, stage / ANSWER.name: ANSWER, stage / ANSWER_PDF.name: ANSWER_PDF}
        template, answer, pdf = moves
        write_template(template)
        write_answer(summary, template, answer, pdf)
        staging.publish_all(moves)
    print(f"saved {TEMPLATE}\nsaved {ANSWER}\nsaved {ANSWER_PDF}\n")
    print("\n".join(summary_lines(summary)))
    print(f"\n特記事項：{remarks(summary)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
