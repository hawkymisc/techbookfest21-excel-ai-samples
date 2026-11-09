"""部署コード対応表（Issue #12）のサンプルを作る。仕様は docs/sample-data.md §1・§2・§2.1・§2.2。

8.4「部署コードがなければ付番する」の題材。会社には部署コードがなく、資料ごとに部署名の書き方が違う。

- 入力（samples/部署名一覧_付番前.xlsx）
  - 「部署名の書き方」シート：売上明細・予算表・代理店報告・内線番号表から部署名を抜き出した一覧（コードはない）
  - 「組織図」シート：本部・部・課の階層。付番の桁（本部・部・課）を決める手がかり
- 正解（samples/answers/部署コード対応表.xlsx）
  - 「部署コード」シート：§2.2 の4桁コードと部署名。売上明細の部署名も持つ（4.2 の辞書 Skill・10.2 で使う）
  - 「名寄せ」シート：入力の各行がどのコードに当たるか。人が確かめるべき書き方には備考を付ける

乱数は使わない（中身はすべてマスタと下の表から決まる）。一時フォルダに保存して個人情報を取り除き、残っていないことを
確認してから samples/ に移す（tools/staging.py）。

    uv run python tools/make_dept_codes.py
"""
import sys
from collections import defaultdict
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import sample_master as master
import staging
from xlsx_meta import scrub_personal_info, verify_no_personal_info

SAMPLES = Path(__file__).resolve().parent.parent / "samples"
INPUT_PATH = SAMPLES / "部署名一覧_付番前.xlsx"
ANSWER_PATH = SAMPLES / "answers" / "部署コード対応表.xlsx"

INPUT_SHEET = "部署名の書き方"
ORG_CHART_SHEET = "組織図"
CODE_SHEET = "部署コード"
MAPPING_SHEET = "名寄せ"

INPUT_HEADER = ["出典", "部署名（資料での書き方）"]
ORG_CHART_HEADER = ["本部", "部・工場", "課・拠点"]
CODE_HEADER = ["部署コード", "本部", "部・工場", "課・拠点", "部署名", "売上明細での部署名"]
MAPPING_HEADER = ["出典", "部署名（資料での書き方）", "部署コード", "部署名", "備考"]

# 資料（出典）。並びは入力シートでの並び
SOURCES = ("売上明細", "予算表", "代理店報告", "内線番号表")

# 内線番号表での書き方（部署コードの正解 → 書き方）。総務部が手で作っている表という想定で、
# 略称・数字の全角半角・空白の有無が揃っていない。マスタにない情報なのでこのスクリプトに持つ
EXTENSION_LIST_NAMES = {
    "1110": "総務課", "1120": "人事課", "1200": "経理部", "1300": "経営企画", "1400": "情報システム",
    "2110": "本社営業", "2120": "北関東営業所", "2130": "仙台営業所", "2210": "大阪支店",
    "2220": "名古屋支店", "2230": "広島営業所", "2240": "福岡営業所", "2300": "海外営業",
    "3110": "第1工場 製造1課", "3120": "第１工場 製造２課", "3130": "第一工場製造三課", "3210": "第二工場 製造課",
    "3300": "生産管理", "3400": "品証", "3500": "資材購買", "4100": "設計", "4200": "開発", "5100": "保守サービス",
}

# 名寄せで人が確かめるべき書き方（8.4 の確認ポイント）。書き方 → 備考
NOTES = {
    "東京": "代理店報告での略称。所在地で書いているので、本社の他部署と取り違えないこと",
    "本社": "代理店報告での略称。管理本部など本社にあるほかの部署ではなく本社営業課",
    "東北": "地域名で書かれている。東北の担当は仙台営業所",
    "中部": "地域名で書かれている。中部の担当は名古屋支店",
    "九州": "地域名で書かれている。九州の担当は福岡営業所",
    "品証": "品質保証部の略称",
    "情報システム": "管理本部直下の情報システム課",
    "第1工場 製造1課": "数字が半角",
    "第１工場 製造２課": "数字が全角",
    "保守サービス": "サービス部 保守サービス課",
}


class CodeRow(NamedTuple):
    code: str
    division: str
    department: str | None
    section: str | None
    name: str
    sales_name: str | None  # 売上明細での部署名（売上を立てる部署だけ）


class NameRow(NamedTuple):
    source: str
    notation: str


class MappingRow(NamedTuple):
    source: str
    notation: str
    code: str
    name: str
    note: str | None


class OrgChartRow(NamedTuple):
    division: str
    department: str | None
    section: str | None


def _check_extension_list() -> None:
    if set(EXTENSION_LIST_NAMES) != set(master.ORG_CODES):
        raise ValueError("EXTENSION_LIST_NAMES の部署コードがマスタ（sample_master.ORG_UNITS）と一致しない")


_check_extension_list()  # 起動時に確かめる


def build_code_table() -> list[CodeRow]:
    """正解の部署コード対応表（§2.2）。コード順。"""
    sales_by_code = {d.answer_code: d.name for d in master.SALES_DEPARTMENTS}
    return [CodeRow(u.code, u.division, u.department, u.section, u.name, sales_by_code.get(u.code))
            for u in sorted(master.ORG_UNITS, key=lambda u: u.code)]


def build_org_chart() -> list[OrgChartRow]:
    """入力に付ける組織図。コードは持たない。"""
    return [OrgChartRow(u.division, u.department, u.section) for u in master.ORG_UNITS]


def _notations_with_codes() -> list[tuple[str, str, str]]:
    """(出典, 書き方, 正解のコード) を入力シートの並びで返す。"""
    rows = []
    for d in master.SALES_DEPARTMENTS:
        rows.append(("売上明細", d.name, d.answer_code))
    for d in master.SALES_DEPARTMENTS:
        rows.append(("予算表", d.budget_name, d.answer_code))
    for d in master.SALES_DEPARTMENTS:
        rows.extend(("代理店報告", n, d.answer_code) for n in d.agency_report_names)
    for u in master.ORG_UNITS:
        rows.append(("内線番号表", EXTENSION_LIST_NAMES[u.code], u.code))
    return rows


def build_input_rows() -> list[NameRow]:
    """付番前の入力：資料ごとの部署名の書き方。"""
    return [NameRow(source, notation) for source, notation, _ in _notations_with_codes()]


def build_mapping() -> list[MappingRow]:
    """正解の名寄せ：入力の各行 → 部署コード。同じ書き方が別のコードに当たっていたら ValueError。"""
    seen: dict[str, str] = {}
    rows = []
    for source, notation, code in _notations_with_codes():
        if seen.setdefault(notation, code) != code:
            raise ValueError(f"書き方「{notation}」が {seen[notation]} と {code} の両方に当たる")
        rows.append(MappingRow(source, notation, code, master.ORG_CODES[code], NOTES.get(notation)))
    unknown = set(NOTES) - set(seen)
    if unknown:
        raise ValueError(f"NOTES に入力にない書き方がある: {sorted(unknown)}")
    return rows


def _write_sheet(sheet, header: list[str], rows: list[tuple], text_columns: str = "") -> None:
    last = len(rows) + 1
    last_col = chr(ord("A") + len(header) - 1)
    for col in text_columns:  # 文字列で持つ列は、値を入れる前に表示形式を「文字列」にする（数値に変換させない）
        sheet.range(f"{col}2:{col}{last}").number_format = "@"
    sheet.range("A1").value = [header] + [list(r) for r in rows]
    head = sheet.range(f"A1:{last_col}1")
    head.font.bold = True
    head.color = (221, 235, 247)
    sheet.range(f"A1:{last_col}{last}").columns.autofit()


def _new_book(app: xw.App, sheet_names: list[str]):
    book = app.books.add()
    while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚から始める
        book.sheets[-1].delete()
    book.sheets[0].name = sheet_names[0]
    for name in sheet_names[1:]:
        book.sheets.add(name, after=book.sheets[-1])
    return book


def _save(book, path: Path) -> None:
    book.sheets[0].activate()
    book.api.RemovePersonalInformation = True
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(str(path))
    book.close()


def write_workbooks(input_path: Path, answer_path: Path) -> None:
    with xw.App(visible=False, add_book=False) as app:
        book = _new_book(app, [INPUT_SHEET, ORG_CHART_SHEET])
        _write_sheet(book.sheets[INPUT_SHEET], INPUT_HEADER, build_input_rows())
        _write_sheet(book.sheets[ORG_CHART_SHEET], ORG_CHART_HEADER, build_org_chart())
        _save(book, input_path)

        book = _new_book(app, [CODE_SHEET, MAPPING_SHEET])
        _write_sheet(book.sheets[CODE_SHEET], CODE_HEADER, build_code_table(), text_columns="A")
        _write_sheet(book.sheets[MAPPING_SHEET], MAPPING_HEADER, build_mapping(), text_columns="C")
        _save(book, answer_path)


def print_summary() -> None:
    inputs = build_input_rows()
    mapping = build_mapping()
    print(f"入力 {len(inputs)} 行（書き方 {len({r.notation for r in inputs})} 通り）　部署 {len(build_code_table())}")
    for source in SOURCES:
        print(f"  {source}\t{sum(r.source == source for r in inputs)} 行")
    notations = defaultdict(list)
    for m in mapping:
        if m.notation not in notations[m.code]:
            notations[m.code].append(m.notation)
    print("\nコード\t部署名\t書き方の数\t書き方")
    for r in build_code_table():
        print(f"{r.code}\t{r.name}\t{len(notations[r.code])}\t{' / '.join(notations[r.code])}")


def main() -> int:
    with staging.staging_dir() as stage:
        moves = {stage / INPUT_PATH.name: INPUT_PATH, stage / ANSWER_PATH.name: ANSWER_PATH}
        write_workbooks(*moves)
        for path in moves:
            scrub_personal_info(path)
            verify_no_personal_info(path)
        staging.publish_all(moves)
    for path in moves.values():
        print(f"saved {path} (personal info removed)")
    print()
    print_summary()
    return 0


if __name__ == "__main__":
    sys.exit(main())
