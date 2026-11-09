"""表記ゆれを含む取引先リスト（Issue #13）のサンプルを作る。仕様は docs/sample-data.md §5。

8.1「表記ゆれ・全角半角・日付形式の統一」の題材。売上明細と同じ得意先120社を、営業が手入力で管理している
別の名簿として作る。売上明細（販売管理システムのマスタから出力）では表記が揃っているが、この名簿では次が混ざる。

- 取引先名：「株式会社」と「㈱」「（株）」「(株)」、前株・後株の入れ替わり、法人格の書き漏れ、半角カナ、余計な空白
- 日付：`2026/1/5`・`2026/01/05`・`2026-01-05`・`R8.1.5`（令和以前は `H30.4.1`）・`2026年1月5日`・全角数字
- 締め日：「20日」と「２０日」、「末日」と「月末」

日付と締め日は、表示形式を「文字列」にした列に入っている（CSV から貼り付けた名簿という想定）。
Excel の日付に変換された値だと、書き方の違いが値ではなく表示形式の違いになり、8.1 の題材にならないため。

- 入力：samples/取引先リスト_手入力.xlsx
- 正解：samples/answers/取引先リスト_整形後.xlsx（取引先名は売上明細の表記、日付は Excel の日付で yyyy/mm/dd、
  締め日は半角数字の「N日」か「末日」。整形ルールを別シートに書く）

乱数は固定シードなので、何度実行しても同じ内容になる。一時フォルダに保存して個人情報を取り除き、残っていないことを
確認してから samples/ に移す（tools/staging.py）。

    uv run python tools/make_customer_list.py
"""
import datetime as dt
import random
import sys
import unicodedata
from collections import Counter
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import make_sales
import sample_master as master
import staging
from xlsx_meta import scrub_personal_info, verify_no_personal_info

SAMPLES = Path(__file__).resolve().parent.parent / "samples"
RAW_PATH = SAMPLES / "取引先リスト_手入力.xlsx"
ANSWER_PATH = SAMPLES / "answers" / "取引先リスト_整形後.xlsx"
SHEET_NAME = "取引先リスト"
RULES_SHEET = "整形ルール"
HEADER = ["No", "取引先名", "担当部署", "取引開始日", "締め日", "最終更新日"]

SEED = 20261113

# 取引開始日・最終更新日の範囲。本文の「今日」は2026年10月なので、どちらも2026年9月までにする。
# 取引開始日は売上明細の初日（make_sales.START＝2024年10月1日）の前日まで。売上明細の得意先は全社が
# 初月から売上を持つとは限らないが、「取引開始日 ≤ その得意先の最初の売上日」がどの社でも成り立つ。
# 令和（R8 など）の書き方の例は、最終更新日（2024年10月〜2026年9月）の側に出る
START_FROM = dt.date(2004, 4, 1)
START_UNTIL = make_sales.START - dt.timedelta(days=1)
UPDATED_FROM = dt.date(2024, 10, 1)
LAST_DAY = dt.date(2026, 9, 30)
REIWA_START = dt.date(2019, 5, 1)

# --- 取引先名の揺れ ---
# 法人格の書き方（会社ごとにどれか1つ）。自治体（上下水道局・水道課）は法人格がないので対象外。
# 残りの会社は売上明細と同じ書き方
LEGAL_FORM_QUOTA = {"㈱": 14, "（株）": 5, "(株)": 3, "前株後株の入れ替え": 7, "法人格なし": 7}
ABBREVIATIONS = {
    "㈱": {"株式会社": "㈱", "有限会社": "㈲"},
    "（株）": {"株式会社": "（株）", "有限会社": "（有）"},
    "(株)": {"株式会社": "(株)", "有限会社": "(有)"},
}
# 法人格と関係なく重なる揺れ（件数）
HALFWIDTH_KANA_QUOTA = 10   # 社名のカタカナを半角カナで書く
INNER_SPACE_QUOTA = 6       # 法人格と社名のあいだに空白（半角か全角）
TRAILING_SPACE_QUOTA = 4    # 末尾に半角空白（見えない揺れ）
MUNICIPAL_SPACE_QUOTA = 3   # 「〇〇市 上下水道局」のように自治体名のあとに空白

# --- 日付の書き方（列ごとの件数。合計は得意先の数） ---
DATE_FORMAT_QUOTA = {
    "2026/1/5": 45, "2026/01/05": 10, "2026-01-05": 30, "R8.1.5": 15, "2026年1月5日": 10, "２０２６/１/５": 10,
}

# --- 締め日 ---
CLOSING_WEIGHTS = {"10日": 10, "15日": 15, "20日": 35, "25日": 15, "末日": 25}
CLOSING_FULLWIDTH_QUOTA = 12  # 「２０日」
CLOSING_GETSUMATSU_QUOTA = 6  # 「月末」

_FULLWIDTH_DIGITS = str.maketrans("0123456789", "０１２３４５６７８９")


def _halfwidth_kana_table() -> dict[str, str]:
    """全角カタカナ → 半角カナ（濁点・半濁点は2文字）。NFKC で全角に戻ることを使って作る。"""
    table = {}
    halfwidth = [chr(c) for c in range(0xFF66, 0xFF9E)]
    for h in halfwidth:
        table.setdefault(unicodedata.normalize("NFKC", h), h)
        for mark in ("ﾞ", "ﾟ"):
            full = unicodedata.normalize("NFKC", h + mark)
            if len(full) == 1:
                table.setdefault(full, h + mark)
    return table


HALFWIDTH_KANA = _halfwidth_kana_table()


def to_halfwidth_kana(text: str) -> str:
    return "".join(HALFWIDTH_KANA.get(ch, ch) for ch in text)


def _has_katakana(text: str) -> bool:
    return any("ァ" <= ch <= "ヺ" for ch in text)


class Entry(NamedTuple):
    no: int
    customer: master.Customer
    raw_name: str
    name_variations: tuple[str, ...]
    start: dt.date
    start_format: str
    closing: str
    closing_text: str
    updated: dt.date
    updated_format: str

    def raw_row(self) -> tuple:
        return (self.no, self.raw_name, self.customer.department,
                format_date(self.start, self.start_format), self.closing_text,
                format_date(self.updated, self.updated_format))

    def answer_row(self) -> tuple:
        return (self.no, self.customer.name, self.customer.department, self.start, self.closing, self.updated)


def format_date(day: dt.date, fmt: str) -> str:
    if fmt == "2026/1/5":
        return f"{day.year}/{day.month}/{day.day}"
    if fmt == "2026/01/05":
        return f"{day:%Y/%m/%d}"
    if fmt == "2026-01-05":
        return day.isoformat()
    if fmt == "R8.1.5":
        era, year = ("R", day.year - 2018) if day >= REIWA_START else ("H", day.year - 1988)
        return f"{era}{year}.{day.month}.{day.day}"
    if fmt == "2026年1月5日":
        return f"{day.year}年{day.month}月{day.day}日"
    if fmt == "２０２６/１/５":
        return f"{day.year}/{day.month}/{day.day}".translate(_FULLWIDTH_DIGITS)
    raise ValueError(f"日付の書き方が不明: {fmt}")


def _split_legal_form(name: str) -> tuple[str | None, str, str | None]:
    """社名を (法人格, 社名の本体, 位置 front/back) に分ける。法人格がなければ (None, name, None)。"""
    for form in ("株式会社", "有限会社"):
        if name.startswith(form):
            return form, name[len(form):], "front"
        if name.endswith(form):
            return form, name[:-len(form)], "back"
    return None, name, None


def _assign(rng: random.Random, candidates: list[int], quota: dict[str, int], label: str) -> dict[int, str]:
    """候補の番号に、件数どおりに種類を割り当てる。候補が足りなければ ValueError。"""
    total = sum(quota.values())
    if total > len(candidates):
        raise ValueError(f"{label}: 割り当てる件数 {total} に対して候補が {len(candidates)} しかない")
    chosen = rng.sample(candidates, total)
    kinds = [kind for kind, n in quota.items() for _ in range(n)]
    return dict(zip(chosen, kinds))


def _raw_names(rng: random.Random, customers: tuple[master.Customer, ...]) -> list[tuple[str, tuple[str, ...]]]:
    parts = [_split_legal_form(c.name) for c in customers]
    companies = [i for i, (form, _, _) in enumerate(parts) if form]
    municipal = [i for i, (form, _, _) in enumerate(parts) if not form]
    legal = _assign(rng, companies, LEGAL_FORM_QUOTA, "法人格の書き方")
    kana = set(rng.sample([i for i, c in enumerate(customers) if _has_katakana(c.name)], HALFWIDTH_KANA_QUOTA))
    written = [i for i in companies if legal.get(i) != "法人格なし"]
    inner = set(rng.sample(written, INNER_SPACE_QUOTA))
    trailing = set(rng.sample(range(len(customers)), TRAILING_SPACE_QUOTA))
    municipal_space = set(rng.sample(municipal, MUNICIPAL_SPACE_QUOTA))

    names = []
    for i, c in enumerate(customers):
        form, body, position = parts[i]
        tags: list[str] = []
        if i in kana:
            body = to_halfwidth_kana(body)
            tags.append("半角カナ")
        if form is None:
            if i in municipal_space:
                head, sep, tail = body.partition("市") if "市" in body else body.partition("町")
                body = f"{head}{sep} {tail}"
                tags.append("自治体名の後の空白")
            name = body
        else:
            kind = legal.get(i)
            written_form = form
            if kind in ABBREVIATIONS:
                written_form = ABBREVIATIONS[kind][form]
                tags.append(written_form)  # 有限会社なら ㈲・（有）・(有)
            elif kind == "前株後株の入れ替え":
                position = "back" if position == "front" else "front"
                tags.append(kind)
            if kind == "法人格なし":
                name = body
                tags.append(kind)
            else:
                sep = ""
                if i in inner:
                    sep = rng.choice((" ", "　"))
                    tags.append("法人格の前後の空白")
                name = f"{written_form}{sep}{body}" if position == "front" else f"{body}{sep}{written_form}"
        if i in trailing:
            name += " "
            tags.append("末尾の空白")
        names.append((name, tuple(tags)))
    return names


def _random_day(rng: random.Random, start: dt.date, end: dt.date) -> dt.date:
    return start + dt.timedelta(days=rng.randint(0, (end - start).days))


def _closings(rng: random.Random, n: int) -> list[tuple[str, str]]:
    closings = rng.choices(list(CLOSING_WEIGHTS), weights=list(CLOSING_WEIGHTS.values()), k=n)
    texts = list(closings)
    numeric = [i for i, c in enumerate(closings) if c != "末日"]
    month_end = [i for i, c in enumerate(closings) if c == "末日"]
    for i in _assign(rng, numeric, {"全角": CLOSING_FULLWIDTH_QUOTA}, "締め日（全角）"):
        texts[i] = closings[i].translate(_FULLWIDTH_DIGITS)
    for i in _assign(rng, month_end, {"月末": CLOSING_GETSUMATSU_QUOTA}, "締め日（月末）"):
        texts[i] = "月末"
    return list(zip(closings, texts))


def build_entries(seed: int = SEED) -> list[Entry]:
    """取引先リストの行を作る。同じシードなら同じ内容になる。No は取引開始日の順。"""
    rng = random.Random(seed)
    customers = master.CUSTOMERS
    n = len(customers)
    if sum(DATE_FORMAT_QUOTA.values()) != n:
        raise ValueError(f"DATE_FORMAT_QUOTA の合計が得意先の数 {n} と一致しない")
    names = _raw_names(rng, customers)
    starts = [_random_day(rng, START_FROM, START_UNTIL) for _ in customers]
    updated = [_random_day(rng, max(s, UPDATED_FROM), LAST_DAY) for s in starts]
    start_formats = _assign(rng, list(range(n)), DATE_FORMAT_QUOTA, "取引開始日の書き方")
    updated_formats = _assign(rng, list(range(n)), DATE_FORMAT_QUOTA, "最終更新日の書き方")
    closings = _closings(rng, n)

    order = sorted(range(n), key=lambda i: (starts[i], customers[i].name))
    return [
        Entry(no, customers[i], names[i][0], names[i][1], starts[i], start_formats[i],
              closings[i][0], closings[i][1], updated[i], updated_formats[i])
        for no, i in enumerate(order, start=1)
    ]


RULES = [
    "整形ルール（8.1 の正解で使ったもの）",
    "取引先名：売上明細（販売管理システム）の得意先名の表記にそろえる",
    "　・㈱・（株）・(株) は「株式会社」、㈲・（有）・(有) は「有限会社」に戻す",
    "　・前株・後株、法人格の書き漏れは、名簿だけでは決められないので売上明細の得意先名に合わせる",
    "　・半角カナは全角に、余計な空白（前後・途中）は取る",
    "取引開始日・最終更新日：Excel の日付（シリアル値）にして、表示形式を yyyy/mm/dd にする",
    "　・R は令和（R1＝2019年）、H は平成（H1＝1989年）として読む。全角数字は半角に",
    "締め日：半角数字の「N日」か「末日」にする（「月末」は「末日」）",
    "担当部署：もとから売上明細の部署名で書かれているので変えない",
]


def _write_list(sheet, rows: list[tuple], text_columns: str, date_columns: str) -> None:
    last = len(rows) + 1
    for col in text_columns:  # 値を入れる前に表示形式を「文字列」にして、Excel に日付へ変換させない
        sheet.range(f"{col}2:{col}{last}").number_format = "@"
    sheet.range("A1").value = [HEADER] + [list(r) for r in rows]
    for col in date_columns:
        sheet.range(f"{col}2:{col}{last}").number_format = "yyyy/mm/dd"
    head = sheet.range("A1:F1")
    head.font.bold = True
    head.color = (221, 235, 247)
    sheet.range(f"A1:F{last}").columns.autofit()


def _new_book(app: xw.App, sheet_name: str):
    book = app.books.add()
    while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
        book.sheets[-1].delete()
    book.sheets[0].name = sheet_name
    return book


def _save(book, path: Path) -> None:
    book.sheets[0].activate()
    book.api.RemovePersonalInformation = True
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(str(path))
    book.close()


def write_workbooks(entries: list[Entry], raw_path: Path, answer_path: Path) -> None:
    with xw.App(visible=False, add_book=False) as app:
        book = _new_book(app, SHEET_NAME)
        _write_list(book.sheets[0], [e.raw_row() for e in entries], text_columns="DEF", date_columns="")
        _save(book, raw_path)

        book = _new_book(app, SHEET_NAME)
        _write_list(book.sheets[0], [e.answer_row() for e in entries], text_columns="", date_columns="DF")
        rules = book.sheets.add(RULES_SHEET, after=book.sheets[0])
        rules.range("A1").value = [[line] for line in RULES]
        rules.range("A1").font.bold = True
        _save(book, answer_path)


def print_summary(entries: list[Entry]) -> None:
    raws = [e.raw_row() for e in entries]
    clean = sum(e.raw_name == e.customer.name for e in entries)
    print(f"取引先 {len(entries)} 社　取引先名が売上明細と同じ {clean} 社　揺れのある社 {len(entries) - clean} 社")
    print("\n取引先名の揺れ（1社に複数重なることがある）")
    for tag, count in Counter(t for e in entries for t in e.name_variations).most_common():
        print(f"  {tag}\t{count}")
    for label, col, fmt in (("取引開始日", 3, "start_format"), ("最終更新日", 5, "updated_format")):
        counts = Counter(getattr(e, fmt) for e in entries)
        heisei = sum(1 for r in raws if r[col].startswith("H"))
        print(f"\n{label}の書き方（うち平成 {heisei} 件）")
        for f in DATE_FORMAT_QUOTA:
            print(f"  {f}\t{counts[f]}")
    print("\n締め日（正解 → 件数、うち全角・月末）")
    for c in CLOSING_WEIGHTS:
        rows = [e for e in entries if e.closing == c]
        print(f"  {c}\t{len(rows)}\t全角 {sum(e.closing_text != c and e.closing_text != '月末' for e in rows)}"
              f"\t月末 {sum(e.closing_text == '月末' for e in rows)}")
    print(f"\n取引開始日 {entries[0].start} 〜 {entries[-1].start}")


def main() -> int:
    entries = build_entries()
    with staging.staging_dir() as stage:
        moves = {stage / RAW_PATH.name: RAW_PATH, stage / ANSWER_PATH.name: ANSWER_PATH}
        write_workbooks(entries, *moves)
        for path in moves:
            scrub_personal_info(path)
            verify_no_personal_info(path)
        staging.publish_all(moves)
    for path in moves.values():
        print(f"saved {path} (personal info removed)")
    print()
    print_summary(entries)
    return 0


if __name__ == "__main__":
    sys.exit(main())
