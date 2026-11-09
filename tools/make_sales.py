"""売上明細（Issue #11）のサンプルブックを作る。仕様は docs/sample-data.md §4。

架空の会社「株式会社ミナセ流体工業」の 2024年10月〜2026年9月の売上を、伝票の明細行単位で作る。
乱数は固定シードなので、何度実行しても同じ内容になる。
読者に配るファイルなので、保存後に作成者名・保存先パスなどの個人情報を取り除き、残っていないことを確認する。
最後に、原稿の検算に使う集計（部署別・月別の金額合計、総行数、総額）を標準出力に出す。

    uv run python tools/make_sales.py

部署・月ごとの計画（MONTH_FACTOR・FY_GROWTH・DEPARTMENT_SHARE・HIROSHIMA_FACTOR など）や商品の出方を変えたら、
シードを選び直す。行ごとの金額のばらつきで、シードによっては傾向（広島営業所だけが前年割れ等）が崩れるため。

1. SEED を 20241001 から1ずつ増やしながら `uv run pytest tests/test_make_sales.py` を回し、
   傾向のテスト（広島・全社・部署ごとの前年比、季節性、前月比）がすべて通る最初のシードを採る
2. `uv run python tools/make_sales.py` でブックを作り直す
3. tests/test_golden.py のゴールデン値（総行数・総額・返品行数・ハッシュ）を、出力に合わせて更新する
"""
import datetime as dt
import random
import sys
from collections import defaultdict
from pathlib import Path
from typing import NamedTuple

import xlwings as xw

import sample_master as master
from xlsx_meta import scrub_personal_info, verify_no_personal_info

OUTPUT = Path(__file__).resolve().parent.parent / "samples" / "売上明細_2024年10月-2026年9月.xlsx"
SHEET_NAME = "売上明細"
# 計画どおりの傾向（広島営業所だけが前年割れ、全社は数%増）が、行ごとのばらつきに埋もれずに出るシードを選んだ。
# 選び直す手順はモジュールの docstring を参照
SEED = 20241001

START = dt.date(2024, 10, 1)
END = dt.date(2026, 9, 30)

HEADER = [
    "売上日", "伝票番号", "部署名", "得意先名", "販売経路", "代理店名",
    "商品コード", "商品名", "商品区分", "数量", "単価", "金額",
]

# 国民の祝日・休日（内閣府「国民の祝日について」の公表値。2024-10〜2026-09 の分だけ）
HOLIDAYS = (
    dt.date(2024, 10, 14), dt.date(2024, 11, 3), dt.date(2024, 11, 4), dt.date(2024, 11, 23),
    dt.date(2025, 1, 1), dt.date(2025, 1, 13), dt.date(2025, 2, 11), dt.date(2025, 2, 23),
    dt.date(2025, 2, 24), dt.date(2025, 3, 20), dt.date(2025, 4, 29), dt.date(2025, 5, 3),
    dt.date(2025, 5, 4), dt.date(2025, 5, 5), dt.date(2025, 5, 6), dt.date(2025, 7, 21),
    dt.date(2025, 8, 11), dt.date(2025, 9, 15), dt.date(2025, 9, 23), dt.date(2025, 10, 13),
    dt.date(2025, 11, 3), dt.date(2025, 11, 23), dt.date(2025, 11, 24),
    dt.date(2026, 1, 1), dt.date(2026, 1, 12), dt.date(2026, 2, 11), dt.date(2026, 2, 23),
    dt.date(2026, 3, 20), dt.date(2026, 4, 29), dt.date(2026, 5, 3), dt.date(2026, 5, 4),
    dt.date(2026, 5, 5), dt.date(2026, 5, 6), dt.date(2026, 7, 20), dt.date(2026, 8, 11),
    dt.date(2026, 9, 21), dt.date(2026, 9, 22), dt.date(2026, 9, 23),
)
_HOLIDAY_SET = frozenset(HOLIDAYS)

# 24か月の総行数の目安（返品行を除く）
TARGET_ROWS = 7_000
# 部署・月ごとの売上計画（改定前の単価で数えた金額）を、次の係数の積で決める。
# 明細行は、その月の計画に届くまで作る。行ごとの金額のばらつきが大きくても、部署・月の合計が計画から外れにくい。
# 月ごとの係数。3月（年度末）と9月（上期末）が多く、8月と1月が少ない（営業日数の違いとは別にかかる）
MONTH_FACTOR = {1: 0.85, 2: 1.00, 3: 1.30, 4: 0.95, 5: 0.95, 6: 1.00,
                7: 1.00, 8: 0.85, 9: 1.25, 10: 1.00, 11: 1.00, 12: 1.05}
# 年度ごとの係数（全社で前年比数%伸びる。2025年10月の単価改定の効果は別に乗る）
FY_GROWTH = {2024: 1.00, 2025: 1.05, 2026: 1.10}
# 部署ごとの比率。本社営業課と大阪支店が大きく、仙台営業所と広島営業所が小さい
DEPARTMENT_SHARE = {
    "本社営業課": 0.22, "北関東営業所": 0.11, "仙台営業所": 0.07, "大阪支店": 0.20,
    "名古屋支店": 0.14, "広島営業所": 0.07, "福岡営業所": 0.10, "海外営業部": 0.09,
}
# 広島営業所は2025年度に1割ほど落ち、2026年度上期も落ち込みが続く。
# 全社の伸びと単価改定を打ち消して、前年同期比が2025年度下期・2026年度上期とも0.9前後になる。
# 本文の「今日」は2026年10月なので、直近の月（2026年9月）でも前年同月を下回るようにしている
HIROSHIMA_FACTOR = {2024: 1.00, 2025: 0.85, 2026: 0.72}


def check_departments(table: dict[str, object]) -> None:
    """部署ごとの表が、マスタの売上部署と同じ部署を同じ順番で持っていることを確かめる。違えば ValueError。"""
    if list(table) != list(master.SALES_DEPARTMENT_NAMES):
        raise ValueError(f"部署の並びがマスタ（sample_master.SALES_DEPARTMENTS）と一致しない: {list(table)}")


check_departments(DEPARTMENT_SHARE)  # 起動時に確かめる

# 部署・月ごとの計画に掛けるゆらぎ（標準偏差）。前月比マイナスの月が自然にできる
MONTHLY_NOISE = 0.03
# 販売経路が代理店になる割合
AGENCY_SHARE = {"海外営業部": 0.40}
DEFAULT_AGENCY_SHARE = 0.30
# 1伝票あたりの明細行数
LINES_PER_SLIP = {1: 0.65, 2: 0.25, 3: 0.10}
# 商品区分ごとの出やすさと数量の分布
CATEGORY_WEIGHT = {"ポンプ": 0.18, "バルブ": 0.30, "交換部品": 0.37, "保守サービス": 0.15}
QUANTITY = {
    "ポンプ": ((1, 2, 3, 4, 5), (50, 25, 12, 8, 5)),
    "バルブ": (tuple(range(1, 11)), (20, 18, 14, 12, 10, 8, 6, 5, 4, 3)),
    "交換部品": (tuple(range(1, 21)), tuple(max(1, 20 - i) for i in range(20))),
    "保守サービス": ((1, 2, 3), (80, 15, 5)),
}
# 返品行（数量マイナス）の割合と、元の売上から返品までの営業日数
RETURN_RATE = 0.005
RETURN_DELAY_DAYS = (3, 20)


class SalesRow(NamedTuple):
    date: dt.date
    slip_no: str
    department: str
    customer: str
    channel: str
    agency: str | None
    product_code: str
    product_name: str
    category: str
    quantity: int
    unit_price: int
    amount: int

    def as_list(self) -> list:
        return list(self)


class _Line(NamedTuple):
    """伝票番号を振る前の明細行。slip は伝票を区別するための通し番号。"""
    date: dt.date
    slip: int
    department: str
    customer: str
    channel: str
    agency: str | None
    product: master.Product
    quantity: int
    unit_price: int


def is_business_day(day: dt.date) -> bool:
    if day.weekday() >= 5 or day in _HOLIDAY_SET:
        return False
    if (day.month, day.day) >= (12, 29) or (day.month, day.day) <= (1, 3):  # 年末年始
        return False
    if day.month == 8 and 13 <= day.day <= 16:  # 夏季休業
        return False
    return True


def business_days(start: dt.date = START, end: dt.date = END) -> list[dt.date]:
    days = []
    day = start
    while day <= end:
        if is_business_day(day):
            days.append(day)
        day += dt.timedelta(days=1)
    return days


def fiscal_year(day: dt.date) -> int:
    return day.year if day.month >= 4 else day.year - 1


def _expected_line_amount() -> float:
    """明細行1行の金額の期待値（改定前の単価）。計画金額から行数を見積もるのに使う。"""
    expected = 0.0
    for category, weight in CATEGORY_WEIGHT.items():
        prices = [p.price for p in master.PRODUCTS if p.category == category]
        values, weights = QUANTITY[category]
        mean_quantity = sum(v * w for v, w in zip(values, weights)) / sum(weights)
        expected += weight * (sum(prices) / len(prices)) * mean_quantity
    return expected


def _monthly_plans(rng: random.Random, days: list[dt.date]) -> dict[tuple[str, int, int], float]:
    """部署・月ごとの売上計画（改定前の単価で数えた金額）。全体で TARGET_ROWS 行ほどになるように合わせる。"""
    days_in_month: dict[tuple[int, int], int] = defaultdict(int)
    for d in days:
        days_in_month[(d.year, d.month)] += 1
    weights = {}
    for (year, month), n in days_in_month.items():
        fy = fiscal_year(dt.date(year, month, 1))
        for dept, share in DEPARTMENT_SHARE.items():
            hiroshima = HIROSHIMA_FACTOR[fy] if dept == "広島営業所" else 1.0
            weights[(dept, year, month)] = share * hiroshima * MONTH_FACTOR[month] * FY_GROWTH[fy] * n
    scale = TARGET_ROWS * _expected_line_amount() / sum(weights.values())
    return {key: w * scale * max(0.5, rng.gauss(1.0, MONTHLY_NOISE)) for key, w in weights.items()}


def _sales_lines(rng: random.Random, days: list[dt.date]) -> list[_Line]:
    customers_by_dept = defaultdict(list)
    for c in master.CUSTOMERS:
        customers_by_dept[c.department].append(c.name)
    agencies_by_dept = defaultdict(list)
    for a in master.AGENCIES:
        agencies_by_dept[a.department].append(a.name)
    products_by_category = defaultdict(list)
    for p in master.PRODUCTS:
        products_by_category[p.category].append(p)
    days_by_month = defaultdict(list)
    for d in days:
        days_by_month[(d.year, d.month)].append(d)
    categories = list(CATEGORY_WEIGHT)
    category_weights = list(CATEGORY_WEIGHT.values())

    slips_by_month: dict[tuple[int, int], list[list[_Line]]] = defaultdict(list)
    for (dept, year, month), plan in _monthly_plans(rng, days).items():
        first_day = days_by_month[(year, month)][0]  # 仮の日付。営業日は後で振り、単価もそのときに決める
        booked = 0
        full = False
        while not full:
            customer = rng.choice(customers_by_dept[dept])
            if rng.random() < AGENCY_SHARE.get(dept, DEFAULT_AGENCY_SHARE):
                channel, agency = "代理店", rng.choice(agencies_by_dept[dept])
            else:
                channel, agency = "直販", None
            n = rng.choices(list(LINES_PER_SLIP), weights=list(LINES_PER_SLIP.values()))[0]
            slip: list[_Line] = []
            used: set[str] = set()
            while len(slip) < n:
                category = rng.choices(categories, weights=category_weights)[0]
                product = rng.choice(products_by_category[category])
                if product.code in used:
                    continue
                used.add(product.code)
                values, weights = QUANTITY[category]
                quantity = rng.choices(values, weights=weights)[0]
                amount = quantity * product.price
                if booked + amount >= plan:
                    # 計画を超える行は、入れたほうが計画に近づくときだけ入れて、その月を締める
                    full = True
                    if booked + amount - plan > plan - booked:
                        break
                slip.append(_Line(first_day, 0, dept, customer, channel, agency, product, quantity, 0))
                booked += amount
                if full:
                    break
            if slip:
                slips_by_month[(year, month)].append(slip)

    # 月の伝票を混ぜてから営業日に均等に割り振る（1日の件数が月の中でそろい、同じ日の並びも部署順にならない）
    slips: list[list[_Line]] = []
    for month, month_slips in sorted(slips_by_month.items()):
        month_days = days_by_month[month]
        rng.shuffle(month_slips)
        for i, slip in enumerate(month_slips):
            day = month_days[i * len(month_days) // len(month_slips)]
            slips.append([line._replace(date=day, unit_price=master.unit_price(line.product, day)) for line in slip])
    return [line._replace(slip=i) for i, slip in enumerate(slips, start=1) for line in slip]


def _return_lines(rng: random.Random, sales: list[_Line], days: list[dt.date], first_slip: int) -> list[_Line]:
    """元の売上の一部を、数日〜数週間後に返品として計上する。単価は元の売上のもの。"""
    day_index = {d: i for i, d in enumerate(days)}
    count = round(len(sales) * RETURN_RATE)
    eligible = [s for s in sales if day_index[s.date] + RETURN_DELAY_DAYS[1] < len(days)]
    originals = sorted(rng.sample(range(len(eligible)), count))
    returns = []
    for offset, i in enumerate(originals):
        s = eligible[i]
        day = days[day_index[s.date] + rng.randint(*RETURN_DELAY_DAYS)]
        quantity = -rng.randint(1, s.quantity)
        returns.append(s._replace(date=day, slip=first_slip + offset, quantity=quantity))
    return returns


def _number_slips(lines: list[_Line]) -> list[SalesRow]:
    """日付順・伝票順に並べ、月ごとに 00001 から伝票番号を振る。"""
    lines = sorted(lines, key=lambda x: (x.date, x.slip))
    slip_numbers: dict[int, str] = {}
    next_seq: dict[tuple[int, int], int] = defaultdict(int)
    rows = []
    for x in lines:
        if x.slip not in slip_numbers:
            month = (x.date.year, x.date.month)
            next_seq[month] += 1
            slip_numbers[x.slip] = f"{x.date:%y%m}-{next_seq[month]:05d}"
        rows.append(SalesRow(
            x.date, slip_numbers[x.slip], x.department, x.customer, x.channel, x.agency,
            x.product.code, x.product.name, x.product.category,
            x.quantity, x.unit_price, x.quantity * x.unit_price,
        ))
    return rows


def build_rows(seed: int = SEED) -> list[SalesRow]:
    """売上明細の行を作る。同じシードなら同じ内容になる。"""
    rng = random.Random(seed)
    days = business_days()
    sales = _sales_lines(rng, days)
    returns = _return_lines(rng, sales, days, first_slip=max(line.slip for line in sales) + 1)
    return _number_slips(sales + returns)


def write_workbook(rows: list[SalesRow], path: Path) -> None:
    with xw.App(visible=False, add_book=False) as app:
        book = app.books.add()
        while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
            book.sheets[-1].delete()
        sheet = book.sheets[0]
        sheet.name = SHEET_NAME
        sheet.range("A1").value = [HEADER] + [r.as_list() for r in rows]
        last = len(rows) + 1
        header = sheet.range("A1:L1")
        header.font.bold = True
        header.color = (221, 235, 247)
        sheet.range(f"A2:A{last}").number_format = "yyyy/mm/dd"
        sheet.range(f"J2:L{last}").number_format = "#,##0"
        sheet.range(f"A1:L{last}").columns.autofit()
        book.api.RemovePersonalInformation = True
        book.save(str(path))
        book.close()


def print_summary(rows: list[SalesRow]) -> None:
    total = sum(r.amount for r in rows)
    returns = [r for r in rows if r.quantity < 0]
    print(f"総行数 {len(rows):,}　総額 {total:,}　返品行 {len(returns)}（{sum(r.amount for r in returns):,}）")

    periods = [("2024年度下期", dt.date(2024, 10, 1), dt.date(2025, 3, 31)),
               ("2025年度", dt.date(2025, 4, 1), dt.date(2026, 3, 31)),
               ("2026年度上期", dt.date(2026, 4, 1), dt.date(2026, 9, 30))]
    print("\n部署別の金額合計")
    print("部署名\t" + "\t".join(p[0] for p in periods) + "\t合計")
    for dept in master.SALES_DEPARTMENT_NAMES:
        sums = [sum(r.amount for r in rows if r.department == dept and s <= r.date <= e) for _, s, e in periods]
        print(f"{dept}\t" + "\t".join(f"{v:,}" for v in sums) + f"\t{sum(sums):,}")
    sums = [sum(r.amount for r in rows if s <= r.date <= e) for _, s, e in periods]
    print("全社\t" + "\t".join(f"{v:,}" for v in sums) + f"\t{sum(sums):,}")

    print("\n月別の件数と金額合計")
    print("年月\t件数\t金額")
    monthly: dict[str, list[int]] = defaultdict(lambda: [0, 0])
    for r in rows:
        m = monthly[f"{r.date:%Y/%m}"]
        m[0] += 1
        m[1] += r.amount
    for month, (count, amount) in sorted(monthly.items()):
        print(f"{month}\t{count}\t{amount:,}")

    print("\n部署別・月別の金額合計")
    months = sorted(monthly)
    print("年月\t" + "\t".join(master.SALES_DEPARTMENT_NAMES))
    by_dept_month: dict[tuple[str, str], int] = defaultdict(int)
    for r in rows:
        by_dept_month[(r.department, f"{r.date:%Y/%m}")] += r.amount
    for month in months:
        print(f"{month}\t" + "\t".join(f"{by_dept_month[(d, month)]:,}" for d in master.SALES_DEPARTMENT_NAMES))


def main() -> int:
    rows = build_rows()
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    write_workbook(rows, OUTPUT)
    scrub_personal_info(OUTPUT)
    verify_no_personal_info(OUTPUT)
    print(f"saved {OUTPUT} (personal info removed)\n")
    print_summary(rows)
    return 0


if __name__ == "__main__":
    sys.exit(main())
