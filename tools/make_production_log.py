"""生産実績ログ（Issue #17）のサンプルブックを作る。仕様は docs/sample-data.md §7（#17 の行）。

架空の会社「株式会社ミナセ流体工業」の第一工場・第二工場で、製造ラインごとにロット単位で記録した
2024年10月〜2026年9月の生産実績。1行が1ロット。数十万行あり、11.1「データが大きすぎると言われたら」と
第11章コラム（コンテキスト上限）の題材にする。丸ごと AI に渡すと大きすぎるが、範囲を絞る（工場・ライン・月）、
先に集計する（不良率を工場別・月別・ライン別に）、サンプル行だけ見せる、のどれでも扱える列構成にしてある。

乱数は固定シードなので、何度実行しても同じ内容になる。一時フォルダに保存して作成者名・保存先パスなどの
個人情報を取り除き、残っていないことを確認してから samples/ に移す（tools/staging.py）。最後に、原稿の検算に使う集計を標準出力に出す。

    uv run python tools/make_production_log.py

集計すると見つかるように仕込んだ傾向（tests/test_make_production_log.py で確かめる）:
- 第二工場は第一工場より不良率が高い
- 夏（7〜8月）は不良率が上がり、第二工場のほうが上がり幅が大きい
- ライン 2V-3（第二工場のバルブライン）だけ、2026年4月から不良率が月を追って上がり、増えた不良の大半が「漏れ」

不良区分は、そのロットで最も多かった不良の区分を1つだけ入れる（不良数が0なら空欄）。
生産量は、売上明細（#11。約7,000行）の販売量とは突き合わせられる規模になっていない。11.1 の題材として行数を優先した。

傾向の係数やロットの作り方を変えたら、`uv run pytest tests/test_make_production_log.py` で傾向のテストが通ることを
確かめ、tests/test_production_golden.py のゴールデン値を出力に合わせて更新する。
"""
import datetime as dt
import math
import random
import sys
import time
from collections import defaultdict
from pathlib import Path
from types import MappingProxyType
from typing import NamedTuple

import xlwings as xw

import sample_master as master
import staging
from make_sales import END, START, business_days
from xlsx_meta import scrub_personal_info, verify_no_personal_info

OUTPUT = Path(__file__).resolve().parent.parent / "samples" / "生産実績ログ_2024年10月-2026年9月.xlsx"
SHEET_NAME = "生産実績"
SEED = 20241001
# 期間（START〜END）は売上明細と同じ（make_sales から import している）

HEADER = [
    "日付", "工場", "ライン", "品目コード", "ロット番号", "計画数", "実績数",
    "良品数", "不良数", "不良区分", "稼働時間（分）", "作業班",
]
# xlwings で一度に書く行数。大きすぎると COM の呼び出しが失敗しやすく、小さすぎると遅い
WRITE_CHUNK_ROWS = 20_000


# --- ライン ---

class Line(NamedTuple):
    name: str                  # 例：1P-1（工場の番号＋区分の頭文字＋連番）
    factory: str
    section: str               # 担当の課（§2 の組織）。ブックには出さない
    category: str              # 商品区分
    products: tuple[str, ...]  # 作る品目の商品コード


_F1, _F2 = "第一工場", "第二工場"
LINES = (
    Line("1P-1", _F1, "製造一課", "ポンプ", ("P-MP100",)),
    Line("1P-2", _F1, "製造一課", "ポンプ", ("P-MP100", "P-MP200")),
    Line("1P-3", _F1, "製造一課", "ポンプ", ("P-MP200",)),
    Line("1P-4", _F1, "製造一課", "ポンプ", ("P-SP050", "P-DP010")),
    Line("1V-1", _F1, "製造二課", "バルブ", ("V-EV025",)),
    Line("1V-2", _F1, "製造二課", "バルブ", ("V-EV025", "V-EV050")),
    Line("1V-3", _F1, "製造二課", "バルブ", ("V-EV050",)),
    Line("1V-4", _F1, "製造二課", "バルブ", ("V-AV025",)),
    Line("1R-1", _F1, "製造三課", "交換部品", ("R-MS100",)),
    Line("1R-2", _F1, "製造三課", "交換部品", ("R-IP100",)),
    Line("1R-3", _F1, "製造三課", "交換部品", ("R-PK001",)),
    Line("1R-4", _F1, "製造三課", "交換部品", ("R-MU200",)),
    Line("2P-1", _F2, "製造課", "ポンプ", ("P-MP100", "P-SP050")),
    Line("2P-2", _F2, "製造課", "ポンプ", ("P-MP200", "P-DP010")),
    Line("2V-1", _F2, "製造課", "バルブ", ("V-AV025",)),
    Line("2V-2", _F2, "製造課", "バルブ", ("V-CV040",)),
    Line("2V-3", _F2, "製造課", "バルブ", ("V-EV025", "V-CV040")),
    Line("2R-1", _F2, "製造課", "交換部品", ("R-MS100", "R-PK001")),
    Line("2R-2", _F2, "製造課", "交換部品", ("R-IP100",)),
    Line("2R-3", _F2, "製造課", "交換部品", ("R-PK001", "R-MU200")),
)
LINE_BY_NAME = MappingProxyType({line.name: line for line in LINES})
LINE_ORDER = MappingProxyType({line.name: i for i, line in enumerate(LINES)})


def check_lines(lines: tuple[Line, ...]) -> None:
    """ラインの品目が商品マスタにあり、区分がラインの区分と一致することを確かめる。違えば ValueError。"""
    by_code = {p.code: p for p in master.PRODUCTS}
    for line in lines:
        for code in line.products:
            if code not in by_code:
                raise ValueError(f"{line.name}: 商品マスタにない品目 {code}")
            if by_code[code].category != line.category:
                raise ValueError(f"{line.name}: {code} の区分が {by_code[code].category}（ラインは {line.category}）")


check_lines(LINES)  # 起動時に確かめる


# --- 作業班と稼働時間 ---

CREWS = ("A班", "B班", "C班")
# 1日2直（日勤・夕勤）。各直の稼働できる時間（分）。休憩を除く
SHIFT_MINUTES = 420


def shift_crews(day: dt.date) -> tuple[str, str]:
    """その日の日勤・夕勤の作業班。3班が週ごとに入れ替わる（両工場とも同じ並び）。"""
    week = (day - (START - dt.timedelta(days=START.weekday()))).days // 7
    return CREWS[week % 3], CREWS[(week + 1) % 3]


# --- ロットの大きさと1個あたりの作業時間 ---

LOT_SIZES = MappingProxyType({
    "ポンプ": (10, 15, 20, 25, 30),
    "バルブ": (20, 30, 40, 50, 60),
    "交換部品": (100, 150, 200, 250, 300),
})
CYCLE_MINUTES = MappingProxyType({"ポンプ": 1.1, "バルブ": 0.45, "交換部品": 0.09})


# --- 不良率の傾向 ---

BASE_DEFECT_RATE = MappingProxyType({"ポンプ": 0.020, "バルブ": 0.015, "交換部品": 0.008})
FACTORY_FACTOR = MappingProxyType({_F1: 1.0, _F2: 1.35})
# 月ごとの係数。夏に上がり、第二工場（三重）のほうが上がり幅が大きい
SEASON_FACTOR = MappingProxyType({
    _F1: MappingProxyType({7: 1.15, 8: 1.15}),
    _F2: MappingProxyType({6: 1.15, 7: 1.5, 8: 1.5, 9: 1.15}),
})
# 2V-3 だけ、2026年4月から不良率が月を追って上がる（増えた分は主に「漏れ」）
ANOMALY_LINE = "2V-3"
ANOMALY_START = dt.date(2026, 4, 1)
ANOMALY_STEP = 0.45   # 2026年4月で 1.45 倍、9月で 3.7 倍
ANOMALY_TYPE = "漏れ"
# ロットごとのばらつき（不良率に掛ける係数の標準偏差）
LOT_NOISE = 0.3

DEFECT_TYPES = MappingProxyType({
    "ポンプ": ("漏れ", "異音・振動", "寸法不良", "外観不良"),
    "バルブ": ("漏れ", "作動不良", "寸法不良", "外観不良"),
    "交換部品": ("寸法不良", "外観不良", "材料不良"),
})
DEFECT_TYPE_WEIGHTS = MappingProxyType({
    "ポンプ": (25, 20, 30, 25),
    "バルブ": (25, 20, 30, 25),
    "交換部品": (45, 35, 20),
})


class ProductionRow(NamedTuple):
    date: dt.date
    factory: str
    line: str
    product_code: str
    lot_no: str
    planned: int
    produced: int
    good: int
    defects: int
    defect_type: str | None
    minutes: int
    crew: str

    def as_list(self) -> list:
        return list(self)


def _anomaly_factor(line: str, day: dt.date) -> float:
    if line != ANOMALY_LINE or day < ANOMALY_START:
        return 1.0
    months = (day.year - ANOMALY_START.year) * 12 + day.month - ANOMALY_START.month + 1
    return 1.0 + ANOMALY_STEP * months


def _defect_type(rng: random.Random, category: str, anomaly: float) -> str:
    types = DEFECT_TYPES[category]
    weights = list(DEFECT_TYPE_WEIGHTS[category])
    if anomaly > 1.0:
        # 増えた分の不良を「漏れ」に寄せる。もとの不良（重み合計）に対し、増えた分（anomaly - 1 倍）を漏れに足す
        weights[types.index(ANOMALY_TYPE)] += sum(weights) * (anomaly - 1.0)
    return rng.choices(types, weights=weights)[0]


def _line_day(rng: random.Random, line: Line, day: dt.date) -> list[ProductionRow]:
    """1ライン・1日分のロット。日勤・夕勤それぞれ、稼働時間を使い切るまでロットを流す。"""
    base = (BASE_DEFECT_RATE[line.category] * FACTORY_FACTOR[line.factory]
            * SEASON_FACTOR[line.factory].get(day.month, 1.0))
    anomaly = _anomaly_factor(line.name, day)
    rows = []
    seq = 0
    for crew in shift_crews(day):
        used = 0
        while used < SHIFT_MINUTES:
            seq += 1
            planned = rng.choice(LOT_SIZES[line.category])
            # 実績は計画の 80%〜105% に収める
            produced = min(math.floor(planned * 1.05), max(math.ceil(planned * 0.8),
                                                           round(planned * rng.gauss(0.985, 0.02))))
            minutes = max(1, round(produced * CYCLE_MINUTES[line.category] * rng.gauss(1.0, 0.08)))
            rate = min(0.5, base * anomaly * max(0.1, rng.gauss(1.0, LOT_NOISE)))
            defects = rng.binomialvariate(produced, rate)
            defect_type = _defect_type(rng, line.category, anomaly) if defects else None
            rows.append(ProductionRow(
                day, line.factory, line.name, rng.choice(line.products),
                f"{line.name.replace('-', '')}-{day:%y%m%d}-{seq:02d}",
                planned, produced, produced - defects, defects, defect_type, minutes, crew,
            ))
            used += minutes
    return rows


def build_rows(seed: int = SEED) -> list[ProductionRow]:
    """生産実績の行を作る。日付・ライン・ロット番号の順に並ぶ。同じシードなら同じ内容になる。"""
    rng = random.Random(seed)
    rows: list[ProductionRow] = []
    for day in business_days(START, END):
        for line in LINES:
            rows.extend(_line_day(rng, line, day))
    return rows


def write_workbook(rows: list[ProductionRow], path: Path) -> None:
    """ブックに書く。数十万行を一度に渡すと COM が重いので、WRITE_CHUNK_ROWS 行ずつ書く。"""
    last = len(rows) + 1
    with xw.App(visible=False, add_book=False) as app:
        app.screen_updating = False
        book = app.books.add()
        while len(book.sheets) > 1:  # 「新しいブックのシート数」設定が2以上でも1枚にする
            book.sheets[-1].delete()
        sheet = book.sheets[0]
        sheet.name = SHEET_NAME
        sheet.range("A1").value = [HEADER]
        for start in range(0, len(rows), WRITE_CHUNK_ROWS):
            chunk = [r.as_list() for r in rows[start:start + WRITE_CHUNK_ROWS]]
            sheet.range(f"A{start + 2}").value = chunk
        header = sheet.range("A1:L1")
        header.font.bold = True
        header.color = (221, 235, 247)
        sheet.range(f"A2:A{last}").number_format = "yyyy/mm/dd"
        sheet.range(f"F2:I{last}").number_format = "#,##0"
        sheet.range(f"K2:K{last}").number_format = "#,##0"
        sheet.range("A1:L1000").columns.autofit()  # 全行で autofit すると遅いので先頭だけで幅を決める
        book.api.RemovePersonalInformation = True
        book.save(str(path))
        book.close()


def _rate(defects: int, produced: int) -> str:
    return f"{defects / produced:.2%}"


def print_summary(rows: list[ProductionRow]) -> None:
    produced = sum(r.produced for r in rows)
    defects = sum(r.defects for r in rows)
    print(f"総行数 {len(rows):,}　実績数 {produced:,}　良品数 {produced - defects:,}　不良数 {defects:,}"
          f"（不良率 {_rate(defects, produced)}）")

    def table(title: str, key) -> None:
        agg: dict[object, list[int]] = defaultdict(lambda: [0, 0, 0])
        for r in rows:
            a = agg[key(r)]
            a[0] += 1
            a[1] += r.produced
            a[2] += r.defects
        print(f"\n{title}\n区分\t行数\t実績数\t不良数\t不良率")
        for k, (n, p, d) in sorted(agg.items()):
            print(f"{k}\t{n:,}\t{p:,}\t{d:,}\t{_rate(d, p)}")

    table("工場別", lambda r: r.factory)
    table("工場別・月別", lambda r: f"{r.factory} {r.date:%Y/%m}")
    table("ライン別（2026年4月〜9月）", lambda r: r.line if r.date >= ANOMALY_START else "（2026年3月以前）")
    table("ライン 2V-3 の月別", lambda r: f"{r.date:%Y/%m}" if r.line == ANOMALY_LINE else "（ほかのライン）")
    table("ライン 2V-3 の不良区分（2026年4月〜9月）",
          lambda r: (r.defect_type or "（不良なし）") if r.line == ANOMALY_LINE and r.date >= ANOMALY_START
          else "（対象外）")
    table("作業班別", lambda r: r.crew)


def main() -> int:
    t0 = time.perf_counter()
    rows = build_rows()
    t1 = time.perf_counter()
    with staging.staging_dir() as stage:
        staged = stage / OUTPUT.name
        write_workbook(rows, staged)
        t2 = time.perf_counter()
        scrub_personal_info(staged)
        verify_no_personal_info(staged)
        t3 = time.perf_counter()
        staging.publish(staged, OUTPUT)
    print(f"saved {OUTPUT} (personal info removed)")
    print(f"生成 {t1 - t0:.1f}秒　Excel書き込み・保存 {t2 - t1:.1f}秒　個人情報の除去・検査 {t3 - t2:.1f}秒"
          f"　ファイルサイズ {OUTPUT.stat().st_size:,} バイト\n")
    print_summary(rows)
    return 0


if __name__ == "__main__":
    sys.exit(main())
