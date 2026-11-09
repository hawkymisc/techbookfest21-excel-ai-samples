"""サンプルデータ共通のマスタ（docs/sample-data.md §1〜§3・§5・§6）。

架空の会社「株式会社ミナセ流体工業」の組織・商品・得意先・販売代理店。
売上明細（#11）をはじめ、#12〜#18 のサンプル生成スクリプトはここから読む。
得意先は乱数で組み立てるが、シードを固定しているので毎回同じ一覧になる。
"""
import datetime as dt
import random
from collections.abc import Iterable
from types import MappingProxyType
from typing import NamedTuple

COMPANY_NAME = "株式会社ミナセ流体工業"


# --- §2 組織 ---
# 組織は ORG_UNITS を唯一の正とし、ORG_CODES・DIVISIONS・SalesDepartment.budget_name はここから導出する。

class OrgUnit(NamedTuple):
    """組織の末端の単位。code は 8.4 で付番したときの正解例（§2.2。4桁：1桁目＝本部、2桁目＝部・工場、3〜4桁目＝課・拠点）。"""
    code: str
    division: str            # 本部
    department: str | None   # 部・工場（本部の直下にある課は None）
    section: str | None      # 課・拠点（部の直下に課がない場合は None）

    @property
    def name(self) -> str:
        """資料での書き方（予算表・付番表）。「東日本営業部 大阪支店」のように部と課を空白でつなぐ。"""
        return " ".join(p for p in (self.department, self.section) if p)


ORG_UNITS = (
    OrgUnit("1110", "管理本部", "総務部", "総務課"),
    OrgUnit("1120", "管理本部", "総務部", "人事課"),
    OrgUnit("1200", "管理本部", "経理部", None),
    OrgUnit("1300", "管理本部", "経営企画部", None),
    OrgUnit("1400", "管理本部", None, "情報システム課"),
    OrgUnit("2110", "営業本部", "東日本営業部", "本社営業課"),
    OrgUnit("2120", "営業本部", "東日本営業部", "北関東営業所"),
    OrgUnit("2130", "営業本部", "東日本営業部", "仙台営業所"),
    OrgUnit("2210", "営業本部", "西日本営業部", "大阪支店"),
    OrgUnit("2220", "営業本部", "西日本営業部", "名古屋支店"),
    OrgUnit("2230", "営業本部", "西日本営業部", "広島営業所"),
    OrgUnit("2240", "営業本部", "西日本営業部", "福岡営業所"),
    OrgUnit("2300", "営業本部", "海外営業部", None),
    OrgUnit("3110", "生産本部", "第一工場", "製造一課"),
    OrgUnit("3120", "生産本部", "第一工場", "製造二課"),
    OrgUnit("3130", "生産本部", "第一工場", "製造三課"),
    OrgUnit("3210", "生産本部", "第二工場", "製造課"),
    OrgUnit("3300", "生産本部", "生産管理部", None),
    OrgUnit("3400", "生産本部", "品質保証部", None),
    OrgUnit("3500", "生産本部", "資材購買部", None),
    OrgUnit("4100", "技術本部", "設計部", None),
    OrgUnit("4200", "技術本部", "開発部", None),
    OrgUnit("5100", "サービス部", "サービス部", "保守サービス課"),
)

# 付番の正解例（コード → 資料での書き方）。書き換えられない辞書
ORG_CODES = MappingProxyType({u.code: u.name for u in ORG_UNITS})

# 本部ごとの人数（§2）
DIVISION_HEADCOUNT = MappingProxyType({
    "管理本部": 45, "営業本部": 70, "生産本部": 235, "技術本部": 45, "サービス部": 25,
})


class Division(NamedTuple):
    name: str
    units: tuple[OrgUnit, ...]
    headcount: int


DIVISIONS = tuple(
    Division(name, tuple(u for u in ORG_UNITS if u.division == name), headcount)
    for name, headcount in DIVISION_HEADCOUNT.items()
)


class SalesDepartment(NamedTuple):
    """売上を立てる部署（§2.1）。資料ごとの名前の揺れと、§2.2 の付番の正解例を持つ。"""
    name: str                             # 販売管理システム上の名前（売上明細）
    agency_report_names: tuple[str, ...]  # 代理店の報告での書き方（#15）
    answer_code: str                      # 8.4 で付番したときの正解例（§2.2）。会社にはコードがない

    @property
    def budget_name(self) -> str:
        """予算表での書き方（#14）。付番表の名前と同じ"""
        return ORG_CODES[self.answer_code]


SALES_DEPARTMENTS = (
    SalesDepartment("本社営業課", ("東京", "本社"), "2110"),
    SalesDepartment("北関東営業所", ("北関東",), "2120"),
    SalesDepartment("仙台営業所", ("仙台", "東北"), "2130"),
    SalesDepartment("大阪支店", ("大阪",), "2210"),
    SalesDepartment("名古屋支店", ("名古屋", "中部"), "2220"),
    SalesDepartment("広島営業所", ("広島",), "2230"),
    SalesDepartment("福岡営業所", ("福岡", "九州"), "2240"),
    SalesDepartment("海外営業部", ("海外",), "2300"),
)
SALES_DEPARTMENT_NAMES = tuple(d.name for d in SALES_DEPARTMENTS)


# --- §3 商品 ---

class Product(NamedTuple):
    code: str
    name: str
    category: str
    price: int  # 単価（円・税抜）。2025年10月の改定前


PRODUCTS = (
    Product("P-MP100", "小型渦巻ポンプ MP-100", "ポンプ", 98_000),
    Product("P-MP200", "小型渦巻ポンプ MP-200", "ポンプ", 148_000),
    Product("P-SP050", "自吸式ポンプ SP-50", "ポンプ", 128_000),
    Product("P-DP010", "定量ポンプ DP-10", "ポンプ", 215_000),
    Product("V-EV025", "電動ボールバルブ EV-25", "バルブ", 46_000),
    Product("V-EV050", "電動ボールバルブ EV-50", "バルブ", 72_000),
    Product("V-AV025", "空気作動バルブ AV-25", "バルブ", 38_000),
    Product("V-CV040", "逆止弁 CV-40", "バルブ", 12_800),
    Product("R-MS100", "メカニカルシール（MP用）", "交換部品", 8_600),
    Product("R-IP100", "インペラ（MP用）", "交換部品", 14_500),
    Product("R-PK001", "パッキンセット", "交換部品", 2_400),
    Product("R-MU200", "モーターユニット（MP-200用）", "交換部品", 36_000),
    Product("S-INSP1", "定期点検", "保守サービス", 45_000),
    Product("S-OVHL1", "オーバーホール", "保守サービス", 120_000),
    Product("S-EMRG1", "緊急出張対応", "保守サービス", 28_000),
)

# 2025年10月の単価改定：ポンプとバルブを3%値上げ（どの単価も1円未満の端数は出ない）
PRICE_REVISION_DATE = dt.date(2025, 10, 1)
PRICE_REVISION_CATEGORIES = frozenset({"ポンプ", "バルブ"})
PRICE_REVISION_PERCENT = 3


def _revised_price(product: Product) -> int:
    revised = product.price * (100 + PRICE_REVISION_PERCENT)
    if revised % 100:
        raise ValueError(f"{product.code}: 改定後の単価に1円未満の端数が出る（{product.price}円）")
    return revised // 100


def check_price_revision(products: Iterable[Product]) -> None:
    """改定対象の商品の改定後単価に端数が出ないことを確かめる。端数があれば ValueError。"""
    for p in products:
        if p.category in PRICE_REVISION_CATEGORIES:
            _revised_price(p)


check_price_revision(PRODUCTS)  # import の時点で確かめる（fail-fast）


def unit_price(product: Product, on: dt.date) -> int:
    """その日に適用される単価。"""
    if on >= PRICE_REVISION_DATE and product.category in PRICE_REVISION_CATEGORIES:
        return _revised_price(product)
    return product.price


# --- §5 得意先 ---

class Customer(NamedTuple):
    name: str
    industry: str
    department: str  # 担当の営業部署（販売管理システム上の名前）
    stem: str        # 社名の造語部分


CUSTOMER_SEED = 20261109
# 社名の造語部分（語幹）は「1字目×2字目」の組み合わせ（144通り）から作る
STEM_FIRST = "楠葵霞椿柚菫凪鴇萌楓槙澪"
STEM_SECOND = "嶺汐澄芽穂環陽浦原里杜波"

# 業種と、社名の後ろにつける業種語
INDUSTRY_WORDS = MappingProxyType({
    "食品": ("食品", "製菓", "乳業", "醸造", "フーズ"),
    "化学": ("化学", "化成", "ケミカル"),
    "水処理": ("水処理", "環境エンジニアリング", "アクアテック"),
    "ビル設備": ("設備", "ビルサービス", "空調"),
    "機械製造": ("製作所", "工業", "精機"),
    "製紙・繊維": ("製紙", "染工"),
    "商社": ("貿易", "商事", "トレーディング"),
})
MUNICIPAL_INDUSTRY = "上下水道（自治体）"
MUNICIPAL_FORMS = ("{stem}市上下水道局", "{stem}町水道課")
DOMESTIC_INDUSTRY_WEIGHTS = MappingProxyType({
    "食品": 25, "化学": 20, "水処理": 15, "ビル設備": 15, "機械製造": 15, "製紙・繊維": 5, MUNICIPAL_INDUSTRY: 5,
})
# 部署ごとの得意先の数（部署の規模に合わせる）。海外営業部は輸出を扱う商社が得意先
CUSTOMERS_PER_DEPARTMENT = MappingProxyType({
    "本社営業課": 24, "北関東営業所": 14, "仙台営業所": 12, "大阪支店": 22,
    "名古屋支店": 16, "広島営業所": 12, "福岡営業所": 12, "海外営業部": 8,
})
# 使わない語幹。実在の有名企業と同じ社名になる語幹が見つかったらここに足す。
# 足しても、その得意先の語幹が予備の語幹に入れ替わるだけで、ほかの得意先と売上明細は変わらない（乱数の消費が変わらないため）。
# 2026-11-09 に生成した120社を目で確認した結果：該当なし（「槙原」「柚原」など名字と同じ形の語幹はあるが、有名企業ではない）
EXCLUDED_CUSTOMER_STEMS: frozenset[str] = frozenset()


def _company_name(stem: str, word: str, form: str) -> str:
    if form == "front":
        return f"株式会社{stem}{word}"
    if form == "back":
        return f"{stem}{word}株式会社"
    return f"有限会社{stem}{word}"


def build_customers(seed: int = CUSTOMER_SEED,
                    excluded_stems: frozenset[str] = EXCLUDED_CUSTOMER_STEMS) -> tuple[Customer, ...]:
    """得意先の一覧を作る。同じシードなら同じ一覧になる。

    語幹をシャッフルし、先頭から得意先の数だけを使う。残りは予備の語幹で、外した語幹の代わりに先頭から順に使う。
    業種・業種語・法人格の決め方は語幹と無関係に乱数を消費するので、語幹を外してもほかの得意先は変わらない。
    """
    if list(CUSTOMERS_PER_DEPARTMENT) != list(SALES_DEPARTMENT_NAMES):
        raise ValueError("CUSTOMERS_PER_DEPARTMENT の部署が SALES_DEPARTMENTS と一致しない")
    rng = random.Random(seed)
    stems = [a + b for a in STEM_FIRST for b in STEM_SECOND]
    rng.shuffle(stems)
    total = sum(CUSTOMERS_PER_DEPARTMENT.values())
    if total > len(stems):
        raise ValueError(f"語幹が足りない（得意先 {total} 社に対して語幹 {len(stems)} 個）")
    primary = stems[:total]
    reserve = iter(s for s in stems[total:] if s not in excluded_stems)
    industries = list(DOMESTIC_INDUSTRY_WEIGHTS)
    weights = list(DOMESTIC_INDUSTRY_WEIGHTS.values())
    customers: list[Customer] = []
    slots = (dept for dept, count in CUSTOMERS_PER_DEPARTMENT.items() for _ in range(count))
    for dept, stem in zip(slots, primary):
        if stem in excluded_stems:
            stem = next(reserve, None)
            if stem is None:
                raise ValueError(f"外した語幹の代わりに使う予備の語幹が足りない（除外 {len(excluded_stems)} 個）")
        industry = "商社" if dept == "海外営業部" else rng.choices(industries, weights=weights)[0]
        if industry == MUNICIPAL_INDUSTRY:
            name = rng.choice(MUNICIPAL_FORMS).format(stem=stem)
        else:
            word = rng.choice(INDUSTRY_WORDS[industry])
            form = rng.choices(("front", "back", "yugen"), weights=(55, 40, 5))[0]
            name = _company_name(stem, word, form)
        customers.append(Customer(name, industry, dept, stem))
    return tuple(customers)


CUSTOMERS = build_customers()


# --- §6 販売代理店 ---

class Agency(NamedTuple):
    name: str
    department: str  # 担当の営業部署（販売管理システム上の名前）


AGENCIES = (
    Agency("株式会社茜舟商会", "本社営業課"),
    Agency("翠庵機材株式会社", "本社営業課"),
    Agency("株式会社琥苑テクノ", "本社営業課"),
    Agency("栞門産業株式会社", "本社営業課"),
    Agency("株式会社晶舟ポンプ商会", "北関東営業所"),
    Agency("紬苑機工株式会社", "北関東営業所"),
    Agency("株式会社雫庵設備", "仙台営業所"),
    Agency("灯門商事株式会社", "仙台営業所"),
    Agency("株式会社縁舟商店", "大阪支店"),
    Agency("茜苑機材株式会社", "大阪支店"),
    Agency("株式会社翠門産業", "大阪支店"),
    Agency("琥舟テクノサービス株式会社", "大阪支店"),
    Agency("株式会社栞苑商会", "名古屋支店"),
    Agency("晶庵機工株式会社", "名古屋支店"),
    Agency("株式会社紬門設備", "名古屋支店"),
    Agency("雫舟産業株式会社", "広島営業所"),
    Agency("株式会社灯苑商会", "広島営業所"),
    Agency("縁庵機材株式会社", "福岡営業所"),
    Agency("株式会社茜門テクノ", "福岡営業所"),
    Agency("翠舟インターナショナル株式会社", "海外営業部"),
)
