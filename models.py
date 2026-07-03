"""project-cost-sheet — 数据模型 & 预设模板"""
from decimal import Decimal, ROUND_HALF_UP
from typing import Optional
from pydantic import BaseModel, Field

Money = Decimal

# ── 币种 ──

CURRENCIES = {
    "CNY": {"symbol": "¥", "name": "人民币"},
    "USD": {"symbol": "$", "name": "美元"},
    "EUR": {"symbol": "€", "name": "欧元"},
    "HKD": {"symbol": "HK$", "name": "港币"},
    "JPY": {"symbol": "¥", "name": "日元"},
    "GBP": {"symbol": "£", "name": "英镑"},
    "TWD": {"symbol": "NT$", "name": "新台币"},
}

TAX_RATES = [
    ("0%", Decimal("0")), ("1%", Decimal("0.01")), ("3%", Decimal("0.03")),
    ("6%", Decimal("0.06")), ("9%", Decimal("0.09")), ("13%", Decimal("0.13")),
]

# ── 单位 ──

UNIT_PRESETS = ["", "天", "小时", "次", "套", "件", "组", "人", "个", "张", "本", "台", "辆", "场", "自定义"]


def m(amount) -> Money:
    return Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def fmt_money(amount: Money, currency: str = "CNY") -> str:
    sym = CURRENCIES.get(currency, {}).get("symbol", "¥")
    if amount == 0:
        return f"{sym}0"
    return f"{sym}{amount:,.2f}"


def calc_tax(subtotal: Money, tax_rate: Money) -> Money:
    return m(subtotal * tax_rate)


def calc_total_with_tax(subtotal: Money, tax_rate: Money) -> Money:
    return m(subtotal + calc_tax(subtotal, tax_rate))


# ── 预设模板 ──

PRINT_CATEGORIES = [
    ("场地费", [
        ("影棚/外景租金", ""), ("电费/空调", ""), ("场地清理", ""),
    ]),
    ("模特费", [
        ("模特出场费", ""), ("模特差旅", ""),
    ]),
    ("妆发造型", [
        ("化妆师", ""), ("发型师", ""), ("妆发耗材", ""),
    ]),
    ("美术", [
        ("服装采购/租赁", ""), ("道具采购/制作", ""),
        ("损耗/干洗", ""), ("美术助理", ""),
    ]),
    ("摄影器材", [
        ("相机/镜头租赁", ""), ("灯光租赁", ""), ("附件/耗材", ""),
    ]),
    ("后期修图", [
        ("修图师", ""), ("修图外包", ""),
    ]),
    ("制片杂费", [
        ("交通/停车", ""), ("餐饮/茶歇", ""),
        ("水/降温物料", ""), ("打印/办公", ""), ("不可预见费", ""),
    ]),
]

VIDEO_CATEGORIES = [
    ("场地费", [
        ("影棚/实景租金", ""), ("电费/空调", ""), ("场地清理/复原", ""),
    ]),
    ("模特/演员", [
        ("模特出场费", ""), ("演员出场费", ""), ("差旅/住宿", ""),
    ]),
    ("妆发造型", [
        ("化妆师", ""), ("发型师", ""), ("妆发耗材", ""),
    ]),
    ("美术", [
        ("服装采购/租赁", ""), ("道具采购/制作", ""),
        ("置景/搭建", ""), ("损耗/干洗", ""), ("美术助理", ""),
    ]),
    ("摄影器材", [
        ("机身/镜头租赁", ""), ("跟焦/监视/图传", ""),
        ("移动组(Dolly/轨/摇臂)", ""), ("附件/耗材", ""),
    ]),
    ("灯光器材", [
        ("灯光租赁", ""), ("发电车", ""), ("灯光附件", ""),
    ]),
    ("后期制作", [
        ("剪辑", ""), ("调色", ""), ("混音/音效", ""),
        ("音乐授权", ""), ("特效/包装", ""),
    ]),
    ("制片杂费", [
        ("交通/停车/燃油", ""), ("餐饮/茶歇", ""),
        ("水/降温物料", ""), ("通讯/对讲机", ""),
        ("打印/办公", ""), ("不可预见费", ""),
    ]),
]


def merge_categories(types: list[str]) -> list:
    """合并多个类型的预设分类，用于套拍"""
    if len(types) == 1:
        return PRINT_CATEGORIES if types[0] == "print" else VIDEO_CATEGORIES

    merged = {}
    for t in types:
        cats = PRINT_CATEGORIES if t == "print" else VIDEO_CATEGORIES
        for cat_name, items in cats:
            if cat_name not in merged:
                merged[cat_name] = {}
            for desc, note in items:
                if desc not in merged[cat_name]:
                    merged[cat_name][desc] = note

    return [(name, [(d, n) for d, n in items.items()]) for name, items in merged.items()]


# ── 数据模型 ──

class Project(BaseModel):
    id: Optional[int] = None
    name: str = ""
    project_types: str = "print"
    currency: str = "CNY"
    tax_rate: Money = Field(default_factory=lambda: Decimal("0"))
    notes: str = ""
    created_at: str = ""
    updated_at: str = ""

    def type_list(self) -> list[str]:
        return [t.strip() for t in self.project_types.split(",") if t.strip()]

    def type_label(self) -> str:
        types = self.type_list()
        if len(types) == 2:
            return "📷🎬 套拍"
        return "📷 平面" if types[0] == "print" else "🎬 视频"


class BudgetCategory(BaseModel):
    id: Optional[int] = None
    project_id: int = 0
    name: str = ""
    sort_order: int = 0


class LineItem(BaseModel):
    id: Optional[int] = None
    category_id: int = 0
    description: str = ""
    unit_price: Money = Field(default_factory=lambda: Decimal("0"))
    quantity: Decimal = Field(default_factory=lambda: Decimal("1"))
    unit: str = ""
    total: Money = Field(default_factory=lambda: Decimal("0"))
    notes: str = ""
    sort_order: int = 0

    def calc_total(self) -> Money:
        return m(self.unit_price * self.quantity)
