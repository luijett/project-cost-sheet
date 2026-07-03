"""
导入引擎 — 粘贴 + Excel
========================
剪贴板文字 / Excel 文件 → 智能解析 → 费用项列表
"""

import re
from decimal import Decimal
from typing import Optional


def _m(v) -> Decimal:
    try: return Decimal(str(v)).quantize(Decimal("0.01"))
    except: return Decimal("0")


def _parse_price(text: str) -> Optional[Decimal]:
    t = text.strip()
    t = re.sub(r'[¥￥\s]', '', t)
    t = t.replace(',', '').replace('，', '')
    if re.match(r'^-?\d+\.?\d*$', t):
        try: return Decimal(t)
        except: pass
    return None


def _is_priceish(text: str) -> bool:
    t = text.strip().replace(',', '').replace('，', '').replace(' ', '')
    t = re.sub(r'^[¥￥]', '', t)
    return bool(re.match(r'^\d+\.?\d*$', t))


def _is_int(text: str) -> bool:
    """小整数 — 很可能是数量而非价格"""
    t = text.strip()
    return bool(re.match(r'^\d{1,2}$', t))


def _looks_like_section(text: str) -> bool:
    t = text.strip()
    if len(t) > 15: return False
    kw = ["费", "造型", "美术", "器材", "后期", "制作", "杂费",
          "模特", "演员", "场地", "灯光", "摄影", "服装", "道具",
          "妆发", "化妆", "发型", "交通", "餐饮", "住宿", "差旅"]
    return any(k in t for k in kw)


def _looks_like_total(text: str) -> bool:
    return any(k in text for k in ["总计","合计","含税","总价","总额","共计","总金额","汇总"])


# ═══════════════════════════════════════════
#  粘贴文字解析
# ═══════════════════════════════════════════

def parse_clipboard_text(text: str) -> list[dict]:
    """解析剪贴板文字 → [{description, unit_price, quantity, unit, total, category}]"""
    lines = text.strip().split("\n")
    items = []
    current_section = "未分类"

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 尝试 tab 分割
        if "\t" in line:
            parts = [p.strip() for p in line.split("\t") if p.strip()]
        else:
            # 空格分割 — 智能合并
            parts = _smart_split(line)

        if not parts:
            continue

        # 检测分类标题
        full = " ".join(parts)
        if _looks_like_section(full) and not any(c.isdigit() for c in full):
            current_section = full
            continue

        # 检测合计行
        if _looks_like_total(full):
            continue

        # 至少 2 列才算明细行
        if len(parts) < 2:
            continue

        # 展开合并的"数字+单位"（如 "1天" → "1", "天"）
        expanded = []
        for p in parts:
            m = re.match(r'^(\d+\.?\d*)([天次套件组人个张本台辆场小时]+)$', p)
            if m:
                expanded.append(m.group(1))
                expanded.append(m.group(2))
            else:
                expanded.append(p)
        parts = expanded

        # 从左到右：描述 → 数字列
        desc_parts = []
        num_parts = []
        in_num = False
        for p in parts:
            if not in_num and (_is_priceish(p) or _is_int(p)):
                in_num = True
            if in_num:
                num_parts.append(p)
            else:
                desc_parts.append(p)

        description = " ".join(desc_parts).strip()
        if not description:
            continue

        # 解析数字列
        prices, qtys, unit_text = [], [], ""
        for p in num_parts:
            pp = _parse_price(p)
            if pp is not None and _is_int(p) and '.' not in p:
                qtys.append(pp)
            elif pp is not None:
                prices.append(pp)
            elif len(p) <= 4 and not _is_priceish(p):
                unit_text = p
            else:
                pp2 = _parse_price(p)
                if pp2 is not None:
                    prices.append(pp2)
                elif len(p) <= 4:
                    unit_text = p

        if not prices and not qtys:
            continue

        up, qty, tot = Decimal("0"), Decimal("1"), Decimal("0")
        if len(prices) >= 2:
            ps = sorted(prices)
            up, tot = ps[0], ps[-1]
            qty = qtys[0] if qtys else Decimal("1")
        elif len(prices) == 1:
            tot = prices[0]
            qty = qtys[0] if qtys else Decimal("1")
            up = tot / qty if qty > 0 else tot
        elif qtys:
            qty = qtys[0]

        if description:
            items.append({
                "description": description,
                "unit_price": _m(up),
                "quantity": _m(qty),
                "unit": unit_text,
                "total": _m(tot if tot > 0 else up * qty),
                "category": current_section,
            })

    return items


def _smart_split(line: str) -> list[str]:
    """智能分割 — 优先多空格 → 单空格 + 合并中文"""
    # 1. 多个空格分割（表格粘贴常见）
    if '  ' in line:
        return [p.strip() for p in re.split(r'\s{2,}', line) if p.strip()]
    # 2. 单空格分割 — 把中文词组合并回去
    parts = line.split(' ')
    result = []
    buf = []
    for p in parts:
        p = p.strip()
        if not p: continue
        # 数字或带单位 → 独立列
        if _is_priceish(p) or re.match(r'^\d+[天次套件组人个张本台辆场小时]?$', p) or re.match(r'^[¥￥]', p):
            if buf:
                result.append(' '.join(buf)); buf.clear()
            result.append(p)
        elif re.match(r'^[天次套件组人个张本台辆场小时]$', p):
            if buf: result.append(' '.join(buf)); buf.clear()
            result.append(p)
        else:
            buf.append(p)
    if buf:
        result.append(' '.join(buf))
    return result


# ═══════════════════════════════════════════
#  Excel 导入
# ═══════════════════════════════════════════

def import_excel(file_path: str) -> list[dict]:
    """导入 Excel → 自动匹配列 → [{description, unit_price, quantity, unit, total, category}]"""
    try:
        import openpyxl
    except ImportError:
        raise ImportError("请安装 openpyxl: pip install openpyxl")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=1, values_only=True))

    if not rows:
        return []

    # 尝试找表头行（前 5 行）
    header_row = None
    header_idx = -1
    col_map = {}  # col_index → field_name

    for ri in range(min(5, len(rows))):
        row = rows[ri]
        if not row:
            continue
        texts = [str(c).strip() if c else "" for c in row]
        full = " ".join(texts)
        # 检查是否包含表头关键词
        header_kw = ["名称", "描述", "项目", "内容", "费用", "单价", "数量", "合计",
                     "金额", "价格", "单位", "总价", "品类", "分类"]
        matches = sum(1 for kw in header_kw if kw in full)
        if matches >= 2:
            header_row = texts
            header_idx = ri
            break

    if header_row:
        # 自动匹配列
        for ci, col_name in enumerate(header_row):
            cn = col_name.lower()
            if any(k in cn for k in ["名称", "描述", "项目", "内容", "费用", "品类"]):
                col_map[ci] = "description"
            elif any(k in cn for k in ["单价", "价格"]):
                col_map[ci] = "unit_price"
            elif any(k in cn for k in ["数量"]):
                col_map[ci] = "quantity"
            elif any(k in cn for k in ["单位"]):
                col_map[ci] = "unit"
            elif any(k in cn for k in ["合计", "总价", "金额", "小计"]):
                col_map[ci] = "total"
            elif any(k in cn for k in ["分类", "类别"]):
                col_map[ci] = "category"

    # 如果有匹配，从表头下一行开始
    start_row = header_idx + 1 if header_idx >= 0 else 0

    # 如果没找到表头，尝试默认映射（第一列=描述，倒数第二/第一列=金额）
    if not col_map:
        # 猜测：第一列为描述，数字列按位置分配
        col_map = _guess_columns(rows)

    if "description" not in col_map.values():
        # 没匹配到描述列 → 用第一列非数字列
        for r in rows[start_row:start_row+3]:
            if r:
                for ci, val in enumerate(r):
                    if val and not _is_priceish(str(val)):
                        col_map[ci] = "description"
                        break
                break

    items = []
    current_section = "未分类"

    for row in rows[start_row:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue

        # 取出各列值
        desc = ""
        up = None
        qty = None
        unit = ""
        tot = None
        cat = ""

        for ci, val in enumerate(row):
            if val is None:
                continue
            sv = str(val).strip()
            field = col_map.get(ci, "")
            if field == "description":
                desc = sv
            elif field == "unit_price":
                up = _parse_price(sv)
            elif field == "quantity":
                p = _parse_price(sv)
                qty = p if p else (Decimal(sv) if sv.isdigit() else None)
            elif field == "unit":
                unit = sv
            elif field == "total":
                tot = _parse_price(sv)
            elif field == "category":
                cat = sv

        # 如果没有明确的描述，用第一个文本列
        if not desc:
            for ci, val in enumerate(row):
                if val and ci not in col_map:
                    sv = str(val).strip()
                    if not _is_priceish(sv) and len(sv) > 1:
                        desc = sv
                        break

        # 如果没有明确的合计，用最后一个数字
        if tot is None:
            for val in reversed(list(row)):
                if val is not None:
                    p = _parse_price(str(val))
                    if p is not None and p > 0:
                        tot = p
                        break

        if not desc:
            continue

        # 检测分类行（只有描述没有数字的单列行）
        if tot is None and up is None:
            if _looks_like_section(desc):
                current_section = desc
            continue

        if tot is None and up is None:
            continue

        up = up or (tot if qty is None or qty == 0 else tot / qty)
        qty_val = qty or Decimal("1")
        tot = tot or (up * qty_val if up else Decimal("0"))

        items.append({
            "description": desc,
            "unit_price": _m(up or tot),
            "quantity": _m(qty_val),
            "unit": unit or "",
            "total": _m(tot),
            "category": cat or current_section,
        })

    wb.close()
    return items


def _guess_columns(rows) -> dict:
    """无表头时猜测列映射"""
    cmap = {}
    for row in rows[:10]:
        if not row: continue
        price_cols = []
        text_cols = []
        for ci, val in enumerate(row):
            if val is None: continue
            sv = str(val).strip()
            if _is_priceish(sv):
                price_cols.append(ci)
            elif not sv.isdigit() and len(sv) > 1:
                text_cols.append(ci)

        if text_cols:
            cmap[text_cols[0]] = "description"
        if len(price_cols) >= 2:
            cmap[price_cols[0]] = "unit_price"
            cmap[price_cols[-1]] = "total"
        elif len(price_cols) == 1:
            cmap[price_cols[0]] = "total"
        break
    return cmap
