"""Core logic for boq-diff: load, match and compare two BOQ versions.

A bill of quantities (BOQ, 工程量清单) is loaded from CSV or XLSX, rows are
matched by item code (清单编码) -- falling back to name + unit when no code is
present -- and the two versions are compared quantity by quantity and price
by price.  Money math uses :class:`decimal.Decimal` so totals do not drift.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

# Header aliases (Chinese names used by common Chinese cost software exports,
# plus English equivalents).  Headers are normalised by stripping whitespace
# and lowering case before lookup.
_HEADER_ALIASES = {
    "code": ("清单编码", "项目编码", "编码", "code", "item_code"),
    "name": ("项目名称", "清单名称", "名称", "name", "item_name"),
    "unit": ("单位", "unit"),
    "quantity": ("工程量", "数量", "quantity", "qty"),
    "unit_price": ("综合单价", "单价", "unit_price", "price"),
    "amount": ("合价", "金额", "总价", "amount", "total"),
}

_ALIAS_TO_FIELD = {
    alias: field_name
    for field_name, aliases in _HEADER_ALIASES.items()
    for alias in aliases
}

_CENT = Decimal("0.01")


def _money(value: Decimal) -> Decimal:
    return value.quantize(_CENT)


def _parse_number(raw, field_name: str, row_no: int) -> Decimal:
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return Decimal("0")
    if isinstance(raw, (int, float, Decimal)):
        return Decimal(str(raw))
    text = str(raw).replace(",", "").replace("，", "").strip()
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"第 {row_no} 行的{field_name}不是数字: {raw!r}") from exc


@dataclass(frozen=True)
class Item:
    """One BOQ line item."""

    code: str
    name: str
    unit: str
    quantity: Decimal
    unit_price: Decimal
    amount: Decimal

    @property
    def key(self) -> str:
        """Matching key: item code when present, else name + unit."""
        if self.code:
            return self.code
        return f"{self.name}|{self.unit}"


@dataclass(frozen=True)
class Change:
    """A line item present in both versions whose numbers moved."""

    old: Item
    new: Item

    @property
    def key(self) -> str:
        return self.old.key

    @property
    def quantity_delta(self) -> Decimal:
        return self.new.quantity - self.old.quantity

    @property
    def unit_price_delta(self) -> Decimal:
        return self.new.unit_price - self.old.unit_price

    @property
    def amount_delta(self) -> Decimal:
        return self.new.amount - self.old.amount


@dataclass
class DiffResult:
    added: list[Item] = field(default_factory=list)
    removed: list[Item] = field(default_factory=list)
    changed: list[Change] = field(default_factory=list)
    unchanged_count: int = 0

    @property
    def has_changes(self) -> bool:
        return bool(self.added or self.removed or self.changed)


def _rows_from_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        return list(csv.reader(fh))


def _rows_from_xlsx(path: Path) -> list[list]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise RuntimeError(
            "读取 .xlsx 需要 openpyxl，请先安装: pip install openpyxl"
        ) from exc
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    return [list(row) for row in ws.iter_rows(values_only=True)]


def _build_item(cells: dict, row_no: int) -> Item:
    code = str(cells.get("code") or "").strip()
    name = str(cells.get("name") or "").strip()
    unit = str(cells.get("unit") or "").strip()
    if not code and not name:
        raise ValueError(f"第 {row_no} 行缺少清单编码和项目名称，无法识别")
    quantity = _parse_number(cells.get("quantity"), "工程量", row_no)
    unit_price = _parse_number(cells.get("unit_price"), "综合单价", row_no)
    raw_amount = cells.get("amount")
    if raw_amount is None or (isinstance(raw_amount, str) and raw_amount.strip() == ""):
        amount = _money(quantity * unit_price)
    else:
        amount = _money(_parse_number(raw_amount, "合价", row_no))
    return Item(
        code=code,
        name=name,
        unit=unit,
        quantity=quantity,
        unit_price=unit_price,
        amount=amount,
    )


def load_boq(path) -> dict[str, Item]:
    """Load a BOQ file (.csv or .xlsx) into ``{key: Item}``.

    The first row must be a header row.  Recognised columns (any language
    variant listed in the aliases): 清单编码/编码, 项目名称, 单位, 工程量,
    综合单价/单价, 合价 (optional; computed as 工程量 × 综合单价 when absent).
    """
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        rows = _rows_from_csv(path)
    elif suffix in (".xlsx", ".xlsm"):
        rows = _rows_from_xlsx(path)
    else:
        raise ValueError(f"不支持的文件类型: {suffix}（仅支持 .csv / .xlsx）")

    rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
    if not rows:
        raise ValueError("文件是空的")

    header = [str(c or "").strip().lower() for c in rows[0]]
    col_field = {}
    for idx, h in enumerate(header):
        field_name = _ALIAS_TO_FIELD.get(h)
        if field_name:
            col_field[idx] = field_name
    if not any(f in col_field.values() for f in ("code", "name")):
        raise ValueError("表头里找不到「清单编码」或「项目名称」列")
    if "quantity" not in col_field.values() or "unit_price" not in col_field.values():
        raise ValueError("表头里找不到「工程量」或「综合单价」列")

    items: dict[str, Item] = {}
    for row_no, row in enumerate(rows[1:], start=2):
        cells = {}
        for idx, field_name in col_field.items():
            if idx < len(row):
                cells[field_name] = row[idx]
        item = _build_item(cells, row_no)
        if item.key in items:
            raise ValueError(
                f"清单中有重复项（编码/名称相同）: {item.key}，请先合并后再对比"
            )
        items[item.key] = item
    return items


def total_amount(items: dict[str, Item] | list[Item]) -> Decimal:
    values = items.values() if isinstance(items, dict) else items
    return _money(sum((it.amount for it in values), Decimal("0")))


def diff_boq(old: dict[str, Item], new: dict[str, Item]) -> DiffResult:
    """Compare two loaded BOQs and return the structured difference."""
    result = DiffResult()
    for key in sorted(new.keys() - old.keys()):
        result.added.append(new[key])
    for key in sorted(old.keys() - new.keys()):
        result.removed.append(old[key])
    for key in sorted(old.keys() & new.keys()):
        o, n = old[key], new[key]
        if o.quantity != n.quantity or o.unit_price != n.unit_price or o.amount != n.amount:
            result.changed.append(Change(old=o, new=n))
        else:
            result.unchanged_count += 1
    return result
