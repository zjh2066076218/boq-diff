"""Core logic for boq-diff: load, match and compare two BOQ versions.

Loads bills of quantities (工程量清单) exported by common Chinese cost
software (广联达 / 鲁班 / 斯维尔 / 品茗 / 新点 ...) as well as plain
hand-made tables.  Real-world exports are messy, so the loader:

* hunts for the header row anywhere in the first rows (title rows above it
  are ignored) and understands two-row headers with a merged parent cell
  (``金额`` spanning ``综合单价`` / ``合价``), as in GB 50500 templates;
* expands merged cells, scans **all** worksheets of a workbook and uses
  every sheet that looks like a BOQ;
* skips chapter / section rows and 小计 / 合计 subtotal rows (they carry no
  quantity and unit price);
* accepts GBK-encoded CSV files and legacy ``.xls`` workbooks (the latter
  needs the optional ``xlrd`` package).

Rows are matched by item code (清单编码 / 项目编码), falling back to
name + unit.  Money math uses :class:`decimal.Decimal` so totals do not
drift.
"""

from __future__ import annotations

import csv
import io
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

# ---------------------------------------------------------------------------
# Header handling
# ---------------------------------------------------------------------------

# Field -> header aliases.  Headers are normalised (whitespace removed,
# parenthesised notes like （元） dropped, lower-cased) before lookup, so
# "综合单价（元）" and "综合单价" are the same alias.
_HEADER_ALIASES = {
    "code": ("清单编码", "项目编码", "分项编码", "编码", "code", "item_code"),
    "name": ("项目名称", "清单名称", "清单项目名称", "分项工程名称", "名称",
             "name", "item_name"),
    "unit": ("单位", "计量单位", "unit"),
    "quantity": ("工程量", "工程数量", "清单工程量", "数量", "quantity", "qty"),
    "unit_price": ("综合单价", "清单单价", "单价", "unit_price", "price"),
    "amount": ("合价", "金额", "小计", "amount", "total"),
}


def _norm_header(text) -> str:
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = re.sub(r"[\s　]+", "", s)
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    return s


_ALIAS_TO_FIELD = {
    _norm_header(alias): field_name
    for field_name, aliases in _HEADER_ALIASES.items()
    for alias in aliases
}

_CENT = Decimal("0.01")
_HEADER_SEARCH_ROWS = 30


def _match_columns(rows: list, idx: int, span: int) -> dict[int, str]:
    """Map column index -> field for a header starting at row ``idx``."""
    width = max((len(r) for r in rows[idx:idx + span]), default=0)
    mapping: dict[int, str] = {}
    for col in range(width):
        parts = []
        for r in rows[idx:idx + span]:
            parts.append(_norm_header(r[col]) if col < len(r) else "")
        # In a two-row header the lower row is the leaf label (综合单价
        # under a merged 金额 parent), so try it before the parent label.
        candidates = list(reversed([p for p in parts if p])) + ["".join(parts)]
        for cand in candidates:
            field_name = _ALIAS_TO_FIELD.get(cand)
            if field_name and field_name not in mapping.values():
                mapping[col] = field_name
                break
    return mapping


def _find_header(rows: list) -> tuple[int, int, dict[int, str]] | None:
    """Find (row_index, span, column mapping) of the BOQ header, or None."""
    best = None
    for idx in range(min(len(rows), _HEADER_SEARCH_ROWS)):
        for span in (1, 2):
            if idx + span > len(rows):
                continue
            mapping = _match_columns(rows, idx, span)
            fields = set(mapping.values())
            if not ({"code", "name"} & fields):
                continue
            if not ({"quantity", "unit_price"} <= fields):
                continue
            score = len(fields)
            if best is None or score > best[3]:
                best = (idx, span, mapping, score)
        if best is not None:
            # Earliest plausible header row wins; don't scan past it.
            break
    if best is None:
        return None
    idx, span, mapping, _ = best
    return idx, span, mapping


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

def _money(value: Decimal) -> Decimal:
    return value.quantize(_CENT)


def _is_blank(raw) -> bool:
    return raw is None or (isinstance(raw, str) and raw.strip() == "")


def _parse_number(raw, field_name: str, row_no: int) -> Decimal:
    if _is_blank(raw):
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


# ---------------------------------------------------------------------------
# Row parsing
# ---------------------------------------------------------------------------

_SUBTOTAL_MARKERS = ("小计", "合计", "本页小计", "总计")


def _parse_block(rows: list, header_idx: int, span: int,
                 mapping: dict[int, str], source: str) -> dict[str, Item]:
    items: dict[str, Item] = {}
    for offset, row in enumerate(rows[header_idx + span:]):
        row_no = header_idx + span + offset + 1  # 1-based, for messages
        cells = {f: (row[c] if c < len(row) else None) for c, f in mapping.items()}
        # Chapter / section banners, 小计 / 合计 rows and blank separators
        # carry neither a quantity nor a unit price: they are not line items.
        if _is_blank(cells.get("quantity")) and _is_blank(cells.get("unit_price")):
            continue
        code = str(cells.get("code") or "").strip()
        name = str(cells.get("name") or "").strip()
        if not code and not name:
            continue
        if any(m in name for m in _SUBTOTAL_MARKERS) and _is_blank(cells.get("code")):
            continue
        unit = str(cells.get("unit") or "").strip()
        quantity = _parse_number(cells.get("quantity"), "工程量", row_no)
        unit_price = _parse_number(cells.get("unit_price"), "综合单价", row_no)
        raw_amount = cells.get("amount")
        if _is_blank(raw_amount):
            amount = _money(quantity * unit_price)
        else:
            amount = _money(_parse_number(raw_amount, "合价", row_no))
        item = Item(code=code, name=name, unit=unit,
                    quantity=quantity, unit_price=unit_price, amount=amount)
        if item.key in items:
            raise ValueError(
                f"{source} 中有重复项（编码/名称相同）: {item.key}，请先合并后再对比"
            )
        items[item.key] = item
    return items


def _load_grid(grid: list, source: str) -> dict[str, Item] | None:
    found = _find_header(grid)
    if found is None:
        return None
    header_idx, span, mapping = found
    return _parse_block(grid, header_idx, span, mapping, source)


# ---------------------------------------------------------------------------
# File loaders
# ---------------------------------------------------------------------------

def _rows_from_csv_bytes(data: bytes) -> list[list]:
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:  # pragma: no cover - undecodable input
        raise ValueError("文件编码无法识别（已尝试 UTF-8 和 GBK）")
    sample = next((ln for ln in text.splitlines() if ln.strip()), "")
    delimiter = "\t" if sample.count("\t") > sample.count(",") else ","
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    return [list(r) for r in reader]


def _rows_from_xlsx(path: Path) -> list[tuple[str, list]]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise RuntimeError(
            "读取 .xlsx 需要 openpyxl，请先安装: pip install openpyxl"
        ) from exc
    wb = load_workbook(path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        # Expand merged cells: every cell of a merged range takes the
        # value of its top-left cell (exports merge titles and headers).
        for rng in ws.merged_cells.ranges:
            top = grid[rng.min_row - 1][rng.min_col - 1]
            for r in range(rng.min_row - 1, rng.max_row):
                for c in range(rng.min_col - 1, rng.max_col):
                    grid[r][c] = top
        sheets.append((ws.title, grid))
    return sheets


def _rows_from_xls(path: Path) -> list[tuple[str, list]]:
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError(
            "读取老式 .xls 需要 xlrd，请先安装: pip install 'boq-diff[xls]'"
        ) from exc
    wb = xlrd.open_workbook(str(path))
    sheets = []
    for ws in wb.sheets():
        grid = [list(ws.row_values(r)) for r in range(ws.nrows)]
        for rlo, rhi, clo, chi in ws.merged_cells:
            if rlo < ws.nrows and clo < ws.ncols:
                top = grid[rlo][clo]
                for r in range(rlo, min(rhi, ws.nrows)):
                    for c in range(clo, min(chi, ws.ncols)):
                        grid[r][c] = top
        sheets.append((ws.name, grid))
    return sheets


def load_boq(path) -> dict[str, Item]:
    """Load a BOQ file (.csv / .txt / .xlsx / .xls) into ``{key: Item}``.

    The header row is located automatically, so title rows, multi-row
    headers and section banners in real software exports are fine.  In a
    workbook, every sheet that contains a BOQ header contributes its items.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix in (".csv", ".txt"):
        grid = _rows_from_csv_bytes(path.read_bytes())
        items = _load_grid(grid, str(path))
        if items is None:
            raise ValueError("找不到清单表头（需要清单编码/项目名称、工程量、综合单价列）")
        return items

    if suffix in (".xlsx", ".xlsm"):
        sheets = _rows_from_xlsx(path)
    elif suffix == ".xls":
        sheets = _rows_from_xls(path)
    else:
        raise ValueError(
            f"不支持的文件类型: {suffix}（仅支持 .csv / .txt / .xlsx / .xls）"
        )

    merged: dict[str, Item] = {}
    used_sheets = []
    for sheet_name, grid in sheets:
        items = _load_grid(grid, f"{path} 工作表「{sheet_name}」")
        if not items:
            continue
        used_sheets.append(sheet_name)
        for key, item in items.items():
            if key in merged:
                raise ValueError(
                    f"不同工作表中有重复项: {key}（工作表「{sheet_name}」），请分开对比"
                )
            merged[key] = item
    if not used_sheets:
        raise ValueError("在任何工作表里都找不到清单表头（已逐表扫描）")
    return merged


# ---------------------------------------------------------------------------
# Diff
# ---------------------------------------------------------------------------

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
