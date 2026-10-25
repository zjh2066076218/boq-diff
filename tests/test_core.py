"""Tests for boq_diff.core.  Runs with unittest or pytest."""

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from boq_diff import diff_boq, load_boq, total_amount

HEADER = "清单编码,项目名称,单位,工程量,综合单价\n"

OLD = HEADER + (
    "001,土方开挖,m3,1000,45.50\n"
    "002,C30混凝土,m3,100,520.00\n"
    "003,脚手架,m2,600,25.00\n"
)

NEW = HEADER + (
    "001,土方开挖,m3,1000,45.50\n"
    "002,C30混凝土,m3,120,520.00\n"
    "004,防水工程,m2,300,65.00\n"
)


def _write(text: str) -> Path:
    fh = tempfile.NamedTemporaryFile(
        "w", suffix=".csv", delete=False, encoding="utf-8"
    )
    fh.write(text)
    fh.close()
    return Path(fh.name)


class TestDiff(unittest.TestCase):
    def setUp(self):
        self.old = load_boq(_write(OLD))
        self.new = load_boq(_write(NEW))
        self.result = diff_boq(self.old, self.new)

    def test_added_removed_changed(self):
        self.assertEqual([i.code for i in self.result.added], ["004"])
        self.assertEqual([i.code for i in self.result.removed], ["003"])
        self.assertEqual([c.key for c in self.result.changed], ["002"])
        self.assertEqual(self.result.unchanged_count, 1)

    def test_change_deltas(self):
        change = self.result.changed[0]
        self.assertEqual(change.quantity_delta, Decimal("20"))
        self.assertEqual(change.amount_delta, Decimal("10400.00"))

    def test_amount_computed_when_missing(self):
        item = self.old["001"]
        self.assertEqual(item.amount, Decimal("45500.00"))
        self.assertEqual(total_amount(self.old), Decimal("112500.00"))

    def test_name_fallback_key(self):
        no_code = "项目名称,单位,工程量,综合单价\n抹灰,m2,10,20.00\n"
        items = load_boq(_write(no_code))
        self.assertIn("抹灰|m2", items)

    def test_duplicate_rejected(self):
        dup = HEADER + "001,甲,m,1,1.00\n001,乙,m,2,2.00\n"
        with self.assertRaises(ValueError):
            load_boq(_write(dup))

    def test_identical_files_have_no_changes(self):
        same = load_boq(_write(OLD))
        result = diff_boq(self.old, same)
        self.assertFalse(result.has_changes)
        self.assertEqual(result.unchanged_count, 3)


if __name__ == "__main__":
    unittest.main()
