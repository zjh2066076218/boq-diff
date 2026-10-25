"""Tests for messy real-world exports: title rows, two-row headers,
merged cells, multiple sheets, GBK encoding, subtotal rows, legacy .xls."""

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from boq_diff import load_boq


def _tmp(suffix: str) -> Path:
    fh = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    fh.close()
    return Path(fh.name)


class TestGbkCsvWithTitles(unittest.TestCase):
    def test_titles_sections_subtotals_and_thousands(self):
        text = (
            "工程量清单表,,,,,,\n"
            "项目名称：示例工程,,,,,,\n"
            "序号,项目编码,项目名称,计量单位,工程数量,综合单价,合价\n"
            ",,一、土石方工程,,,,\n"
            '1,001001,土方开挖,m3,"1,000",45.50,"45,500.00"\n'
            "2,001002,C30混凝土,m3,100,520.00,52000.00\n"
            ",,分部分项小计,,,,97500.00\n"
        )
        path = _tmp(".csv")
        path.write_bytes(text.encode("gbk"))
        items = load_boq(path)
        self.assertEqual(sorted(items), ["001001", "001002"])
        self.assertEqual(items["001001"].quantity, Decimal("1000"))
        self.assertEqual(items["001001"].amount, Decimal("45500.00"))


class TestTwoRowHeaderXlsx(unittest.TestCase):
    def test_merged_parent_header_and_cover_sheet(self):
        openpyxl = __import__("openpyxl")
        wb = openpyxl.Workbook()
        cover = wb.active
        cover.title = "封面"
        cover["A1"] = "某工程 工程量清单"
        cover["A2"] = "编制单位：示例公司"

        ws = wb.create_sheet("分部分项工程量清单")
        ws.merge_cells("A1:H1")
        ws["A1"] = "分部分项工程量清单表"
        headers = ["序号", "项目编码", "项目名称", "项目特征",
                   "计量单位", "工程数量", "金额", "金额"]
        for col, h in enumerate(headers, start=1):
            ws.cell(row=2, column=col, value=h)
        ws.merge_cells("G2:H2")
        ws.cell(row=3, column=7, value="综合单价")
        ws.cell(row=3, column=8, value="合价")
        ws.append([])  # row 4 placeholder to keep appends aligned below
        data = [
            [1, "010101001001", "土方开挖", "土壤类别:三类土", "m3", 1000, 45.5, 45500],
            [None, None, "一、混凝土工程", None, None, None, None, None],
            [2, "010302001001", "C30混凝土", "强度等级:C30", "m3", 100, 520, 52000],
            [None, None, "本页小计", None, None, None, None, 97500],
        ]
        for r, row in enumerate(data, start=5):
            for c, v in enumerate(row, start=1):
                ws.cell(row=r, column=c, value=v)
        path = _tmp(".xlsx")
        wb.save(path)

        items = load_boq(path)
        self.assertEqual(sorted(items), ["010101001001", "010302001001"])
        self.assertEqual(items["010302001001"].unit_price, Decimal("520"))
        self.assertEqual(items["010302001001"].amount, Decimal("52000"))


class TestLegacyXls(unittest.TestCase):
    def test_xls_with_title_row(self):
        try:
            import xlwt
        except ImportError:
            self.skipTest("xlwt not available to build the fixture")
        wb = xlwt.Workbook()
        ws = wb.add_sheet("清单")
        ws.write(0, 0, "工程量清单（老式 .xls 导出）")
        headers = ["项目编码", "项目名称", "单位", "工程量", "综合单价"]
        for c, h in enumerate(headers):
            ws.write(1, c, h)
        ws.write(2, 0, "001")
        ws.write(2, 1, "土方开挖")
        ws.write(2, 2, "m3")
        ws.write(2, 3, 1000)
        ws.write(2, 4, 45.5)
        path = _tmp(".xls")
        wb.save(str(path))
        items = load_boq(path)
        self.assertEqual(items["001"].amount, Decimal("45500.00"))


if __name__ == "__main__":
    unittest.main()
