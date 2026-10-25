"""Command line interface: ``boq-diff OLD NEW [--out-dir DIR]``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import diff_boq, load_boq, total_amount
from .report import render_markdown, summary_lines, write_xlsx


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="boq-diff",
        description="对比两个版本的工程量清单（CSV / XLSX），输出差异报告。",
    )
    parser.add_argument("old", help="旧版清单文件（.csv / .xlsx）")
    parser.add_argument("new", help="新版清单文件（.csv / .xlsx）")
    parser.add_argument(
        "--out-dir", default="boq-report",
        help="报告输出目录（默认: boq-report）",
    )
    args = parser.parse_args(argv)

    try:
        old_items = load_boq(args.old)
        new_items = load_boq(args.new)
    except (ValueError, RuntimeError, FileNotFoundError) as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2

    result = diff_boq(old_items, new_items)

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    md_path = out_dir / "report.md"
    md_path.write_text(
        render_markdown(old_items, new_items, result,
                        old_label=str(args.old), new_label=str(args.new)),
        encoding="utf-8",
    )
    xlsx_path = write_xlsx(out_dir / "report.xlsx", old_items, new_items, result)

    print("对比完成：")
    for line in summary_lines(old_items, new_items, result):
        print(line)
    print(f"旧版总价 {total_amount(old_items):,.2f} 元 → "
          f"新版总价 {total_amount(new_items):,.2f} 元")
    print(f"报告已生成: {md_path} 和 {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
