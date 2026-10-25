# boq-diff

**工程量清单版本对比工具** — 把两个版本的工程量清单（招标版 / 施工版 / 结算版……）
丢进来，几秒钟自动找完所有不同，再也不用拿着几千行 Excel 一行一行对。

*A command-line tool that diffs two versions of a bill of quantities (BOQ)
and generates a colour-highlighted change report.*

## 它解决什么问题

工程量清单动不动几千行，而且会改好几版。人工逐项核对两个版本——哪项新增了、
哪项工程量变了、哪项单价涨了、总价差了多少——费时又容易漏，漏一项就可能几万块
对不上账。boq-diff 把这件事自动化：

- 🟢 **新增项**：新版有、旧版没有
- 🔴 **删除项**：旧版有、新版没有
- 🟡 **变更项**：工程量变化、综合单价变化、合价变化，逐项列明
- 📊 **汇总**：两版总价和差额，一眼看清

## 快速开始

```bash
pip install openpyxl   # 唯一的运行依赖
git clone https://github.com/zjh2066076218/boq-diff.git
cd boq-diff

python -m boq_diff examples/old.csv examples/new.csv
```

输出：

```text
对比完成：
- 旧版清单：5 项，合价总计 215,400.00 元
- 新版清单：5 项，合价总计 231,900.00 元
- 总价变化：+16,500.00 元
- 新增 1 项 / 删除 1 项 / 变更 2 项 / 未变 2 项
报告已生成: boq-report/report.md 和 boq-report/report.xlsx
```

同时生成两份报告（默认在 `boq-report/` 目录，可用 `--out-dir` 修改）：

- `report.md` — Markdown 差异报告，可直接贴进文档或发给同事
- `report.xlsx` — 带颜色标亮的 Excel（新增=绿、删除=红、变更=黄），含「汇总」
  和「差异明细」两张表

也可以安装成命令：

```bash
pip install .
boq-diff 旧版清单.xlsx 新版清单.xlsx --out-dir 报告目录
```

## 输入格式

支持 `.csv` 和 `.xlsx`，第一行必须是表头。识别以下列（中英文均可）：

| 含义 | 可用的表头名 | 必需 |
|---|---|---|
| 清单编码 | 清单编码 / 项目编码 / 编码 / code | 与项目名称二选一 |
| 项目名称 | 项目名称 / 清单名称 / 名称 / name | 与清单编码二选一 |
| 单位 | 单位 / unit | 否 |
| 工程量 | 工程量 / 数量 / quantity | ✅ |
| 综合单价 | 综合单价 / 单价 / unit_price | ✅ |
| 合价 | 合价 / 金额 / amount | 否，缺省时按 工程量 × 综合单价 计算 |

匹配规则：优先按**清单编码**配对；没有编码时按**项目名称 + 单位**配对。
同一份清单里出现重复编码会直接报错提醒，避免静默算错。

金额计算内部使用 `Decimal`，不会出现浮点误差。

## 作为库使用

```python
from boq_diff import load_boq, diff_boq

old = load_boq("招标清单.xlsx")
new = load_boq("结算清单.xlsx")
result = diff_boq(old, new)

print(len(result.added), len(result.removed), len(result.changed))
for change in result.changed:
    print(change.new.name, change.amount_delta)
```

## 运行测试

```bash
python -m unittest discover -s tests
```

## 路线图（欢迎 PR）

- [ ] 支持广联达等造价软件导出的多层级清单结构
- [ ] 单价分析：人材机组成变化对比
- [ ] 网页版：拖拽上传，在线看报告
- [ ] 变化幅度阈值过滤（只看变化超过 x% 的项）

## License

MIT
