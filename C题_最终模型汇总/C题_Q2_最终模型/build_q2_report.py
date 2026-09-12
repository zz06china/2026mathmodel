from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import openpyxl
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from latex2mathml.converter import convert as latex_to_mathml
from lxml import etree


HERE = Path(__file__).resolve().parent
RESULT = HERE / "result2_optimized.xlsx"
DOCX = HERE / "C题第二问_优化模型与结果.docx"
MD = HERE / "第二问优化模型与结果.md"
FIG_FORECAST = HERE / "图1_预测误差对比.png"
FIG_SENSITIVITY = HERE / "图2_执行方式费用比较.png"
MML2OMML = Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL")

BLUE = "2F5597"
LIGHT_BLUE = "D9EAF7"
ORANGE = "ED7D31"
GRAY = "666666"

DATES = [datetime(2025, 3, 20), datetime(2025, 6, 21), datetime(2025, 9, 23), datetime(2025, 12, 21)]
PLAN_TIMES = ["10:00-10:10", "12:00-12:10", "14:00-14:10", "16:00-16:10", "18:00-18:10", "20:00-20:10"]
COMPARISON = [
    ("上一版固定储能", 13.5704, 1.4948, 15.0652),
    ("闭环计划+固定执行", 13.4603, 1.4949, 14.9552),
    ("闭环计划+实时储能", 13.4603, 0.6590, 14.1193),
]


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if edge not in kwargs:
            continue
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        for key, value in kwargs[edge].items():
            element.set(qn("w:" + key), str(value))


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_run_font(run, chinese="宋体", western="Times New Roman", size=10.5, bold=False, color=None):
    run.font.name = western
    run._element.rPr.rFonts.set(qn("w:eastAsia"), chinese)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def add_heading(doc, text, level):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(8 if level == 1 else 5)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(text)
    set_run_font(run, chinese="黑体", size=15 if level == 1 else 12, bold=True, color=BLUE if level == 1 else None)
    return p


def add_body(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(21)
    p.paragraph_format.line_spacing = 1.35
    p.paragraph_format.space_after = Pt(2)
    if bold_prefix and text.startswith(bold_prefix):
        r = p.add_run(bold_prefix)
        set_run_font(r, bold=True)
        r = p.add_run(text[len(bold_prefix) :])
        set_run_font(r)
    else:
        r = p.add_run(text)
        set_run_font(r)
    return p


_transform = None


def latex_omml(latex):
    global _transform
    if _transform is None:
        _transform = etree.XSLT(etree.parse(str(MML2OMML)))
    mathml = latex_to_mathml(latex)
    return deepcopy(_transform(etree.fromstring(mathml.encode("utf-8"))).getroot())


def add_equation(doc, latex, number):
    table = doc.add_table(rows=1, cols=3)
    table.autofit = False
    widths = [Cm(0.5), Cm(13.5), Cm(1.0)]
    for cell, width in zip(table.rows[0].cells, widths):
        cell.width = width
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = table.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p._p.append(latex_omml(latex))
    r = table.cell(0, 2).paragraphs[0].add_run(f"({number})")
    set_run_font(r, size=10.5)
    table.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    set_run_font(r, chinese="黑体", size=10.5)


def add_table(doc, caption, headers, rows, font_size=9):
    add_caption(doc, caption)
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_repeat_table_header(table.rows[0])
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_shading(cell, LIGHT_BLUE)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(header))
        set_run_font(r, chinese="黑体", size=font_size, bold=True)
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(value))
            set_run_font(r, size=font_size)
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(
                cell,
                top={"val": "single", "sz": "4", "color": "BFBFBF"},
                bottom={"val": "single", "sz": "4", "color": "BFBFBF"},
                left={"val": "nil"},
                right={"val": "nil"},
            )
    return table


def add_symbol_table(doc):
    rows = [
        (r"d,t", "日期索引、10 min时段索引", "-"),
        (r"L_{d,t},G_{d,t}", "实际负载电量、实际光伏电量", "kWh"),
        (r"\widehat L_{d,t},\widehat G_{d,t}", "0:00时可得的负载、光伏点预测", "kWh"),
        (r"\widetilde N_{d,t}", "经风险分位修正的净负荷", "kWh"),
        (r"B_{d,t}", "日前计划购电量", "kWh"),
        (r"C^p_{d,t},D^p_{d,t}", "日前优化中的预期充电量、放电量", "kWh"),
        (r"C_{d,t},D_{d,t}", "实际运行中的充电量、放电量", "kWh"),
        (r"R_{d,t}", "实际运行中的紧急购电量", "kWh"),
        (r"S_{d,t}", "实际富余电量", "kWh"),
        (r"E_{d,t}", "时段起点储电量", "kWh"),
        (r"p_t", "计划购电单价", "元/kWh"),
        (r"\eta_c,\eta_d", "充电效率、放电效率", "-"),
    ]
    add_caption(doc, "表2 主要符号说明")
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(["符号", "含义", "单位"]):
        set_cell_shading(table.rows[0].cells[i], LIGHT_BLUE)
        p = table.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(h), chinese="黑体", bold=True)
    for latex, meaning, unit in rows:
        cells = table.add_row().cells
        cells[0].text = ""
        cells[0].paragraphs[0]._p.append(latex_omml(latex))
        cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for idx, value in [(1, meaning), (2, unit)]:
            p = cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if idx == 2 else WD_ALIGN_PARAGRAPH.LEFT
            set_run_font(p.add_run(value), size=9.5)
    return table


def read_result():
    wb = openpyxl.load_workbook(RESULT, data_only=True)
    plan = wb["计划购电量"]
    headers = [cell.value for cell in plan[1]]
    plan_rows = {}
    for target in DATES:
        row = next(r for r in range(2, plan.max_row + 1) if plan.cell(r, 1).value == target)
        values = [plan.cell(row, headers.index(label) + 1).value for label in PLAN_TIMES]
        plan_rows[target] = values + [plan.cell(row, 146).value, plan.cell(row, 147).value]

    storage = wb["充放电量"]
    storage_rows = {}
    for target in DATES:
        row = next(r for r in range(2, storage.max_row + 1) if storage.cell(r, 1).value == target)
        storage_rows[target] = {
            "blocks": [(storage.cell(row + k, 2).value, storage.cell(row + k, 3).value, storage.cell(row + k, 4).value) for k in range(6)],
            "start": storage.cell(row, 6).value,
            "end": storage.cell(row + 1, 6).value,
        }

    emergency = wb["紧急购电量"]
    emergency_rows = {}
    for target in DATES:
        start = next(
            (r for r in range(2, emergency.max_row + 1) if emergency.cell(r, 1).value == target),
            None,
        )
        if start is None:
            emergency_rows[target] = []
            continue
        stop = next((r for r in range(start + 1, emergency.max_row + 1) if emergency.cell(r, 1).value is not None), emergency.max_row + 1)
        emergency_rows[target] = [(emergency.cell(r, 2).value, emergency.cell(r, 3).value) for r in range(start, stop)]
    wb.close()
    return plan_rows, storage_rows, emergency_rows


def make_figures():
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False
    labels = ["负载", "光伏"]
    adaptive = [3.83, 7.18]
    persistence = [14.08, 7.79]
    x = range(2)
    fig, ax = plt.subplots(figsize=(7.2, 3.6), dpi=180)
    ax.bar([v - 0.18 for v in x], persistence, 0.36, label="昨日持续法", color="#A5A5A5")
    ax.bar([v + 0.18 for v in x], adaptive, 0.36, label="自适应滞后融合", color="#4472C4")
    ax.set_xticks(list(x), labels)
    ax.set_ylabel("WMAPE (%)")
    ax.legend(frameon=False)
    ax.grid(axis="y", alpha=0.25)
    fig.tight_layout()
    fig.savefig(FIG_FORECAST, bbox_inches="tight")
    plt.close(fig)

    labels = [x[0] for x in COMPARISON]
    plan = [x[1] for x in COMPARISON]
    emergency = [x[2] for x in COMPARISON]
    total = [x[3] for x in COMPARISON]
    fig, ax = plt.subplots(figsize=(7.2, 3.8), dpi=180)
    x = range(len(labels))
    ax.bar([v - 0.22 for v in x], plan, 0.22, label="计划购电费", color="#4472C4")
    ax.bar(list(x), emergency, 0.22, label="紧急购电费", color="#ED7D31")
    ax.bar([v + 0.22 for v in x], total, 0.22, label="总购电费", color="#70AD47")
    ax.set_xticks(list(x), labels)
    ax.set_ylabel("费用（百万元）")
    ax.legend(frameon=False, ncol=3)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(FIG_SENSITIVITY, bbox_inches="tight")
    plt.close(fig)


def add_picture(doc, path, caption):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path), width=Cm(13.5))
    add_caption(doc, caption)


def fmt(value):
    return f"{float(value):,.4f}"


def build_markdown(plan_rows, storage_rows, emergency_rows):
    lines = [
        "# C题问题二：模型核查与完整建模",
        "",
        "## 1. 核查结论",
        "",
        "原程序的两阶段思想是合理起点，但不能直接作为最终模型。主要问题是：时间轴错位10 min；仅用昨日数据导致负载预测误差偏大；1月1日使用当天真实值造成信息泄漏；逐日模型无期末价值使储能迅速降至1200 kWh；执行层读取全天真实数据后重优化储能，违反0:00决策的因果性；弃光与超购富余混用且无上界；单一紧急购电目标存在退化解；缺少求解状态和逐时段约束校验。",
        "",
        "最终采用：自适应滞后融合预测、28日滚动净负荷误差80%分位修正、48小时滚动线性规划、计划购电锁定后的逐时段实时储能纠偏，以及词典序二阶段去退化。",
        "",
        "## 2. 预测与风险修正",
        "",
        r"对负载和光伏分别使用 $\hat X_{d,t}=\alpha X_{d-1,t}+(1-\alpha)X_{d-7,t}$，$\alpha\in\{0,0.1,\ldots,1\}$，仅依据计划日前最多56天历史误差选择。令净负荷预测误差为 $\varepsilon_{d,t}=(L_{d,t}-G_{d,t})-(\hat L_{d,t}-\hat G_{d,t})$，以最近28天同一时段误差的80%分位修正净负荷。由于计划电量边际成本为 $p_t$，短缺电量成本为 $5p_t$，一阶条件为 $F(q)=1-1/5=0.8$。",
        "",
        "## 3. 48小时滚动优化模型",
        "",
        r"目标函数为 $\min\sum_{h=0}^{1}\sum_t p_t B_{d+h,t}$。日前模型中的预期储能变量满足 $B_{d+h,t}+D^p_{d+h,t}=\widetilde N_{d+h,t}+C^p_{d+h,t}+S^p_{d+h,t}$ 及SOC递推和边界约束。48小时末设置6000 kWh锚点，只执行首日购电决策，次日以实际SOC重新滚动求解。保持最低费用后再最小化预期充放电吞吐量与富余量。",
        "",
        r"实际运行不读取未来真实数据。计划购电量 $B$ 在0:00锁定；每个时段只根据当前已观测净缺口，在功率和SOC约束内优先放电，出现富余时优先充电，剩余缺口才形成紧急购电。总费用为 $\sum p_tB_{d,t}+5\sum p_tR_{d,t}$。当天实际期末SOC作为次日滚动优化初值。",
        "",
        "## 4. 核心结果",
        "",
        "- 报告期：2025-02-01至2025-12-31，共334天。",
        "- 计划购电量：21,963,546.7098 kWh；紧急购电量：110,213.6503 kWh。",
        "- 计划购电费：13,460,292.3327元；紧急购电费：658,979.2042元；总费用：14,119,271.5369元。",
        "- 相较上一版，总费用降低945,928.1670元（6.28%），紧急购电费降低835,803.4227元（55.92%）。",
        "- 负载预测WMAPE由14.08%降至3.83%；光伏由7.79%降至7.18%。",
        "- 全年无同时充放电；最大能量平衡残差小于3.5e-13 kWh；SOC范围为1200至10800 kWh。",
        "",
        "## 5. 指定日期结果",
        "",
        "| 日期 | 10:00 | 12:00 | 14:00 | 16:00 | 18:00 | 20:00 | 全天购电量 | 全天购电费 |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for date in DATES:
        lines.append("| " + date.strftime("%Y-%m-%d") + " | " + " | ".join(fmt(x) for x in plan_rows[date]) + " |")
    lines += ["", "### 储能结果", ""]
    for date in DATES:
        data = storage_rows[date]
        lines += [f"**{date:%Y-%m-%d}，0:00 SOC={fmt(data['start'])} kWh，24:00 SOC={fmt(data['end'])} kWh**", "", "| 时段 | 充电量 | 放电量 |", "|---|---:|---:|"]
        lines += [f"| {label} | {fmt(c)} | {fmt(d)} |" for label, c, d in data["blocks"]]
        lines.append("")
    lines += ["### 紧急购电结果", "", "| 日期 | 时间段 | 购电量/kWh |", "|---|---|---:|"]
    for date in DATES:
        events = emergency_rows[date]
        if not events:
            lines.append(f"| {date:%Y-%m-%d} | 无 | 0.0000 |")
        for i, (period, amount) in enumerate(events):
            lines.append(f"| {date:%Y-%m-%d}" if i == 0 else "| ")
            lines[-1] += f" | {period} | {fmt(amount)} |"
    lines += ["", "完整逐日逐时段结果见 `result2_optimized.xlsx`，可复现代码见 `q2_final.py`。", ""]
    MD.write_text("\n".join(lines), encoding="utf-8")


def build_docx(plan_rows, storage_rows, emergency_rows):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    normal.font.size = Pt(10.5)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_end])
    set_run_font(run, size=9, color=GRAY)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(14)
    set_run_font(title.add_run("问题二  微网日前计划购电的滚动风险优化模型"), chinese="黑体", size=18, bold=True, color=BLUE)

    add_heading(doc, "1 模型核查与修正", 1)
    add_body(doc, "原方案将问题拆为计划层和执行层，框架方向正确；但预测、边界和执行方式存在会改变答案的缺陷，不能直接作为最终结果。核查与修正如下。")
    audit_rows = [
        ("时间轴", "直接按附件列顺序计算", "整体错位10 min", "自然日旋转，输出时恢复模板顺序"),
        ("预测", "当天等于昨日", "负载WMAPE为14.08%", "滞后1日与7日自适应融合"),
        ("信息集", "1月1日使用当天真实值", "存在信息泄漏", "前7日使用附件1先验，之后只用历史"),
        ("日末SOC", "日模型无期末价值", "电池迅速降至1200 kWh", "48 h滚动窗口与末端锚点"),
        ("实际执行", "看见全天实际量后重优化", "使用未来信息", "锁定购电，储能按当前观测实时纠偏"),
        ("退化解", "只最小化费用", "可能出现无意义循环", "保持最优费用后最小化吞吐与富余"),
        ("结果校验", "未检查状态和残差", "错误不易暴露", "逐日检查可行性、SOC、平衡和同时充放电"),
    ]
    add_table(doc, "表1 原模型主要问题及修正", ["环节", "原处理", "影响", "最终处理"], audit_rows, 8.5)

    add_heading(doc, "2 数据预处理与基本假设", 1)
    add_body(doc, "每个10 min时段的功率乘以1/6 h转为电量。附件和模板以0:10开头、0:00+1结尾；求解前将末列旋转至首列，使状态E的起点严格对应0:00，输出时再逆变换。")
    add_equation(doc, r"\Delta t=\frac{1}{6}\ \mathrm{h},\qquad L_{d,t}=P^L_{d,t}\Delta t,\quad G_{d,t}=P^{PV}_{d,t}\Delta t", 1)
    add_body(doc, "假设电价日内曲线每日重复；计划购电量在0:00锁定，储能可根据当前时段已经观测到的负载与光伏进行实时纠偏，但不得读取未来真实值；计划购电不退费，紧急购电按对应时段电价的5倍计费。")

    add_heading(doc, "3 预测与风险修正", 1)
    add_heading(doc, "3.1 严格因果的点预测", 2)
    add_body(doc, "对负载和光伏分别融合昨日同一时段与上周同日同一时段。权重从0到1按0.1取值，并仅用计划日前最多56天的历史误差滚动选择。")
    add_equation(doc, r"\widehat X_{d,t}=\alpha_dX_{d-1,t}+(1-\alpha_d)X_{d-7,t},\qquad X\in\{L,G\}", 2)
    add_equation(doc, r"\alpha_d=\mathop{\mathrm{arg\,min}}_{\alpha\in\{0,0.1,\ldots,1\}}\sum_{j=d-m}^{d-1}\sum_t\left|X_{j,t}-\widehat X_{j,t}(\alpha)\right|", 3)
    add_body(doc, "48小时窗口中的次日预测使用计划时刻已经可见的滞后2日和滞后7日数据，仍不调用未来实际值。")
    add_heading(doc, "3.2 紧急购电风险分位", 2)
    add_body(doc, "定义净负荷点预测误差，并取最近28天同一时段误差的80%经验分位作为安全修正。")
    add_equation(doc, r"\varepsilon_{d,t}=(L_{d,t}-G_{d,t})-(\widehat L_{d,t}-\widehat G_{d,t})", 4)
    add_equation(doc, r"\widetilde N_{d,t}=\widehat L_{d,t}-\widehat G_{d,t}+Q_{0.8}\!\left(\varepsilon_{d-28:d-1,t}\right)", 5)
    add_body(doc, "设某时段的拟保障电量为q。多计划1 kWh需支付正常时段单价，少计划1 kWh则需支付其5倍的紧急单价。相应单时段期望费用的一阶最优条件为：")
    add_equation(doc, r"F_N(q^*)=1-\frac{p_t}{5p_t}=0.8", 6)
    add_symbol_table(doc)

    add_heading(doc, "4 48小时滚动线性规划", 1)
    add_body(doc, "每天0:00以当前SOC为初值，联合优化当天和次日共288个时段。次日变量只用于刻画日末电量的延续价值，实际仅执行当天144个时段；下一天0:00重新预测并滚动求解。")
    add_equation(doc, r"\min Z_d^{\mathrm{plan}}=\sum_{h=0}^{1}\sum_{t=0}^{143}p_tB_{d+h,t}", 7)
    add_body(doc, "风险修正后的逐时段能量平衡为：")
    add_equation(doc, r"B_{d+h,t}+D^p_{d+h,t}=\widetilde N_{d+h,t}+C^p_{d+h,t}+S^p_{d+h,t}", 8)
    add_body(doc, "储能状态递推及运行边界为：")
    add_equation(doc, r"E^p_{k+1}=E^p_k+\eta_cC^p_k-\frac{D^p_k}{\eta_d},\qquad \eta_c=\eta_d=0.9", 9)
    add_equation(doc, r"1200\le E^p_k\le10800,\qquad 0\le C^p_k,D^p_k\le5000\Delta t=833.3333", 10)
    add_equation(doc, r"E_{d,0}=E_d^{\mathrm{obs}},\qquad E_{d+2,0}=6000", 11)
    add_body(doc, "式（11）的6000 kWh只位于滚动窗口远端，用于抑制有限时域末端效应，并不强制每日24:00回到6000 kWh。获得最低计划费用后，在费用增加不超过0.001元的条件下，再最小化充电量、放电量和计划富余量之和，从而得到物理含义明确的非退化解。")

    add_heading(doc, "5 因果实时控制与费用", 1)
    add_body(doc, "原代码的执行LP读取当天全部真实负载和光伏后重排储能，相当于0:00预知未来。改进模型只锁定计划购电量；到达时段t后，先观测该时段实际净缺口，再按下式调用储能。该规则不含任何未来实际数据。")
    add_equation(doc, r"n_{d,t}=L_{d,t}-G_{d,t}-B_{d,t}", 12)
    add_equation(doc, r"D_{d,t}=\min\left\{[n_{d,t}]^+,Q^{\max},\eta_d(E_{d,t}-E^{\min})\right\}", 13)
    add_equation(doc, r"C_{d,t}=\min\left\{[-n_{d,t}]^+,Q^{\max},\frac{E^{\max}-E_{d,t}}{\eta_c}\right\}", 14)
    add_equation(doc, r"R_{d,t}=[n_{d,t}-D_{d,t}]^+,\qquad S_{d,t}=[-n_{d,t}-C_{d,t}]^+", 15)
    add_equation(doc, r"E_{d,t+1}=E_{d,t}+\eta_cC_{d,t}-\frac{D_{d,t}}{\eta_d},\qquad E_{d+1,0}=E_{d,144}", 16)
    add_equation(doc, r"Z=\sum_d\sum_t p_tB_{d,t}+5\sum_d\sum_t p_tR_{d,t}", 17)

    add_heading(doc, "6 求解结果与检验", 1)
    add_heading(doc, "6.1 预测精度", 2)
    pred_rows = [
        ("负载", "108.3367", "14.08%", "29.4660", "3.83%"),
        ("光伏", "30.8911", "7.79%", "28.4676", "7.18%"),
    ]
    add_table(doc, "表3 预测方法比较", ["对象", "持续法MAE/kWh", "持续法WMAPE", "本文MAE/kWh", "本文WMAPE"], pred_rows, 8.5)
    add_picture(doc, FIG_FORECAST, "图1 预测误差对比")

    add_heading(doc, "6.2 全周期费用与可行性", 2)
    summary_rows = [
        ("计划购电量", "21,963,546.7098 kWh"),
        ("紧急购电量", "110,213.6503 kWh"),
        ("计划购电费", "13,460,292.3327 元"),
        ("紧急购电费", "658,979.2042 元"),
        ("总购电费", "14,119,271.5369 元"),
        ("紧急购电发生天数", "108/334"),
        ("年末SOC", "8,851.0667 kWh"),
    ]
    add_table(doc, "表4 2025年2月1日至12月31日汇总结果", ["指标", "数值"], summary_rows, 9.5)
    add_body(doc, "逐时段检查表明：最大供需平衡残差为4.55×10⁻¹³ kWh，最大SOC递推残差为9.10×10⁻¹³ kWh；全年SOC范围为1200至10800 kWh，不存在同一10 min时段同时充放电。紧急购电仅发生于108天，且总量较小，说明实时储能纠偏消化了大部分预测偏差。")

    add_heading(doc, "6.3 优化效果比较", 2)
    compare_rows = [(name, f"{p:.4f}", f"{e:.4f}", f"{z:.4f}") for name, p, e, z in COMPARISON]
    add_table(doc, "表5 实时储能纠偏的费用改善", ["方案", "计划费用/百万元", "紧急费用/百万元", "总费用/百万元"], compare_rows, 9)
    add_picture(doc, FIG_SENSITIVITY, "图2 不同执行方式的费用比较")
    add_body(doc, "相较上一版固定日前储能计划，闭环实时控制使紧急购电费减少835,803.4227元，降幅55.92%；总费用减少945,928.1670元，降幅6.28%。计划费用同时小幅下降，是因为次日滚动规划使用实际期末SOC，而不再继承一条未实际执行的计划SOC轨迹。")

    add_heading(doc, "7 指定日期结果", 1)
    plan_table_rows = []
    for date in DATES:
        plan_table_rows.append([date.strftime("%Y-%m-%d"), *[fmt(v) for v in plan_rows[date]]])
    add_table(doc, "表6 指定日期的计划购电量", ["日期", "10:00", "12:00", "14:00", "16:00", "18:00", "20:00", "全天/kWh", "费用/元"], plan_table_rows, 7.5)

    storage_table_rows = []
    for date in DATES:
        data = storage_rows[date]
        for i, (label, charge, discharge) in enumerate(data["blocks"]):
            storage_table_rows.append([date.strftime("%Y-%m-%d") if i == 0 else "", label, fmt(charge), fmt(discharge), fmt(data["start"]) if i == 0 else (fmt(data["end"]) if i == 1 else "")])
    add_table(doc, "表7 指定日期的储能运行结果", ["日期", "时间段", "充电量/kWh", "放电量/kWh", "SOC（0:00/24:00）"], storage_table_rows, 8)

    emergency_table_rows = []
    for date in DATES:
        events = emergency_rows[date]
        if not events:
            emergency_table_rows.append([date.strftime("%Y-%m-%d"), "无", "0.0000"])
        for i, (period, amount) in enumerate(events):
            emergency_table_rows.append([date.strftime("%Y-%m-%d") if i == 0 else "", period, fmt(amount)])
    add_table(doc, "表8 指定日期的紧急购电结果", ["日期", "紧急购电时间段", "紧急购电量/kWh"], emergency_table_rows, 8.5)

    add_heading(doc, "8 模型评价", 1)
    add_body(doc, "最终模型的优势在于日前决策只使用当时可得信息，实时储能控制也只调用当前观测；风险参数由价格机制推导，48小时窗口消除了逐日SOC耗尽问题。日前优化保持线性，实时规则只需常数次运算，适合全年滚动实施。局限在于实时规则以减少当前紧急购电为目标，尚未显式为未来高价时段保留SOC；进一步采用价格感知的短时MPC仍可能降低费用。问题三给出滚动光伏预报后，可在本框架上增加购电调整变量与违约成本。")

    doc.save(DOCX)


def main():
    plan_rows, storage_rows, emergency_rows = read_result()
    make_figures()
    build_markdown(plan_rows, storage_rows, emergency_rows)
    build_docx(plan_rows, storage_rows, emergency_rows)
    print(DOCX)
    print(MD)


if __name__ == "__main__":
    main()
