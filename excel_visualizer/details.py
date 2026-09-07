"""每张图的统计说明与实时列映射核对。"""

from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTextBrowser, QPushButton, QHBoxLayout, QFormLayout, QComboBox
from html import escape

from excel_visualizer.data_service import FIELDS, FIELD_LABELS, guess_columns
from excel_visualizer.issues_service import ISSUE_LABELS, ISSUE_FIELDS, guess_issue_columns


def monthly_titles(year, month):
    return tuple(f"{prefix}{year}年{month}月Uptime&Run货量" for prefix in
                 ("PX产品", "光学产品", "T2～未入质保已经Run机台"))


def detail_text(page, index, monthly):
    if monthly:
        title = monthly_titles(page.year_box.value(), page.month_box.value())[index]
        fields, labels = FIELDS, FIELD_LABELS
        purpose = "比较各机台可用率与跑货量"
        filters = ("未标黄，山头属于 XRF、XPS、XRD、AFM", "未标黄，山头属于 BFI、DFI、DBO、IBO、eMBI、PC、MBI", "客户单元格直接黄色填充，不限制山头")[index]
        filters += "；表格1须为所选月份数据，未提供日期列，不自动按日期筛选"
        x, y = "机台编码 / 山头 / 客户名称（三行）", "左轴 Uptime（%）折线；右轴 Run货量柱形"
        formula = "逐行保留；比例 Uptime × 100，百分制原值；Run货量原值，不合计、不平均"
        notes = "空值保留为缺口，提示机台编码和 Excel 行号；错误数值阻止生成。自动模式把 0～1 视为比例；若原值 0.5 表示 0.5%，请选择百分制。黄色识别不支持条件格式。每个分类须有有效数值才可导出三页。"
        guess = guess_columns
    else:
        titles = ("各山头总问题数", "各山头月度问题出现次数及其变化趋势", "各个山头的问题密度", "各山头关闭率")
        title = f"{page.year}年{titles[index]}"
        fields = (("product", "created"), ("product", "created"), ("product", "created", "machine"), ("product", "created", "status"))[index]
        labels = [dict(zip(ISSUE_FIELDS, ISSUE_LABELS))[f] for f in fields]
        purpose = ("比较问题总量", "观察各山头月度变化", "比较每台机器的问题数量", "比较问题关闭情况")[index]
        filters = f"创建时间为 {page.year} 年；仅绘制有效数据实际出现的山头"
        if index == 1:
            filters += f"；月份连续显示 {page.start_month}～{page.end_month} 月"
        x = "月份" if index == 1 else "山头"
        y = ("问题数", "每月问题数", "问题数 / 台", "关闭率（%）")[index]
        formula = ("按山头分组的有效行数，每行一个问题", "按山头和月份分组计数；缺失月份补0", "该山头有效行数 / 非空机台编码去重数", "状态为申请关闭、已结束、已取消的行数 / 该山头有效行数 × 100%")[index]
        notes = "四图使用同一年份，年度指标包含全年有效记录；状态去除首尾空格；零分母显示 N/A；未知山头提示后保留；无效日期列出行号并阻止生成。"
        guess = guess_issue_columns
    rows = [("图表名称", title), ("统计目的", purpose), ("数据筛选条件", filters), ("横坐标", x), ("纵坐标", y), ("计算公式", formula), ("涉及字段", "、".join(labels))]
    guesses = guess(page.sheet) if page.sheet else {}
    mappings = []
    for field, label in zip(fields, labels):
        i = guesses.get(field, -1)
        auto = page.sheet.columns[i] if page.sheet and i >= 0 else "未匹配，请手动选择"
        current = page.column_boxes[field].currentText() or "未选择"
        mappings.append(f"{label}：自动匹配 {auto}；当前选择 {current}")
    rows += [("当前自动匹配的Excel列", "\n".join(mappings)), ("注意事项", notes + " 同一步骤共用字段映射，修改后所有相关图表失效，需重新生成。")]
    return "".join(f"<p><b>{escape(k)}</b><br>{escape(v).replace(chr(10), '<br>')}</p>" for k, v in rows)


def show_detail(page, index, monthly):
    dialog = QDialog(page)
    dialog.setWindowTitle("图表 detail · 统计口径与列映射")
    dialog.resize(760, 680)
    layout = QVBoxLayout(dialog)
    browser = QTextBrowser()
    browser.setHtml(detail_text(page, index, monthly))
    layout.addWidget(browser)
    form = QFormLayout()
    fields = FIELDS if monthly else (("product", "created"), ("product", "created"),
                                    ("product", "created", "machine"), ("product", "created", "status"))[index]
    labels = dict(zip(FIELDS if monthly else ISSUE_FIELDS, FIELD_LABELS if monthly else ISSUE_LABELS))
    for field in fields:
        source = page.column_boxes[field]
        box = QComboBox()
        for i in range(source.count()):
            box.addItem(source.itemText(i), source.itemData(i))
        box.setCurrentIndex(source.currentIndex())
        box.setEnabled(page.sheet is not None)
        def update(value, original=source):
            original.setCurrentIndex(value)
            browser.setHtml(detail_text(page, index, monthly))
        box.currentIndexChanged.connect(update)
        form.addRow(labels[field], box)
    layout.addLayout(form)
    close = QPushButton("关闭")
    close.clicked.connect(dialog.accept)
    layout.addWidget(close)
    dialog.exec()


def add_chart_actions(layout, page, count, generate, *, monthly):
    row = QHBoxLayout()
    buttons = []
    for index in range(count):
        button = QPushButton(f"制作图{index + 1}")
        detail = QPushButton("detail")
        def run(checked=False, i=index):
            generate()
            if page.report is not None:
                page.tabs.setCurrentIndex(i if monthly else i + 1)
        button.clicked.connect(run)
        detail.clicked.connect(lambda checked=False, i=index: show_detail(page, i, monthly))
        row.addWidget(button)
        row.addWidget(detail)
        buttons.append(button)
    layout.addLayout(row)
    return buttons
