"""第二份 Excel 的独立导入、字段确认、统计核对与导出页面。"""

from pathlib import Path

from PyQt6.QtCore import pyqtSignal
from PyQt6.QtWidgets import (
    QComboBox, QFileDialog, QGridLayout, QHBoxLayout, QHeaderView, QLabel,
    QMessageBox, QPushButton, QScrollArea, QSpinBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from excel_visualizer.data_service import ExcelDataError, NON_OPTICAL, OPTICAL, header_hint, list_sheets, read_sheet
from excel_visualizer.issues_service import ISSUE_ALIASES, ISSUE_FIELDS, ISSUE_LABELS, build_issues_report, guess_issue_columns
from excel_visualizer.issues_widget import IssuesChartWidget


class IssuesPage(QWidget):
    report_changed = pyqtSignal()

    def __init__(self) -> None:
        super().__init__()
        self.file_path = self.sheet = self.report = None
        self.products = NON_OPTICAL + OPTICAL
        self.open_button = QPushButton("导入第二份 Excel")
        self.file_label = QLabel("尚未选择问题统计文件")
        self.file_label.setWordWrap(True)
        self.sheet_box = QComboBox()
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 1048576)
        self.reload_button = QPushButton("读取工作表 / 刷新")
        self.column_boxes = {field: QComboBox() for field in ISSUE_FIELDS}
        for box in (self.sheet_box, *self.column_boxes.values()):
            box.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
            box.setMinimumContentsLength(10)
        self.scope_box = QComboBox()
        self.scope_box.addItem("全部年份（按整份问题表统计）", True)
        self.scope_box.addItem("仅 2026 年", False)
        self.plot_button = QPushButton("生成四张问题统计图表")
        self.export_button = QPushButton("导出第二步 PPT（1 页）")
        self.status_label = QLabel("请上传含山头、创建时间、产品序列号/机台编码、服务请求状态的 .xlsx。")
        self.status_label.setWordWrap(True)
        self.tabs = QTabWidget()
        self.overview = IssuesChartWidget()
        self.overview.setMinimumSize(880, 440)
        overview_scroll = QScrollArea()
        overview_scroll.setWidgetResizable(True)
        overview_scroll.setWidget(self.overview)
        self.tabs.addTab(overview_scroll, "四图预览")
        self.charts = [IssuesChartWidget(i) for i in range(4)]
        for title, chart in zip(("年度问题数", "月度趋势", "问题密度", "关闭率"), self.charts):
            self.tabs.addTab(chart, title)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.tabs.addTab(self.table, "数据核对")
        root = QVBoxLayout(self)
        files = QHBoxLayout()
        files.addWidget(self.open_button)
        files.addWidget(self.file_label, 1)
        root.addLayout(files)
        source = QHBoxLayout()
        source.addWidget(QLabel("工作表："))
        source.addWidget(self.sheet_box, 1)
        source.addWidget(QLabel("表头行："))
        source.addWidget(self.header_row)
        source.addWidget(self.reload_button)
        root.addLayout(source)
        fields = QGridLayout()
        for i, (field, label) in enumerate(zip(ISSUE_FIELDS, ISSUE_LABELS)):
            fields.addWidget(QLabel(label + "："), i // 2, i % 2 * 2)
            fields.addWidget(self.column_boxes[field], i // 2, i % 2 * 2 + 1)
        fields.addWidget(QLabel("密度 / 关闭率范围："), 2, 0)
        fields.addWidget(self.scope_box, 2, 1, 1, 3)
        root.addLayout(fields)
        hint = QLabel("前两图固定统计 2026 年，折线从 1 月到本月；PPT 四图在上半页按 2×2 排列，下半页留白。\n"
                      "问题逐行计数；机器按山头内非空编号去重；关闭状态：申请关闭、已结束、已取消。")
        hint.setWordWrap(True)
        root.addWidget(hint)
        actions = QHBoxLayout()
        actions.addWidget(self.plot_button)
        actions.addWidget(self.export_button)
        actions.addStretch()
        root.addLayout(actions)
        root.addWidget(self.status_label)
        root.addWidget(self.tabs, 1)
        self.open_button.clicked.connect(self._select_file)
        self.reload_button.clicked.connect(self._load_sheet)
        self.sheet_box.currentIndexChanged.connect(self._source_changed)
        self.header_row.valueChanged.connect(self._source_changed)
        self.scope_box.currentIndexChanged.connect(self._mapping_changed)
        for box in self.column_boxes.values():
            box.currentIndexChanged.connect(self._mapping_changed)
        self.plot_button.clicked.connect(self._generate_report)
        self.export_button.clicked.connect(self._export_report)
        for widget in (self.sheet_box, self.header_row, self.reload_button):
            widget.setEnabled(False)
        self._set_mapping_enabled(False)
        self.export_button.setEnabled(False)

    def set_products(self, products: tuple[str, ...]) -> None:
        names = tuple(dict.fromkeys(NON_OPTICAL + OPTICAL + products))
        if names != self.products:
            self.products = names
            self._mapping_changed()
            self.status_label.setText("第一步的山头名单已变化，请重新生成问题统计。")

    def _set_mapping_enabled(self, enabled: bool) -> None:
        for widget in (*self.column_boxes.values(), self.scope_box, self.plot_button):
            widget.setEnabled(enabled)

    def _invalidate_report(self) -> None:
        self.report = None
        self.export_button.setEnabled(False)
        for chart in (self.overview, *self.charts):
            chart.show_placeholder()
        self.report_changed.emit()

    def _source_changed(self) -> None:
        self.sheet = None
        self._invalidate_report()
        self._set_mapping_enabled(False)
        self.table.setRowCount(0)
        self.status_label.setText("工作表或表头行已改变，请点击“读取工作表 / 刷新”。")

    def _mapping_changed(self) -> None:
        self._invalidate_report()
        if self.sheet is not None:
            self._show_source_table()
            self.status_label.setText("字段或统计范围已改变，请重新生成问题统计。")

    def _select_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择第二份 Excel：问题统计", "", "Excel 文件 (*.xlsx)")
        if path:
            self.open_excel(path)

    def open_excel(self, path: str | Path) -> None:
        try:
            names = list_sheets(path)
        except ExcelDataError as error:
            QMessageBox.critical(self, "读取失败", str(error))
            return
        self.file_path = Path(path)
        self.file_label.setText(str(path))
        self.sheet_box.blockSignals(True)
        self.sheet_box.clear()
        self.sheet_box.addItems(names)
        self.sheet_box.blockSignals(False)
        for widget in (self.sheet_box, self.header_row, self.reload_button):
            widget.setEnabled(True)
        self._load_sheet()

    def _load_sheet(self) -> None:
        if self.file_path is None:
            return
        self._source_changed()
        try:
            sheet = read_sheet(self.file_path, self.sheet_box.currentText(), self.header_row.value())
        except ExcelDataError as error:
            self.status_label.setText("读取失败，请检查工作表和表头行。")
            QMessageBox.warning(self, "读取失败", str(error))
            return
        self.sheet = sheet
        guesses = guess_issue_columns(sheet)
        for field, box in self.column_boxes.items():
            previous_text = box.currentText()
            box.blockSignals(True)
            box.clear()
            box.addItem("请选择对应列", -1)
            for i, name in enumerate(sheet.columns):
                box.addItem(name, i)
            previous = box.findText(previous_text)
            box.setCurrentIndex(previous if previous > 0 else guesses[field] + 1)
            box.blockSignals(False)
        self._set_mapping_enabled(True)
        self._show_source_table()
        hint = header_hint(sheet, ISSUE_ALIASES, guesses)
        self.status_label.setText(
            f"已读取 {len(sheet.rows)} 行、{len(sheet.columns)} 列；当前表头为第 {self.header_row.value()} 行。"
            + ("\n" + hint if hint else "请确认四个字段，再生成问题统计。"))

    def _set_table(self, headers, rows) -> None:
        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(rows))
        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                self.table.setItem(i, j, QTableWidgetItem("" if value is None else str(value)))

    def _show_source_table(self) -> None:
        self._set_table(["Excel 行号", *self.sheet.columns],
                        [(number, *(cell.value for cell in row))
                         for number, row in zip(self.sheet.row_numbers, self.sheet.rows)])

    def _generate_report(self) -> None:
        if self.sheet is None:
            return
        self._invalidate_report()
        try:
            report = build_issues_report(self.sheet,
                {field: box.currentData() for field, box in self.column_boxes.items()},
                products=self.products, all_years=self.scope_box.currentData())
            for chart in (self.overview, *self.charts):
                chart.plot_report(report)
        except (ExcelDataError, ValueError) as error:
            self.status_label.setText("生成失败，请根据提示修正 Excel 或字段选择。")
            QMessageBox.warning(self, "无法生成问题统计", str(error))
            return
        self.report = report
        self.export_button.setEnabled(True)
        self.report_changed.emit()
        self.status_label.setText(
            f"2026年问题 {sum(m.annual_total for m in report.metrics)} 条；"
            f"月度趋势 {len(report.trend_products)} 个山头；密度及关闭率：{report.scope}。\n"
            + "\n".join(report.notes))
        self._set_table(["山头", "2026年问题数", f"总问题数（{report.scope}）", "去重机器数", "问题密度",
                         "关闭数", "关闭率", *(f"2026-{month:02}" for month in report.months)],
                        [(m.product, m.annual_total, m.total, m.machines,
                          "N/A" if m.density is None else f"{m.density:.2f}", m.closed,
                          "N/A" if m.closure_rate is None else f"{m.closure_rate:.1%}", *m.monthly)
                         for m in report.metrics])
        self.tabs.setCurrentIndex(0)

    def _export_report(self) -> None:
        if self.report is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "保存问题统计 PPT", "问题统计.pptx", "PowerPoint (*.pptx)")
        if not path:
            return
        try:
            from excel_visualizer.ppt_service import export_issues_pptx
            result = export_issues_pptx(self.report, path)
        except Exception as error:
            QMessageBox.critical(self, "导出失败", f"{error}\n若目标 PPT 正在打开，请关闭后重试。")
            return
        QMessageBox.information(self, "导出完成", f"已生成 1 页 PPT：四张可编辑图表在上半页，下半页留白。\n{result}")
