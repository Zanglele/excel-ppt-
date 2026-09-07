"""三步月报工作流；已实现前两步，第三步预留。"""

from dataclasses import replace
from datetime import date
from pathlib import Path

from PyQt6.QtGui import QColor, QFont, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QColorDialog, QComboBox, QFileDialog, QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPushButton, QSpinBox, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from excel_visualizer.chart_widget import MonthlyChartWidget
from excel_visualizer.chart_style import ChartColors, DEFAULT_COLORS
from excel_visualizer.issues_page import IssuesPage
from excel_visualizer.data_service import (
    ALIASES, ExcelDataError, FIELDS, FIELD_LABELS, GROUP_TITLES, MonthlyReport, SheetData,
    build_report, guess_columns, header_hint, list_sheets, read_sheet,
)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setFont(QFont("Microsoft YaHei", 10))
        self.file_path: Path | None = None
        self.sheet: SheetData | None = None
        self.report: MonthlyReport | None = None
        self.chart_colors = DEFAULT_COLORS
        self.setWindowTitle("Excel 月报图表与 PPT 工具")
        self.resize(1280, 900)
        self.open_button = QPushButton("导入第一份 Excel")
        self.file_label = QLabel("尚未选择文件")
        self.sheet_box = QComboBox()
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 1048576)
        self.reload_button = QPushButton("读取工作表 / 刷新")
        self.column_boxes = {field: QComboBox() for field in FIELDS}
        self.uptime_mode = QComboBox()
        self.uptime_mode.addItem("自动识别（0～1比例；大于1百分制）", "auto")
        self.uptime_mode.addItem("按 Excel 格式（普通数值 98 = 98%）", "excel")
        self.uptime_mode.addItem("普通数值全部按百分制（98 = 98%）", "points")
        self.uptime_mode.addItem("普通数值全部按小数比例（0.98 = 98%）", "fraction")
        self.period_edit = QLineEdit(date.today().strftime("%Y-%m"))
        self.period_edit.setMaxLength(12)
        self.period_edit.setReadOnly(True)
        self.year_box = QSpinBox()
        self.year_box.setRange(1900, 9999)
        self.year_box.setValue(date.today().year)
        self.month_box = QSpinBox()
        self.month_box.setRange(1, 12)
        self.month_box.setValue(date.today().month)
        self.start_month_box = QSpinBox()
        self.start_month_box.setRange(1, 12)
        self.line_color_button = QPushButton()
        self.bar_color_button = QPushButton()
        self.reset_colors_button = QPushButton("恢复默认配色")
        self.plot_button = QPushButton("生成三张月报图表")
        self.export_button = QPushButton("导出第一步 PPT（3 页）")
        self.combined_export_button = QPushButton("合并导出前两步 PPT（4 页）")
        self.combined_export_button.setEnabled(False)
        self.status_label = QLabel("请选择每月的 .xlsx 文件。黄色客户单元格表示未入质保。")
        self.status_label.setWordWrap(True)
        self.tabs = QTabWidget()
        self.charts = [MonthlyChartWidget() for _ in GROUP_TITLES]
        for title, chart in zip(GROUP_TITLES, self.charts):
            self.tabs.addTab(chart, title)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.tabs.addTab(self.table, "数据核对")
        self.workflow_tabs = QTabWidget()
        self.monthly_page = QWidget()
        self.issues_page = IssuesPage()
        self.unresolved_page = QWidget()
        self.workflow_tabs.addTab(self.monthly_page, "1. Uptime 与跑货量（3 页）")
        self.workflow_tabs.addTab(self.issues_page, "2. 问题统计（1 页）")
        self.workflow_tabs.addTab(self.unresolved_page, "3. 未解决的问题统计（预留）")
        self.workflow_tabs.setTabToolTip(1, "第二份 Excel，1 页 PPT，上半页 4 张图表，下半页留白")
        self.workflow_tabs.setTabToolTip(2, "预留：第三份 Excel，1 页 PPT，3 张图表")
        self.issues_page.set_period(self.year_box.value(), self.start_month_box.value(), self.month_box.value())
        self._build_layout()
        for box in (self.year_box, self.month_box, self.start_month_box):
            box.valueChanged.connect(self._period_changed)
        self._update_color_buttons()
        self.line_color_button.clicked.connect(lambda: self._choose_color("line"))
        self.bar_color_button.clicked.connect(lambda: self._choose_color("bar"))
        self.reset_colors_button.clicked.connect(lambda: self.set_chart_colors(DEFAULT_COLORS))
        self.open_button.clicked.connect(self._select_excel_file)
        self.reload_button.clicked.connect(self._load_sheet)
        self.sheet_box.currentIndexChanged.connect(self._source_changed)
        self.header_row.valueChanged.connect(self._source_changed)
        for box in self.column_boxes.values():
            box.currentIndexChanged.connect(self._mapping_changed)
        self.uptime_mode.currentIndexChanged.connect(self._mapping_changed)
        self.plot_button.clicked.connect(self._generate_report)
        self.export_button.clicked.connect(self._export_report)
        self.combined_export_button.clicked.connect(self._export_combined_report)
        self.issues_page.report_changed.connect(self._update_combined_export)
        self.sheet_box.setEnabled(False)
        self.header_row.setEnabled(False)
        self.reload_button.setEnabled(False)
        self._set_mapping_enabled(False)
        self.export_button.setEnabled(False)

    def _build_layout(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        workflow_hint = QLabel(
            "月报工作流：三份 Excel，目标五页 PPT。当前可用前两步，可合并导出四页；第三步待开发。")
        workflow_hint.setWordWrap(True)
        root.addWidget(workflow_hint)
        appearance = QGroupBox("汇报设置")
        appearance_layout = QVBoxLayout(appearance)
        color_row = QHBoxLayout()
        color_row.addWidget(QLabel("汇报月份："))
        color_row.addWidget(self.year_box)
        color_row.addWidget(QLabel("年"))
        color_row.addWidget(self.month_box)
        color_row.addWidget(QLabel("月；趋势起始月："))
        color_row.addWidget(self.start_month_box)
        color_row.addWidget(self.bar_color_button)
        color_row.addWidget(self.line_color_button)
        color_row.addWidget(self.reset_colors_button)
        color_row.addStretch()
        appearance_layout.addLayout(color_row)
        color_hint = QLabel("点击色块按钮选择颜色，统一应用于第一步的三张预览图和导出 PPT；修改后即时生效。")
        color_hint.setWordWrap(True)
        appearance_layout.addWidget(color_hint)
        appearance_layout.addWidget(self.combined_export_button)
        root.addWidget(appearance)
        root.addWidget(self.workflow_tabs, 1)
        monthly_layout = QVBoxLayout(self.monthly_page)
        files = QHBoxLayout()
        files.addWidget(self.open_button)
        files.addWidget(self.file_label, 1)
        monthly_layout.addLayout(files)
        source = QHBoxLayout()
        source.addWidget(QLabel("工作表："))
        source.addWidget(self.sheet_box, 1)
        source.addWidget(QLabel("表头行："))
        source.addWidget(self.header_row)
        source.addWidget(self.reload_button)
        monthly_layout.addLayout(source)
        fields = QGridLayout()
        for index, (field, label) in enumerate(zip(FIELDS, FIELD_LABELS)):
            fields.addWidget(QLabel(f"{label}："), index // 2, (index % 2) * 2)
            fields.addWidget(self.column_boxes[field], index // 2, (index % 2) * 2 + 1)
        fields.addWidget(QLabel("Uptime 格式："), 3, 0)
        fields.addWidget(self.uptime_mode, 3, 1)
        monthly_layout.addLayout(fields)
        hint = QLabel("质保按客户名字单元格的黄色填充识别；左轴为 Uptime 折线，右轴为跑货量柱形。")
        hint.setWordWrap(True)
        monthly_layout.addWidget(hint)
        actions = QHBoxLayout()
        actions.addWidget(self.plot_button)
        actions.addWidget(self.export_button)
        actions.addStretch()
        monthly_layout.addLayout(actions)
        from excel_visualizer.details import add_chart_actions
        self.chart_buttons = add_chart_actions(monthly_layout, self, 3, self._generate_report, monthly=True)
        monthly_layout.addWidget(self.status_label)
        monthly_layout.addWidget(self.tabs, 1)
        self.setCentralWidget(central)

    def _period_changed(self) -> None:
        self.period_edit.setText(f"{self.year_box.value()}-{self.month_box.value():02}")
        self._mapping_changed()
        self.issues_page.set_period(self.year_box.value(), self.start_month_box.value(), self.month_box.value())

    def _update_color_buttons(self) -> None:
        for button, label, color in (
            (self.bar_color_button, "柱状图 / 跑货量", self.chart_colors.bar),
            (self.line_color_button, "折线图 / Uptime", self.chart_colors.line),
        ):
            swatch = QPixmap(20, 20)
            swatch.fill(QColor(color))
            button.setIcon(QIcon(swatch))
            button.setText(f"{label}  {color}")
            button.setToolTip("点击选择颜色，也可在颜色对话框中输入 HTML 十六进制色号。")
        self.reset_colors_button.setEnabled(self.chart_colors != DEFAULT_COLORS)

    def _choose_color(self, field: str) -> None:
        title = "选择 Uptime 折线颜色" if field == "line" else "选择跑货量柱状图颜色"
        color = QColorDialog.getColor(QColor(getattr(self.chart_colors, field)), self, title,
                                     QColorDialog.ColorDialogOption.DontUseNativeDialog)
        if color.isValid():
            self.set_chart_colors(replace(self.chart_colors, **{field: color.name()}))

    def set_chart_colors(self, colors: ChartColors) -> None:
        """改变外观无需重新读取或计算数据，也不改变当前核对页签。"""
        if colors == self.chart_colors:
            return
        self.chart_colors = colors
        self._update_color_buttons()
        if self.report is not None:
            for chart, group in zip(self.charts, self.report.groups):
                chart.plot_group(group, self.chart_colors)

    def _set_mapping_enabled(self, enabled: bool) -> None:
        for box in self.column_boxes.values():
            box.setEnabled(enabled)
        self.uptime_mode.setEnabled(enabled)
        self.plot_button.setEnabled(enabled)
        for button in self.chart_buttons:
            button.setEnabled(enabled)

    def _invalidate_report(self) -> None:
        self.report = None
        self.export_button.setEnabled(False)
        self._update_combined_export()
        for chart in self.charts:
            chart.show_placeholder()

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
            self.status_label.setText("字段或数值格式已改变，请重新生成三张图表。")

    def _select_excel_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "选择月报 Excel", "", "Excel 文件 (*.xlsx)")
        if file_path:
            self.open_excel(file_path)

    def open_excel(self, file_path: str | Path) -> None:
        try:
            names = list_sheets(file_path)
        except ExcelDataError as error:
            QMessageBox.critical(self, "读取失败", str(error))
            return
        self.file_path = Path(file_path)
        self.file_label.setText(str(self.file_path))
        self.sheet_box.blockSignals(True)
        self.sheet_box.clear()
        self.sheet_box.addItems(names)
        self.sheet_box.blockSignals(False)
        self.sheet_box.setEnabled(True)
        self.header_row.setEnabled(True)
        self.reload_button.setEnabled(True)
        self._load_sheet()

    def _load_sheet(self) -> None:
        if self.file_path is None:
            return
        self._source_changed()
        try:
            sheet = read_sheet(self.file_path, self.sheet_box.currentText(), self.header_row.value())
        except ExcelDataError as error:
            self.status_label.setText("读取失败，请检查工作表和表头行后重试。")
            QMessageBox.warning(self, "读取失败", str(error))
            return
        self.sheet = sheet
        guesses = guess_columns(sheet)
        for field, box in self.column_boxes.items():
            old_name = box.currentText()
            box.blockSignals(True)
            box.clear()
            box.addItem("请选择对应列", -1)
            for index, name in enumerate(sheet.columns):
                box.addItem(name, index)
            previous = box.findText(old_name)
            box.setCurrentIndex(previous if previous > 0 else guesses[field] + 1)
            box.blockSignals(False)
        self._set_mapping_enabled(True)
        self._show_source_table()
        hint = header_hint(sheet, ALIASES, guesses)
        self.status_label.setText(
            f"已读取 {len(sheet.rows)} 行、{len(sheet.columns)} 列；当前表头为第 {self.header_row.value()} 行。"
            + ("\n" + hint if hint else "请确认五个字段和 Uptime 格式，再生成图表。"))

    def _show_source_table(self) -> None:
        if self.sheet is None:
            return
        self.table.clear()
        self.table.setColumnCount(len(self.sheet.columns) + 1)
        self.table.setHorizontalHeaderLabels(["Excel 行号", *self.sheet.columns])
        self.table.setRowCount(len(self.sheet.rows))
        for index, (number, row) in enumerate(zip(self.sheet.row_numbers, self.sheet.rows)):
            self.table.setItem(index, 0, QTableWidgetItem(str(number)))
            for column, cell in enumerate(row, 1):
                item = QTableWidgetItem("" if cell.value is None else str(cell.value))
                item.setToolTip(f"{cell.coordinate}；Excel 数字格式：{cell.number_format}")
                if cell.rgb:
                    item.setBackground(QColor(f"#{cell.rgb}"))
                    item.setForeground(QColor("#000000"))
                self.table.setItem(index, column, item)

    def _generate_report(self) -> None:
        if self.sheet is None:
            return
        self._invalidate_report()
        try:
            if self.start_month_box.value() > self.month_box.value():
                raise ExcelDataError("趋势起始月不能大于结束月。")
            report = build_report(self.sheet, {field: box.currentData() for field, box in self.column_boxes.items()},
                                  self.uptime_mode.currentData())
            from excel_visualizer.details import monthly_titles
            report = replace(report, groups=tuple(replace(group, title=title) for group, title in
                             zip(report.groups, monthly_titles(self.year_box.value(), self.month_box.value()))))
            empty = [g.title for g in report.groups if not g.records or not any(r.uptime is not None or r.volume is not None for r in g.records)]
            if empty:
                raise ExcelDataError("当前时间范围无有效数据：" + "、".join(empty) + "\n" + "\n".join(report.notes))
            if report.notes:
                QMessageBox.warning(self, "数据预警（缺失值保持为空）", "\n".join(report.notes))
            for chart, group in zip(self.charts, report.groups):
                chart.plot_group(group, self.chart_colors)
        except (ExcelDataError, ValueError) as error:
            self.status_label.setText("生成失败，请根据提示修正 Excel 或字段选择。")
            QMessageBox.warning(self, "无法生成月报", str(error))
            return
        self.report = report
        self.issues_page.set_products(tuple(record.product for group in report.groups for record in group.records))
        self._update_combined_export()
        self.export_button.setEnabled(True)
        summary = "；".join(f"{group.title} {len(group.records)} 条" for group in report.groups)
        self.status_label.setText(summary + ("\n" + "\n".join(report.notes) if report.notes else ""))
        self._show_report_table(report)
        self.tabs.setCurrentIndex(0)

    def _show_report_table(self, report: MonthlyReport) -> None:
        self.table.clear()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Excel 行号", "分类", "产品", "客户名字", "机台编码", "Uptime", "跑货量"])
        self.table.setRowCount(sum(len(group.records) for group in report.groups))
        row = 0
        for group_index, group in enumerate(report.groups):
            for record in group.records:
                values = (record.source_row, group.title, record.product, record.customer, record.machine,
                          "缺失" if record.uptime is None else f"{record.uptime:g}%",
                          "缺失" if record.volume is None else f"{record.volume:,.2f}".rstrip("0").rstrip("."))
                for column, value in enumerate(values):
                    item = QTableWidgetItem(str(value))
                    if group_index == 2:
                        item.setBackground(QColor("#FFF2CC"))
                        item.setForeground(QColor("#000000"))
                    self.table.setItem(row, column, item)
                row += 1

    def _update_combined_export(self) -> None:
        self.combined_export_button.setEnabled(self.report is not None and self.issues_page.report is not None)

    def _export_combined_report(self) -> None:
        self._export_report(combined=True)

    def _export_report(self, checked: bool = False, *, combined: bool = False) -> None:
        if self.report is None:
            return
        if combined and self.issues_page.report is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "保存月报 PPT", "客户月报.pptx", "PowerPoint 演示文稿 (*.pptx)")
        if not path:
            return
        try:
            from excel_visualizer.ppt_service import export_pptx
            result = export_pptx(self.report, path, self.period_edit.text(), self.chart_colors,
                                 issues_report=self.issues_page.report if combined else None)
        except ImportError:
            QMessageBox.critical(self, "缺少导出依赖", "请在当前 Python 环境执行：python -m pip install -r requirements.txt")
            return
        except Exception as error:
            QMessageBox.critical(self, "导出失败", f"{error}\n若目标 PPT 正在打开，请关闭后重试。")
            return
        message = "已生成前两步的四页 PPT" if combined else "已生成第一步的三页 PPT"
        QMessageBox.information(self, "导出完成", f"{message}，图表可在 PowerPoint 中编辑数据：\n{result}")
