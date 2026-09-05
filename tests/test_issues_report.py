"""用已知计数验证第二份 Excel、预览、PPT 内嵌数据和跨步骤失效处理。"""

import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook
from pptx import Presentation
from PyQt6.QtWidgets import QApplication, QMessageBox

from excel_visualizer.data_service import ExcelDataError, NON_OPTICAL, OPTICAL, build_report, guess_columns, read_sheet
from excel_visualizer.issues_service import build_issues_report, guess_issue_columns
from excel_visualizer.issues_charts import issue_charts, series_color
from excel_visualizer.main_window import MainWindow
from excel_visualizer.ppt_service import export_issues_pptx, export_pptx
from test_monthly_report import fixture as monthly_fixture


TODAY = date(2026, 9, 5)


def issues_fixture(path):
    book = Workbook()
    sheet = book.active
    sheet.title = "问题"
    sheet.append(["山头", "创建时间", "产品序列号/机台编码", "服务请求状态"])
    for row in [
        ("XRF", "2026-1-2", "001", "申请关闭"),
        (" xrf ", "2026-1-3", "001", "已结束"),
        ("XRF", datetime(2026, 9, 1), "002", "处理中"),
        ("XRF", "2025-12-31", "003", "已取消"),
        ("XRF", "2026-12-1", "002", "已取消"),
        ("BFI", "2026-2-4 09:30:00", "001", "已取消"),
        ("BFI", "2026-2-4", "001", "已关闭"),
        ("AFM", "2026-3-2", None, "处理中"),
        ("额外山头", "2026-9-1", 123, None),
        ("额外山头", "2026-9-2", 123.0, "已结束"),
        ("XPS", "2025-1-1", "old", "已结束"),
    ]:
        sheet.append(row)
    book.save(path)
    book.close()


class IssuesReportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "issues.xlsx"
        issues_fixture(self.path)

    def report(self, **kwargs):
        sheet = read_sheet(self.path, "问题")
        return build_issues_report(sheet, guess_issue_columns(sheet), today=TODAY, **kwargs)

    def write_rows(self, rows):
        book = Workbook()
        book.active.append(["山头", "创建时间", "产品序列号/机台编码", "服务请求状态"])
        for row in rows:
            book.active.append(row)
        book.active.title = "问题"
        book.save(self.path)
        book.close()

    def test_counts_months_distinct_machines_and_exact_closed_statuses(self):
        report = self.report()
        metrics = {m.product: m for m in report.metrics}
        xrf = metrics["XRF"]
        self.assertEqual((xrf.annual_total, xrf.total, xrf.machines, xrf.closed), (4, 5, 3, 4))
        self.assertEqual(xrf.monthly, (2, 0, 0, 0, 0, 0, 0, 0, 1))
        self.assertAlmostEqual(xrf.density, 5 / 3)
        self.assertEqual(xrf.closure_rate, 0.8)
        self.assertEqual((metrics["BFI"].total, metrics["BFI"].machines, metrics["BFI"].closed), (2, 1, 1))
        self.assertEqual(metrics["额外山头"].machines, 1)
        self.assertEqual(metrics["额外山头"].density, 2)
        self.assertIsNone(metrics["AFM"].density)
        self.assertEqual(metrics["AFM"].closure_rate, 0)
        self.assertEqual(metrics["XRD"].annual_total, 0)
        self.assertIsNone(metrics["XRD"].closure_rate)
        self.assertEqual(metrics["XPS"].monthly, (0,) * 9)
        self.assertIn("XPS", report.trend_products)
        self.assertEqual(len(report.trend_products), 5)
        self.assertEqual(tuple(metrics)[:10], NON_OPTICAL + OPTICAL)
        self.assertTrue(any("后续月份" in note for note in report.notes))
        self.assertTrue(any("缺少机器" in note for note in report.notes))

    def test_optional_year_scope_changes_only_density_closure_population(self):
        report = self.report(all_years=False)
        xrf = report.metrics[0]
        self.assertEqual((xrf.total, xrf.machines, xrf.closed), (4, 2, 3))
        self.assertEqual(xrf.density, 2)
        self.assertEqual(xrf.closure_rate, .75)
        self.assertEqual(xrf.monthly, self.report().metrics[0].monthly)
        self.assertEqual(report.trend_products, self.report().trend_products)

    def test_first_step_extra_names_and_missing_products_preserved(self):
        report = self.report(products=("第一步新增", "XRF"))
        self.assertIn("第一步新增", [m.product for m in report.metrics])
        self.assertEqual(next(m for m in report.metrics if m.product == "第一步新增").annual_total, 0)

    def test_future_year_caps_trend_at_december_and_january_has_one_month(self):
        sheet = read_sheet(self.path, "问题")
        mapping = guess_issue_columns(sheet)
        self.assertEqual(build_issues_report(sheet, mapping, today=date(2027, 2, 1)).months, tuple(range(1, 13)))
        self.assertEqual(build_issues_report(sheet, mapping, today=date(2026, 1, 1)).months, (1,))
        with self.assertRaisesRegex(ExcelDataError, "尚未进入"):
            build_issues_report(sheet, mapping, today=date(2025, 1, 1))

    def test_invalid_dates_and_required_values_report_excel_row(self):
        for created in ("2026-2-30", "2026-13-1", "2026/1/1", None, "坏日期", "=TODAY()"):
            with self.subTest(created=created):
                self.write_rows([("XRF", created, "001", "已结束")])
                with self.assertRaisesRegex(ExcelDataError, "第 2 行"):
                    self.report()
        self.write_rows([(None, "2026-1-1", "001", "已结束")])
        with self.assertRaisesRegex(ExcelDataError, "第 2 行"):
            self.report()

    def test_distinct_codes_preserve_text_leading_zeros_and_machine_scope(self):
        self.write_rows([("XRF", "2026-1-1", "001", "已结束"),
                         ("XRF", "2026-1-1", 1, "已结束"),
                         ("XRF", "2026-1-1", " 001 ", "已结束"),
                         ("BFI", "2026-1-1", "001", "已结束")])
        report = self.report()
        self.assertEqual(report.metrics[0].machines, 2)
        self.assertEqual(report.metrics[4].machines, 1)

    def test_missing_mapping_duplicate_mapping_empty_and_no_2026(self):
        sheet = read_sheet(self.path, "问题")
        mapping = guess_issue_columns(sheet)
        self.assertEqual(mapping, dict(product=0, created=1, machine=2, status=3))
        for bad_mapping in ({}, {**mapping, "machine": 1}):
            with self.assertRaises(ExcelDataError):
                build_issues_report(sheet, bad_mapping, today=TODAY)
        self.write_rows([])
        with self.assertRaisesRegex(ExcelDataError, "没有符合"):
            self.report()
        self.write_rows([("XRF", "2025-1-1", "001", "已结束")])
        report = self.report()
        self.assertEqual(report.metrics[0].annual_total, 0)
        self.assertEqual(report.metrics[0].monthly, (0,) * 9)
        # 无 2026 问题时仍可导出年度零值和全年的后两图。
        output = export_issues_pptx(report, Path(self.temp.name) / "old.pptx")
        self.assertEqual(len(Presentation(output).slides), 1)

    def test_ppt_has_four_editable_charts_only_in_upper_half(self):
        report = self.report()
        path = export_issues_pptx(report, Path(self.temp.name) / "issues")
        self.assertEqual(path.suffix, ".pptx")
        ppt = Presentation(path)
        self.assertEqual(len(ppt.slides), 1)
        slide = ppt.slides[0]
        shapes = [s for s in slide.shapes if s.has_chart]
        self.assertEqual(len(shapes), 4)
        for shape in slide.shapes:
            self.assertLessEqual(shape.top + shape.height, ppt.slide_height // 2)
            self.assertLessEqual(shape.left + shape.width, ppt.slide_width)
        self.assertEqual(shapes[0].top, shapes[1].top)
        self.assertEqual(shapes[2].top, shapes[3].top)
        self.assertEqual(shapes[0].left, shapes[2].left)
        self.assertLess(shapes[0].left + shapes[0].width, shapes[1].left)
        self.assertLess(shapes[0].top + shapes[0].height, shapes[2].top)
        self.assertEqual(tuple(shapes[0].chart.series[0].values), tuple(m.annual_total for m in report.metrics))
        trend = shapes[1].chart
        self.assertEqual([s.name for s in trend.series], list(report.trend_products))
        self.assertEqual([c.label for c in trend.plots[0].categories], [f"{m}月" for m in report.months])
        self.assertEqual(len({str(s.format.line.color.rgb) for s in trend.series}), len(report.trend_products))
        for shape, spec in zip(shapes, issue_charts(report)):
            from openpyxl import load_workbook
            book = load_workbook(BytesIO(shape.chart.part.chart_workbook.xlsx_part.blob), data_only=True)
            rows = list(book.active.values)
            for i, (_, expected) in enumerate(spec.series, 1):
                for row, value in zip(rows[1:], expected):
                    if value is None:
                        self.assertIsNone(row[i])
                    else:
                        self.assertAlmostEqual(row[i], value)
            book.close()

    def test_combined_report_has_three_original_pages_then_four_chart_page(self):
        monthly_path = Path(self.temp.name) / "monthly.xlsx"
        monthly_fixture(monthly_path)
        sheet = read_sheet(monthly_path, "月报", 2)
        monthly = build_report(sheet, guess_columns(sheet))
        result = export_pptx(monthly, Path(self.temp.name) / "combined.pptx", issues_report=self.report())
        ppt = Presentation(result)
        self.assertEqual(len(ppt.slides), 4)
        self.assertEqual([sum(s.has_chart for s in slide.shapes) for slide in ppt.slides], [1, 1, 1, 4])
        self.assertIn("1 / 4", [s.text for s in ppt.slides[0].shapes if s.has_text_frame])

    def test_failed_issue_save_preserves_previous_file(self):
        output = Path(self.temp.name) / "existing.pptx"
        output.write_bytes(b"existing")
        with patch("excel_visualizer.ppt_service.os.replace", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError):
                export_issues_pptx(self.report(), output)
        self.assertEqual(output.read_bytes(), b"existing")
        self.assertEqual(len(list(output.parent.glob("*.pptx"))), 1)

    def test_gui_independent_import_export_and_stale_combined_report(self):
        window = MainWindow()
        self.addCleanup(window.close)
        page = window.issues_page
        with patch.object(QMessageBox, "warning") as warning:
            page.open_excel(self.path)
            page._generate_report()
        self.assertFalse(warning.called, str(warning.call_args))
        self.assertIsNotNone(page.report)
        self.assertFalse(window.combined_export_button.isEnabled())
        self.assertTrue(page.export_button.isEnabled())
        self.assertEqual(len(page.overview.figure.axes), 4)
        self.assertEqual(len(page.charts[1].figure.axes[0].lines), 5)
        page.overview.draw()
        old_issues = page.report
        monthly_path = Path(self.temp.name) / "monthly.xlsx"
        monthly_fixture(monthly_path)
        window.header_row.setValue(2)
        with patch.object(QMessageBox, "warning"):
            window.open_excel(monthly_path)
        window.sheet_box.setCurrentText("月报")
        window._load_sheet()
        window._generate_report()
        # 第一份表新增“其他产品”，旧问题报告必须重新生成以补齐横轴。
        self.assertIsNone(page.report)
        page._generate_report()
        self.assertIsNot(page.report, old_issues)
        self.assertTrue(window.combined_export_button.isEnabled())
        self.assertIn("其他产品", [m.product for m in page.report.metrics])
        output = Path(self.temp.name) / "gui.pptx"
        with patch("excel_visualizer.main_window.QFileDialog.getSaveFileName", return_value=(str(output), "")), \
             patch.object(QMessageBox, "information"):
            window.combined_export_button.click()
        self.assertEqual(len(Presentation(output).slides), 4)
        with patch("excel_visualizer.issues_page.QFileDialog.getSaveFileName", return_value=(str(output), "")), \
             patch.object(QMessageBox, "information"):
            page.export_button.click()
        self.assertEqual(len(Presentation(output).slides), 1)
        monthly_report = window.report
        page.scope_box.setCurrentIndex(1)
        self.assertFalse(window.combined_export_button.isEnabled())
        self.assertIs(window.report, monthly_report)
        self.assertFalse(page.export_button.isEnabled())
        page._generate_report()
        self.assertTrue(window.combined_export_button.isEnabled())
        window.uptime_mode.setCurrentIndex(1)
        self.assertFalse(window.combined_export_button.isEnabled())
        self.assertTrue(page.export_button.isEnabled())
        page.header_row.setValue(2)
        self.assertIsNone(page.sheet)
        self.assertFalse(page.plot_button.isEnabled())

    def test_extra_series_get_distinct_colors(self):
        self.assertEqual(len({series_color(i) for i in range(30)}), 30)


if __name__ == "__main__":
    unittest.main()
