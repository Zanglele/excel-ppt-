"""生成明确标记的模拟 Excel、PPT 和离屏界面截图，供版式人工核验。"""

import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFontDatabase

from excel_visualizer.data_service import NON_OPTICAL, OPTICAL, build_report, guess_columns, read_sheet
from excel_visualizer.issues_service import build_issues_report, guess_issue_columns
from excel_visualizer.main_window import MainWindow
from excel_visualizer.ppt_service import export_issues_pptx, export_pptx
from test_issues_report import issues_fixture


def main():
    folder = Path(__file__).resolve().parents[1] / "outputs" / "问题统计功能验证"
    folder.mkdir(parents=True, exist_ok=True)
    issues_path = folder / "问题统计模拟数据.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "问题统计（模拟）"
    sheet.append(["山头", "创建时间", "产品序列号/机台编码", "服务请求状态"])
    for i, name in enumerate(NON_OPTICAL + OPTICAL):
        for month in range(1, 10):
            for j in range((i * 3 + month * (i % 4 + 1)) % 20 + 2):
                sheet.append([name, f"2026-{month}-{j + 1}", f"{name}-{j % (i + 2):03}",
                              ("申请关闭", "已结束", "已取消", "处理中", "处理中")[(j + i) % 5]])
    for col in ("A", "B", "C", "D"):
        sheet.column_dimensions[col].width = 28
    book.save(issues_path)
    book.close()
    monthly_path = folder / "第一步模拟数据.xlsx"
    book = Workbook()
    sheet = book.active
    sheet.title = "月报（模拟）"
    sheet.append(["客户名字", "山头", "uptime", "跑货量"])
    for i, name in enumerate(NON_OPTICAL + OPTICAL):
        sheet.append([f"模拟客户{i + 1}", name, 95 + i / 3, 1000 + i * 123])
    book.save(monthly_path)
    book.close()
    issue_sheet = read_sheet(issues_path, "问题统计（模拟）")
    report = build_issues_report(issue_sheet, guess_issue_columns(issue_sheet), today=date(2026, 9, 5))
    export_issues_pptx(report, folder / "问题统计示例（模拟数据）.pptx")
    monthly_sheet = read_sheet(monthly_path, "月报（模拟）")
    monthly = build_report(monthly_sheet, guess_columns(monthly_sheet))
    export_pptx(monthly, folder / "前两步合并示例（模拟数据）.pptx", "2026-09", issues_report=report)
    edge_path = folder / "边界情况模拟数据.xlsx"
    issues_fixture(edge_path)
    edge_sheet = read_sheet(edge_path, "问题")
    edge_report = build_issues_report(edge_sheet, guess_issue_columns(edge_sheet), today=date(2026, 9, 5))
    export_issues_pptx(edge_report, folder / "边界情况示例（模拟数据）.pptx")
    app = QApplication.instance() or QApplication([])
    # 离屏 Qt 平台不自动加载系统中文字体，显式加载仅用于截图验证。
    QFontDatabase.addApplicationFont("C:/Windows/Fonts/msyh.ttc")
    window = MainWindow()
    window.show()
    app.processEvents()
    window.open_excel(monthly_path)
    window._generate_report()
    app.processEvents()
    window.workflow_tabs.setCurrentIndex(1)
    app.processEvents()
    window.issues_page.open_excel(issues_path)
    window.issues_page._generate_report()
    app.processEvents()
    window.grab().save(str(folder / "第二步界面.png"))
    window.resize(1024, 768)
    app.processEvents()
    window.grab().save(str(folder / "第二步界面-1024.png"))
    window.close()
    print("Generated simulated workbooks, PPTX files, and Qt screenshots.")


if __name__ == "__main__":
    main()
