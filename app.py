"""桌面应用启动入口。"""

import sys

from PyQt6.QtWidgets import QApplication

from excel_visualizer.main_window import MainWindow


def main() -> int:
    """创建 Qt 应用并进入事件循环。"""
    application = QApplication(sys.argv)
    application.setApplicationName("Excel 月报图表与 PPT 工具")

    window = MainWindow()
    window.show()

    return application.exec()


if __name__ == "__main__":
    sys.exit(main())
