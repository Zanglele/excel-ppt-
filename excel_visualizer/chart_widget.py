"""月报双轴图的 Qt 预览。"""

from matplotlib import rcParams
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from matplotlib.ticker import PercentFormatter, StrMethodFormatter

from excel_visualizer.data_service import ReportGroup
from excel_visualizer.chart_style import ChartColors, DEFAULT_COLORS


rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
rcParams["axes.unicode_minus"] = False


class MonthlyChartWidget(FigureCanvasQTAgg):
    def __init__(self) -> None:
        self.figure = Figure(figsize=(12, 6), layout="constrained")
        super().__init__(self.figure)
        self.show_placeholder()

    def show_placeholder(self, message: str = "导入 Excel 并确认字段后，生成三张月报图表") -> None:
        self.figure.clear()
        self.axes = self.figure.add_subplot(111)
        self.axes.text(0.5, 0.5, message, ha="center", va="center", transform=self.axes.transAxes)
        self.axes.set_axis_off()
        self.draw_idle()

    def plot_group(self, group: ReportGroup, colors: ChartColors = DEFAULT_COLORS) -> None:
        if not group.records:
            self.show_placeholder(f"{group.title}\n本月无符合条件的数据")
            return
        self.figure.clear()
        self.axes = self.figure.add_subplot(111)
        volume_axes = self.axes.twinx()
        positions = list(range(len(group.records)))
        bars = volume_axes.bar(positions, [r.volume if r.volume is not None else float("nan") for r in group.records], width=0.58,
                               color=colors.bar, label="跑货量")
        self.axes.set_zorder(volume_axes.get_zorder() + 1)
        self.axes.patch.set_visible(False)
        line, = self.axes.plot(positions, [r.uptime for r in group.records], color=colors.line,
                               marker="o", linewidth=2, markersize=4, label="Uptime")
        self.axes.set_ylim(0, 105)
        self.axes.set_yticks(range(0, 101, 20))
        self.axes.yaxis.set_major_formatter(PercentFormatter(xmax=100))
        self.axes.set_ylabel("Uptime (%)", color=colors.line)
        volume_axes.set_ylabel("跑货量", color=colors.bar)
        volume_axes.set_ylim(0, max(1, max((r.volume for r in group.records if r.volume is not None), default=0) * 1.18))
        volume_axes.yaxis.set_major_formatter(StrMethodFormatter("{x:,.0f}"))
        self.axes.set_xlim(-0.65, len(positions) - 0.35)
        self.axes.set_xticks(positions, [r.label for r in group.records],
                             rotation=45 if len(positions) > 12 else 0,
                             ha="right" if len(positions) > 12 else "center",
                             fontsize=8 if len(positions) > 20 else 10)
        self.axes.set_xlabel("机台编码 / 山头 / 客户")
        self.axes.set_title(group.title, pad=38, fontname="Microsoft YaHei", fontsize=16, fontweight="bold")
        self.axes.grid(axis="y", linestyle="--", alpha=0.25)
        self.axes.legend([line, bars], ["Uptime（左轴）", "跑货量（右轴）"],
                         loc="lower center", bbox_to_anchor=(0.5, 1.01), ncol=2, frameon=False)
        self.draw_idle()
