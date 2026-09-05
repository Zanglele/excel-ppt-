"""问题统计四图预览，可切换到独立图表查看细节。"""

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator, PercentFormatter

from excel_visualizer.issues_charts import issue_charts, series_color
from excel_visualizer.issues_service import IssuesReport


class IssuesChartWidget(FigureCanvasQTAgg):
    def __init__(self, chart_index: int | None = None) -> None:
        self.figure = Figure(figsize=(12, 6), layout="constrained")
        self.chart_index = chart_index
        super().__init__(self.figure)
        self.setMinimumSize(600, 300)
        self.show_placeholder()

    def show_placeholder(self) -> None:
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.text(0.5, 0.5, "导入第二份 Excel 并生成问题统计", ha="center", va="center")
        ax.set_axis_off()
        self.draw_idle()

    def plot_report(self, report: IssuesReport) -> None:
        self.figure.clear()
        specs = issue_charts(report)
        selected = specs if self.chart_index is None else (specs[self.chart_index],)
        for index, spec in enumerate(selected):
            ax = self.figure.add_subplot(2, 2, index + 1) if len(selected) == 4 else self.figure.add_subplot(111)
            positions = list(range(len(spec.categories)))
            if spec.kind == "line":
                for i, (name, values) in enumerate(spec.series):
                    ax.plot(positions, values, label=name, color=series_color(i), marker="o",
                            markersize=3, linewidth=1.5)
                ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18),
                          ncol=min(5, len(spec.series)), fontsize=7, frameon=False)
            else:
                values = spec.series[0][1]
                bars = ax.bar(positions, [v if v is not None else 0 for v in values], color=series_color(0))
                for bar, value in zip(bars, values):
                    label = ("N/A" if value is None else f"{value:.1%}" if spec.number_format.endswith("%")
                             else f"{value:.2f}" if spec.number_format == "0.00" else str(value))
                    ax.annotate(label, (bar.get_x() + bar.get_width() / 2, bar.get_height()),
                                xytext=(0, 3), textcoords="offset points", ha="center", fontsize=7)
            ax.set_title(spec.title, fontsize=11)
            ax.set_ylabel(spec.ylabel, fontsize=9)
            ax.set_xticks(positions, spec.categories, fontsize=8,
                          rotation=30 if len(positions) > 12 else 0)
            ax.tick_params(axis="y", labelsize=8)
            ax.set_ylim(bottom=0)
            ax.margins(y=0.18)
            if spec.number_format.endswith("%"):
                ax.set_ylim(0, 1.15)
                ax.yaxis.set_major_formatter(PercentFormatter(1))
            elif spec.number_format == "0":
                ax.yaxis.set_major_locator(MaxNLocator(integer=True))
            ax.grid(axis="y", alpha=0.2)
            ax.set_axisbelow(True)
        self.draw_idle()
