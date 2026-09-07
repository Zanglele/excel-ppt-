"""问题统计的图表定义；Qt 预览和可编辑 PPT 共用同一份数据和配色。"""

from colorsys import hsv_to_rgb
from dataclasses import dataclass

from excel_visualizer.issues_service import IssuesReport


@dataclass(frozen=True)
class IssueChart:
    title: str
    categories: tuple[str, ...]
    series: tuple[tuple[str, tuple[float | None, ...]], ...]
    ylabel: str
    kind: str = "bar"
    number_format: str = "0"


def series_color(index: int) -> str:
    palette = ("#5B9BD5", "#D45A32", "#548235", "#8064A2", "#008C95",
               "#BF9000", "#C04B87", "#404E78", "#88734B", "#747474")
    if index < len(palette):
        return palette[index]
    # 额外山头仍各有颜色，不循环复用固定十色。
    rgb = hsv_to_rgb((index * 0.61803398875) % 1, 0.68, 0.72)
    return "#" + "".join(f"{round(channel * 255):02X}" for channel in rgb)


def issue_charts(report: IssuesReport) -> tuple[IssueChart, ...]:
    names = tuple(item.product for item in report.metrics)
    by_name = {item.product: item for item in report.metrics}
    return (
        IssueChart(f"{report.year}年各山头总问题数", names,
                   (("总问题数", tuple(item.annual_total for item in report.metrics)),), "问题数"),
        IssueChart(f"{report.year}年各山头月度问题出现次数及其变化趋势",
                   tuple(f"{month}月" for month in report.months),
                   tuple((name, by_name[name].monthly) for name in report.trend_products),
                   "问题数", "line"),
        IssueChart(f"{report.year}年各个山头的问题密度", names,
                   (("问题密度", tuple(item.density for item in report.metrics)),),
                   "问题数 / 台", number_format="0.00"),
        IssueChart(f"{report.year}年各山头关闭率", names,
                   (("关闭率", tuple(item.closure_rate for item in report.metrics)),),
                   "关闭率", number_format="0.0%"),
    )
