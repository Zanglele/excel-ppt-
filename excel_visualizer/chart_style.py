"""预览与 PPT 共用的图表配色，不依赖 Qt 或 PowerPoint。"""

from dataclasses import dataclass
import re


@dataclass(frozen=True)
class ChartColors:
    line: str = "#D45A32"
    bar: str = "#4F81BD"

    def __post_init__(self) -> None:
        for field in ("line", "bar"):
            color = getattr(self, field)
            if not isinstance(color, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
                raise ValueError("图表颜色必须为 #RRGGBB 格式。")
            object.__setattr__(self, field, color.upper())


DEFAULT_COLORS = ChartColors()
