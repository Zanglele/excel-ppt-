"""第二份 Excel 的问题次数、月度趋势、机器去重和关闭率统计。"""

from dataclasses import dataclass
from datetime import date, datetime
import re
import pandas as pd

from excel_visualizer.data_service import ExcelDataError, NON_OPTICAL, OPTICAL, SheetData, _text


ISSUE_FIELDS = ("product", "created", "machine", "status")
ISSUE_LABELS = ("山头", "创建时间", "产品序列号/机台编码", "服务请求状态")
ISSUE_ALIASES = {
    "product": ("山头", "山头/产品", "产品", "产品名称"),
    "created": ("创建时间", "创建日期"),
    "machine": ("产品序列号/机台编码", "产品序列号／机台编码", "产品序列号", "机台编码"),
    "status": ("服务请求状态",),
}
CLOSED_STATUSES = frozenset(("申请关闭", "已结束", "已取消"))


@dataclass(frozen=True)
class IssueMetrics:
    product: str
    annual_total: int
    monthly: tuple[int, ...]
    total: int
    machines: int
    closed: int

    @property
    def density(self) -> float | None:
        return self.total / self.machines if self.machines else None

    @property
    def closure_rate(self) -> float | None:
        return self.closed / self.total if self.total else None


@dataclass(frozen=True)
class IssuesReport:
    metrics: tuple[IssueMetrics, ...]
    months: tuple[int, ...]
    trend_products: tuple[str, ...]
    as_of: date
    source: str
    sheet_name: str
    notes: tuple[str, ...]

    year: int

    @property
    def scope(self) -> str:
        return f"{self.year}年"


def guess_issue_columns(sheet: SheetData) -> dict[str, int]:
    names = [re.sub(r"\s+", "", name.rsplit(" [", 1)[0]).casefold() for name in sheet.columns]
    return {field: next((i for i, name in enumerate(names)
                        if name in {alias.casefold() for alias in ISSUE_ALIASES[field]}), -1)
            for field in ISSUE_FIELDS}


def build_issues_report(sheet: SheetData, mapping: dict[str, int], *,
                        products: tuple[str, ...] = NON_OPTICAL + OPTICAL,
                        today: date | None = None,
                        year: int | None = None, start_month: int = 1,
                        end_month: int | None = None) -> IssuesReport:
    today = today or date.today()
    year = today.year if year is None else year
    end_month = today.month if end_month is None else end_month
    if not 1900 <= year <= 9999 or not 1 <= start_month <= end_month <= 12:
        raise ExcelDataError("请设置有效年份及月份：1 ≤ 起始月 ≤ 结束月 ≤ 12。")
    if any(mapping.get(field, -1) not in range(len(sheet.columns)) for field in ISSUE_FIELDS):
        raise ExcelDataError("请为山头、创建时间、产品序列号/机台编码、服务请求状态选择对应列。")
    if len({mapping[field] for field in ISSUE_FIELDS}) != 4:
        raise ExcelDataError("四个字段必须对应四个不同的列。")
    canonical = {name.casefold(): name for name in NON_OPTICAL + OPTICAL}
    months = tuple(range(start_month, end_month + 1))
    errors, notes, records = [], [], []
    for row_number, row in zip(sheet.row_numbers, sheet.rows):
        cells = {field: row[mapping[field]] for field in ISSUE_FIELDS}
        if all(cell.value is None and not cell.formula for cell in cells.values()):
            continue
        try:
            created = _text(cells["created"])
            match = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?", created)
            if not match:
                raise ExcelDataError(f"{cells['created'].coordinate} 日期无法解析：{created}")
            try:
                parsed = datetime.fromisoformat(created) if isinstance(cells['created'].value, datetime) else date(*map(int, match.groups()))
            except ValueError:
                raise ExcelDataError(f"{cells['created'].coordinate} 日期无法解析：{created}") from None
            raw = _text(cells["product"])
            product = canonical.get(raw.casefold(), raw)
            if raw.casefold() not in canonical:
                notes.append(f"第 {row_number} 行：发现名单外山头 {raw}，有效年份内保留统计。")
            if parsed.year != year:
                continue
            values = {}
            for field in ("machine", "status"):
                cell = cells[field]
                values[field] = _text(cell) if cell.formula or (cell.value is not None and str(cell.value).strip()) else ""
            if isinstance(cells['machine'].value, (int, float)) and not isinstance(cells['machine'].value, bool):
                values['machine'] = format(cells['machine'].value, '.15g')
            if not values['machine']:
                notes.append(f"第 {row_number} 行：机台编码为空，计入问题数但不计入机器数。")
            if not values['status']:
                notes.append(f"第 {row_number} 行：状态为空，计入问题数但不计入关闭数。")
            records.append(dict(product=product, month=parsed.month, machine=values['machine'] or None,
                                closed=values['status'].strip() in CLOSED_STATUSES))
        except ExcelDataError as error:
            errors.append(f"第 {row_number} 行：{error}")
    if errors:
        raise ExcelDataError("请修正以下异常行后重新生成：\n" + "\n".join(errors))
    if not records:
        raise ExcelDataError("当前时间范围无有效数据")
    frame = pd.DataFrame.from_records(records)
    if not frame['month'].between(start_month, end_month).any():
        raise ExcelDataError("当前时间范围无有效数据（月度趋势范围内无问题记录）")
    grouped = frame.groupby('product', sort=False)
    totals = grouped.size()
    machines = grouped['machine'].nunique(dropna=True)
    closed = grouped['closed'].sum()
    monthly = frame.groupby(['product', 'month']).size()
    names = [name for name in NON_OPTICAL + OPTICAL if name in totals.index]
    names += [name for name in totals.index if name not in names]
    metrics = tuple(IssueMetrics(name, int(totals[name]),
                                tuple(int(monthly.get((name, month), 0)) for month in months),
                                int(totals[name]), int(machines[name]), int(closed[name])) for name in names)
    notes.append("四图统一按所选年份筛选；趋势仅显示起始月至结束月，年度指标包含该年全部有效记录。")
    return IssuesReport(metrics, months, tuple(names), date(year, end_month, 1),
                        sheet.source, sheet.sheet_name, tuple(dict.fromkeys(notes)), year)
