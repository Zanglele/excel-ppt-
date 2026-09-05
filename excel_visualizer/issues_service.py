"""第二份 Excel 的问题次数、月度趋势、机器去重和关闭率统计。"""

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date, datetime
import re

from excel_visualizer.data_service import ExcelDataError, NON_OPTICAL, OPTICAL, SheetData, _text


ISSUE_YEAR = 2026
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
    all_years: bool
    source: str
    sheet_name: str
    notes: tuple[str, ...]

    @property
    def scope(self) -> str:
        return "全部年份" if self.all_years else "2026年"


def guess_issue_columns(sheet: SheetData) -> dict[str, int]:
    names = [re.sub(r"\s+", "", name.rsplit(" [", 1)[0]).casefold() for name in sheet.columns]
    return {field: next((i for i, name in enumerate(names)
                        if name in {alias.casefold() for alias in ISSUE_ALIASES[field]}), -1)
            for field in ISSUE_FIELDS}


def build_issues_report(sheet: SheetData, mapping: dict[str, int], *,
                        products: tuple[str, ...] = NON_OPTICAL + OPTICAL,
                        today: date | None = None, all_years: bool = True) -> IssuesReport:
    today = today or date.today()
    if today.year < ISSUE_YEAR:
        raise ExcelDataError("当前日期尚未进入 2026 年，无法生成 2026 年月度趋势。")
    if any(mapping.get(field, -1) not in range(len(sheet.columns)) for field in ISSUE_FIELDS):
        raise ExcelDataError("请为山头、创建时间、产品序列号/机台编码、服务请求状态选择对应列。")
    if len({mapping[field] for field in ISSUE_FIELDS}) != 4:
        raise ExcelDataError("四个字段必须对应四个不同的列。")
    names = list(dict.fromkeys(NON_OPTICAL + OPTICAL + tuple(products)))
    canonical = {name.casefold(): name for name in names}
    months = tuple(range(1, (today.month if today.year == ISSUE_YEAR else 12) + 1))
    annual, totals, closed, monthly = Counter(), Counter(), Counter(), Counter()
    machines = defaultdict(set)
    errors, extra_names = [], []
    observed_names = set()
    excluded = missing_machine = missing_status = later_months = 0
    records = 0
    for row_number, row in zip(sheet.row_numbers, sheet.rows):
        cells = {field: row[mapping[field]] for field in ISSUE_FIELDS}
        if all(cell.value is None and not cell.formula for cell in cells.values()):
            continue
        try:
            created = _text(cells["created"])
            # 按用户规定先判断前四个字符；Excel 日期对象的字符串也以年份开头。
            if not re.match(r"^\d{4}", created):
                raise ExcelDataError(f"{cells['created'].coordinate} 创建时间须以四位年份开头")
            raw_product = _text(cells["product"])
            product = canonical.get(raw_product.casefold(), raw_product)
            if product.casefold() not in canonical:
                canonical[product.casefold()] = product
                names.append(product)
                extra_names.append(product)
            observed_names.add(product)
            in_year = created[:4] == str(ISSUE_YEAR)
            if not in_year and not all_years:
                excluded += 1
                continue
            month = None
            if in_year:
                value = cells["created"].value
                if isinstance(value, (date, datetime)):
                    parsed_date = value
                else:
                    match = re.fullmatch(r"(2026)-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?", created)
                    if not match:
                        raise ExcelDataError(f"{cells['created'].coordinate} 创建时间格式应为 2026-月-日（可带时间）")
                    try:
                        parsed_date = date(*(int(part) for part in match.groups()))
                    except ValueError:
                        raise ExcelDataError(f"{cells['created'].coordinate} 创建时间不是有效日期：{created}") from None
                month = parsed_date.month
            machine_cell, status_cell = cells["machine"], cells["status"]
            machine = _text(machine_cell) if machine_cell.value is not None or machine_cell.formula else ""
            status = _text(status_cell) if status_cell.value is not None or status_cell.formula else ""
            # 数字形式的编号 123 与 123.0 视为同一编号；文本前导零保留。
            if isinstance(machine_cell.value, (int, float)) and not isinstance(machine_cell.value, bool):
                machine = format(machine_cell.value, ".15g")
            records += 1
            totals[product] += 1
            closed[product] += status in CLOSED_STATUSES
            if machine.strip():
                machines[product].add(machine.strip())
            else:
                missing_machine += 1
            missing_status += not bool(status.strip())
            if in_year:
                annual[product] += 1
                monthly[product, month] += 1
                later_months += month not in months
            else:
                excluded += 1
        except ExcelDataError as error:
            errors.append(f"第 {row_number} 行：{error}")
    if errors:
        raise ExcelDataError(f"发现 {len(errors)} 行数据问题，请修正后生成（未跳过错误行）：\n" + "\n".join(errors[:20]))
    if not records:
        raise ExcelDataError("没有符合统计范围的问题记录，请检查年份、工作表和表头行。")
    notes = []
    if excluded:
        notes.append(f"{excluded} 条非 2026 年记录未计入前两图" + ("。" if not all_years else "，已计入后两图。"))
    if later_months:
        notes.append(f"{later_months} 条 2026 年后续月份记录计入年度总数，未计入截至本月的折线图。")
    if missing_machine:
        notes.append(f"{missing_machine} 条记录缺少机器编号：问题数仍计入，空编号不计为机器；密度可能偏高。")
    if missing_status:
        notes.append(f"{missing_status} 条记录缺少状态：计入总问题数，不计入关闭数。")
    if extra_names:
        notes.append("发现名单外山头，已追加保留：" + "、".join(extra_names))
    notes.append("机器数为问题表中非空编号的去重数；无机器或无问题时，相应比值显示 N/A。")
    return IssuesReport(tuple(IssueMetrics(name, annual[name], tuple(monthly[name, m] for m in months),
                                           totals[name], len(machines[name]), closed[name]) for name in names),
                        months, tuple(name for name in names if name in observed_names), today, all_years,
                        sheet.source, sheet.sheet_name, tuple(notes))
