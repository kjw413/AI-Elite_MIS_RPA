"""Production collection contracts shared by the UI and workbook readers."""
from datetime import date
import re


VERIFIED_PERIOD_MARKER = "__DATE_RANGE_VERIFIED_V1__"


def validate_grid_dates(headers, start: date, end: date) -> set[int]:
    """Reject a different period/duplicate day, allowing omitted nonproduction days."""
    if start > end or (start.year, start.month) != (end.year, end.month):
        raise ValueError(f"생산 조회 기간은 같은 월이어야 합니다: {start} ~ {end}")
    if not any(re.fullmatch(r"item\s*code", str(h).strip(), re.IGNORECASE) for h in headers):
        raise ValueError("생산 그리드 Item Code 머리글이 없습니다.")
    days: set[int] = set()
    for header in headers:
        match = re.fullmatch(r"(\d{1,2})\s*일?", str(header).strip())
        if not match:
            continue
        day = int(match.group(1))
        if day in days or not start.day <= day <= end.day:
            raise ValueError(f"생산 그리드 일자 오류: {header!r} ({start} ~ {end})")
        days.add(day)
    return days
