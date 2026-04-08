from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta
from typing import Callable, Iterable, List, Optional, Set


# -----------------------------
# Result object
# -----------------------------

@dataclass(frozen=True)
class WorkCalendarResult:
    from_date: date
    to_date: date

    theoretical_work_days_count: int
    theoretical_work_days_dates: List[date]

    holiday_days_count: int
    holiday_dates: List[date]

    approved_leave_days_count: int
    approved_leave_dates: List[date]

    net_work_days_count: int
    net_work_days_dates: List[date]


# -----------------------------
# Service
# -----------------------------

HolidayDatesProvider = Callable[[date, date], Iterable[date]]
ApprovedLeaveDatesProvider = Callable[[int, date, date], Iterable[date]]


class WorkCalendarService:
    """
    Pure business-logic service for calculating work days within a period.

    Design guarantees:
    - Deterministic: same inputs -> same outputs.
    - No UI/Qt dependencies.
    - No direct DB dependency: data is injected via providers.
    - No double-subtraction: holidays removed first, then approved leaves.
    """

    def __init__(
        self,
        holiday_dates_provider: Optional[HolidayDatesProvider] = None,
        approved_leave_dates_provider: Optional[ApprovedLeaveDatesProvider] = None,
    ) -> None:
        self._holiday_dates_provider = holiday_dates_provider
        self._approved_leave_dates_provider = approved_leave_dates_provider

    def calculate_period(
        self,
        from_date: date,
        to_date: date,
        work_days_indices: List[int],
        employee_id: Optional[int] = None,
        include_holidays: bool = True,
        include_leaves: bool = True,
    ) -> WorkCalendarResult:
        self._validate_inputs(from_date, to_date, work_days_indices, employee_id, include_holidays, include_leaves)

        theoretical_set = self._build_theoretical_work_days_set(from_date, to_date, work_days_indices)
        theoretical_dates = sorted(theoretical_set)
        theoretical_count = len(theoretical_dates)

        holiday_set: Set[date] = set()
        if include_holidays:
            holiday_set = self._get_holiday_dates_set(from_date, to_date)
            holiday_set = holiday_set.intersection(theoretical_set)

        leave_set: Set[date] = set()
        if include_leaves and employee_id is not None:
            leave_set = self._get_approved_leave_dates_set(employee_id, from_date, to_date)
            # Prevent double subtraction: leaves only apply to theoretical days that are not already holidays
            leave_set = leave_set.intersection(theoretical_set.difference(holiday_set))

        net_set = theoretical_set.difference(holiday_set).difference(leave_set)
        net_dates = sorted(net_set)
        net_count = len(net_dates)

        holiday_dates = sorted(holiday_set)
        leave_dates = sorted(leave_set)

        return WorkCalendarResult(
            from_date=from_date,
            to_date=to_date,

            theoretical_work_days_count=theoretical_count,
            theoretical_work_days_dates=theoretical_dates,

            holiday_days_count=len(holiday_dates),
            holiday_dates=holiday_dates,

            approved_leave_days_count=len(leave_dates),
            approved_leave_dates=leave_dates,

            net_work_days_count=net_count,
            net_work_days_dates=net_dates,
        )

    # -----------------------------
    # Internals
    # -----------------------------

    @staticmethod
    def _validate_inputs(
        from_date: date,
        to_date: date,
        work_days_indices: List[int],
        employee_id: Optional[int],
        include_holidays: bool,
        include_leaves: bool,
    ) -> None:
        if from_date is None or to_date is None:
            raise ValueError("from_date and to_date are required.")

        if from_date > to_date:
            raise ValueError("from_date must be <= to_date.")

        if not isinstance(work_days_indices, list) or len(work_days_indices) == 0:
            raise ValueError("work_days_indices must be a non-empty list of integers 0..6.")

        invalid = [x for x in work_days_indices if not isinstance(x, int) or x < 0 or x > 6]
        if invalid:
            raise ValueError(f"work_days_indices contains invalid values: {invalid}. Expected integers 0..6.")

        if employee_id is not None and (not isinstance(employee_id, int) or employee_id <= 0):
            raise ValueError("employee_id must be a positive int when provided.")

        if include_leaves and employee_id is None:
            # Not an error—just a clear rule: leaves require employee_id.
            # We keep it strict to avoid silent misunderstandings.
            raise ValueError("include_leaves=True requires employee_id to be provided.")

        if not isinstance(include_holidays, bool) or not isinstance(include_leaves, bool):
            raise ValueError("include_holidays and include_leaves must be booleans.")

    @staticmethod
    def _build_theoretical_work_days_set(
        from_date: date,
        to_date: date,
        work_days_indices: List[int],
    ) -> Set[date]:
        indices = set(work_days_indices)
        result: Set[date] = set()

        current = from_date
        while current <= to_date:
            if current.weekday() in indices:
                result.add(current)
            current += timedelta(days=1)

        return result

    def _get_holiday_dates_set(self, from_date: date, to_date: date) -> Set[date]:
        if self._holiday_dates_provider is None:
            # Explicit failure is safer than silently ignoring holidays.
            raise RuntimeError("holiday_dates_provider is not configured, but include_holidays=True was requested.")

        dates = self._holiday_dates_provider(from_date, to_date)
        return self._normalize_dates_iterable(dates)

    def _get_approved_leave_dates_set(self, employee_id: int, from_date: date, to_date: date) -> Set[date]:
        if self._approved_leave_dates_provider is None:
            raise RuntimeError("approved_leave_dates_provider is not configured, but include_leaves=True was requested.")

        dates = self._approved_leave_dates_provider(employee_id, from_date, to_date)
        return self._normalize_dates_iterable(dates)

    @staticmethod
    def _normalize_dates_iterable(dates: Iterable[date]) -> Set[date]:
        result: Set[date] = set()
        for d in dates:
            if not isinstance(d, date):
                raise ValueError(f"Provider returned a non-date value: {d!r}")
            result.add(d)
        return result
