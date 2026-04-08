from datetime import date
from typing import Iterable
from database import DatabaseManager


def holiday_dates_provider(db: DatabaseManager):
    """
    دالة تقرأ العطل الرسمية من قاعدة البيانات
    وتعيد جميع تواريخ العطل ضمن فترة معينة
    """

    def _provider(from_date: date, to_date: date) -> Iterable[date]:
        results = []

        rows = db.fetch_all(
            "SELECT holiday_date, type, year FROM holidays"
        )

        for hdate, htype, hyear in rows:
            if htype == 'fixed':
                # عطلة ثابتة (MM-DD) تتكرر كل سنة
                month, day = map(int, hdate.split('-'))
                for y in range(from_date.year, to_date.year + 1):
                    d = date(y, month, day)
                    if from_date <= d <= to_date:
                        results.append(d)
            else:
                # عطلة متغيرة (YYYY-MM-DD)
                y, m, d = map(int, hdate.split('-'))
                d = date(y, m, d)
                if from_date <= d <= to_date:
                    results.append(d)

        return results

    return _provider
