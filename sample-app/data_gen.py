import random
from datetime import datetime, timedelta
from typing import List, Dict


STATUSES = ["New", "Ok", "Hold"]


def _random_datetime(start: datetime, end: datetime) -> datetime:
    if start > end:
        start, end = end, start
    delta = end - start
    if delta.total_seconds() <= 0:
        return start
    offset = random.randint(0, int(delta.total_seconds()))
    return start + timedelta(seconds=offset)


def generate_rows(
    count: int,
    date_from: datetime,
    date_to: datetime,
    category: str,
    min_amount: float,
) -> List[Dict[str, str]]:
    rows = []
    for i in range(1, count + 1):
        ts = _random_datetime(date_from, date_to)
        amount = round(random.uniform(min_amount, min_amount + 1000.0), 2)
        rows.append(
            {
                "Id": i,
                "Timestamp": ts.strftime("%Y-%m-%d %H:%M:%S"),
                "Category": category,
                "Customer": f"Customer-{i:03d}",
                "Amount": amount,
                "Status": random.choice(STATUSES),
            }
        )
    return rows
