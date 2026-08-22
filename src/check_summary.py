"""수표정리 - the check summary block that tallies checks by amount.

Mirrors how the cash section works: a fixed set of predefined amount rows that
always appear (even at zero), followed by any amount that turned up in the
batch but is not predefined, appended in ascending order after the last
predefined row.
"""

from collections import Counter
from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence

# Confirm this list against the paper form before the church test.
DEFAULT_PREDEFINED_AMOUNTS: Sequence[float] = (5, 10, 20, 25, 50, 100)


@dataclass
class SummaryRow:
    amount: float
    count: int
    total: float
    predefined: bool

    def as_tuple(self):
        return (self.amount, self.count, self.total)


def build_check_summary(amounts: Iterable[Optional[float]],
                        predefined: Sequence[float] = DEFAULT_PREDEFINED_AMOUNTS) -> List[SummaryRow]:
    """Tally check amounts into summary rows.

    Predefined amounts always appear, in the order given, even when no check
    matched them. Anything else is appended afterwards in ascending order, so
    the familiar rows never move around on the operator's form.

    None amounts (unreadable checks) are skipped - they are surfaced through
    the review sidecar instead of being silently bucketed as zero.
    """
    predefined_list = list(predefined)
    predefined_set = set(predefined_list)

    counts = Counter(a for a in amounts if a is not None)

    rows = [
        SummaryRow(amount=amount, count=counts.get(amount, 0),
                   total=round(amount * counts.get(amount, 0), 2), predefined=True)
        for amount in predefined_list
    ]
    extras = sorted(a for a in counts if a not in predefined_set)
    rows.extend(
        SummaryRow(amount=amount, count=counts[amount],
                   total=round(amount * counts[amount], 2), predefined=False)
        for amount in extras
    )
    return rows


def summary_totals(rows: Sequence[SummaryRow]):
    """(total check count, total dollar amount) across the summary."""
    return sum(r.count for r in rows), round(sum(r.total for r in rows), 2)


def write_check_summary(worksheet, rows: Sequence[SummaryRow], start_row: int,
                        amount_col: int, count_col: Optional[int] = None,
                        total_col: Optional[int] = None) -> int:
    """Write summary rows into a worksheet, returning the next free row.

    Anchor coordinates are arguments rather than constants because the
    production template has not been confirmed yet. Predefined rows are
    written in place; extras extend downward from the last predefined row.
    """
    count_col = count_col if count_col is not None else amount_col + 1
    total_col = total_col if total_col is not None else amount_col + 2

    row_number = start_row
    for row in rows:
        worksheet.cell(row=row_number, column=amount_col, value=row.amount)
        worksheet.cell(row=row_number, column=count_col, value=row.count)
        worksheet.cell(row=row_number, column=total_col, value=row.total)
        row_number += 1
    return row_number
