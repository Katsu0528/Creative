#!/usr/bin/env python3
"""Generate a payment certificate from CSV data.

Reads columns C (item name), G (amount), K (date) and aggregates amounts by month.
Outputs a markdown payment certificate with monthly totals and per-item breakdown.
"""
from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable


PAYMENT_DATES_BY_MONTH = {
    (2024, 12): date(2025, 1, 31),
    (2025, 1): date(2025, 2, 28),
    (2025, 2): date(2025, 3, 31),
    (2025, 3): date(2025, 4, 30),
    (2025, 4): date(2025, 5, 30),
    (2025, 5): date(2025, 6, 30),
    (2025, 6): date(2025, 7, 31),
    (2025, 7): date(2025, 8, 29),
    (2025, 8): date(2025, 9, 30),
    (2025, 9): date(2025, 10, 31),
    (2025, 10): date(2025, 11, 28),
    (2025, 11): date(2025, 12, 30),
}


@dataclass
class Row:
    item: str
    amount: Decimal
    day: date


def parse_amount(raw: str) -> Decimal:
    cleaned = raw.replace(',', '').replace('¥', '').replace('円', '').strip()
    if not cleaned:
        return Decimal(0)
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        raise ValueError(f"Invalid amount: {raw}")


def parse_date(raw: str) -> date:
    value = raw.strip()
    for fmt in (
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%Y.%m.%d",
        "%Y年%m月%d日",
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    ):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Unsupported date format: {raw}")


def read_rows(path: Path) -> Iterable[Row]:
    encodings = ["shift_jis", "cp932", "utf-8-sig"]
    last_error: Exception | None = None
    for encoding in encodings:
        try:
            with path.open("r", encoding=encoding, newline="") as file:
                reader = csv.reader(file)
                for row in reader:
                    if len(row) < 11:
                        continue
                    item = row[2].strip()
                    if not item:
                        continue
                    try:
                        amount = parse_amount(row[6])
                        day = parse_date(row[10])
                    except ValueError:
                        continue
                    yield Row(item=item, amount=amount, day=day)
            return
        except Exception as exc:  # pragma: no cover - fallback for encodings
            last_error = exc
            continue
    raise RuntimeError(f"Failed to read CSV: {path}") from last_error


def format_currency(amount: Decimal) -> str:
    return f"{int(amount):,}円"


def payment_date_for(month: int, year: int) -> date:
    key = (year, month)
    if key in PAYMENT_DATES_BY_MONTH:
        return PAYMENT_DATES_BY_MONTH[key]
    if month == 12:
        return date(year, month, 31)
    return date(year, month + 1, 1) - timedelta(days=1)


def generate_certificate(rows: Iterable[Row]) -> str:
    summary: dict[tuple[int, int], dict[str, Decimal]] = {}
    totals: dict[tuple[int, int], Decimal] = {}
    for row in rows:
        key = (row.day.year, row.day.month)
        summary.setdefault(key, {})
        summary[key][row.item] = summary[key].get(row.item, Decimal(0)) + row.amount
        totals[key] = totals.get(key, Decimal(0)) + row.amount

    today = date.today()
    date_line = f"{today.year}年{today.month}月{today.day}日"
    output = []
    output.append(f"{'':>40}{date_line}")
    output.append("宛名欄として様")
    output.append(f"{'':>40}〒141-0031")
    output.append(f"{'':>40}東京都品川区西五反田1−3−8")
    output.append(f"{'':>40}五反田PLACE　2F")
    output.append(f"{'':>40}株式会社OTONARI")
    output.append("")
    output.append(f"{'支払証明書':^40}")
    output.append("")
    output.append("下記の支払を行いましたことを、本状にて証明いたします。")
    output.append("")
    output.append("| 支払日 | 金額 | 内容 |")
    output.append("| --- | --- | --- |")

    for (year, month) in sorted(totals.keys()):
        total = totals[(year, month)]
        if total == 0:
            continue
        payment_day = payment_date_for(month, year)
        items = summary[(year, month)]
        detail_lines = [f"{name}：{format_currency(amount)}" for name, amount in items.items() if amount != 0]
        detail_cell = "<br>".join(detail_lines) if detail_lines else ""
        output.append(
            f"| {payment_day.year}年{payment_day.month}月{payment_day.day}日 | {format_currency(total)} | {detail_cell} |"
        )

    return "\n".join(output) + "\n"


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate payment certificate from CSV.")
    parser.add_argument("csv", type=Path, help="Input CSV file")
    parser.add_argument("-o", "--output", type=Path, default=Path("payment_certificate.md"))
    args = parser.parse_args()

    content = generate_certificate(read_rows(args.csv))
    args.output.write_text(content, encoding="utf-8")
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
