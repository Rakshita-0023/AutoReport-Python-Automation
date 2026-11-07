"""Regenerate the deterministic synthetic sales export."""

import csv
from datetime import date, timedelta
from pathlib import Path
from random import Random

from src.cleaner import COLUMNS, DATE_FORMATS

PRODUCTS = {
    "Laptop Stand": 1499,
    "Wireless Mouse": 849,
    "USB-C Hub": 2199,
    "Desk Mat": 649,
    "Mechanical Keyboard": 3299,
    "Webcam": 2499,
    "USB-C Cable": 399,
    "Monitor Riser": 1199,
}


def generate(path: Path) -> None:
    rng = Random(42)
    start = date(2025, 1, 1)
    rows = []
    for index in range(970):
        day = start + timedelta(days=rng.randrange(181))
        product = rng.choice(list(PRODUCTS))
        region = rng.choices(["North", "South", "West", "East"], [29, 32, 25, 14])[0]
        quantity = rng.choices([1, 2, 3, 4, 5, 6], [35, 28, 17, 10, 6, 4])[0]
        price = PRODUCTS[product] - rng.choice([0, 0, 0, 50, 100])
        rows.append([f"S-{index + 1:05}", day.strftime(rng.choice(DATE_FORMATS)),
                     product, region, quantity, f"{price}.00"])

    for index in range(10):
        rows[index][5] = ""
    for index in range(10, 18):
        rows[index][3] = ""
    for index in range(18, 24):
        rows[index][1] = "31/02/2025"
    for index in range(24, 28):
        rows[index][4] = "two"
    for index in range(28, 36):
        rows[index][2] = f" {rows[index][2]} "
    rows.extend(row.copy() for row in rng.sample(rows[36:], 30))
    rng.shuffle(rows)

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as target:
        writer = csv.writer(target, lineterminator="\n")
        writer.writerow(COLUMNS)
        writer.writerows(rows)


if __name__ == "__main__":
    generate(Path(__file__).parent / "data" / "sample_sales.csv")
