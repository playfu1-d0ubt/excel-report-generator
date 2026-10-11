"""Generate a fake sales CSV for testing the report generator."""

import random

import pandas as pd

random.seed(42)  # same data every time

products = {"Notebook": 3.5, "Pen set": 6.0, "Desk lamp": 24.0, "Backpack": 39.0, "Headphones": 59.0}
regions = ["North", "South", "East", "West"]

rows = []
for order_id in range(1, 401):
    product = random.choice(list(products))
    quantity = random.randint(1, 6)
    rows.append({
        "order_id": order_id,
        "date": pd.Timestamp("2026-01-01") + pd.Timedelta(days=random.randint(0, 270)),
        "region": random.choice(regions),
        "product": product,
        "quantity": quantity,
        "unit_price": products[product],
    })

pd.DataFrame(rows).sort_values("date").to_csv("sample_data/sales.csv", index=False)
print("Saved sample_data/sales.csv")