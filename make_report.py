"""Build a formatted Excel sales report from a sales CSV."""

import pandas as pd


def load_sales(path):
    df = pd.read_csv(path, parse_dates=["date"])
    df["revenue"] = df["quantity"] * df["unit_price"]
    df["month"] = df["date"].dt.to_period("M").astype(str)  # e.g. "2026-03"
    return df


def summarise(df):
    """Return the summary tables that will become sheets in the report."""
    by_month = df.groupby("month", as_index=False)["revenue"].sum()

    by_product = (
        df.groupby("product", as_index=False)
        .agg(orders=("order_id", "count"), units=("quantity", "sum"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )

    by_region = (
        df.groupby("region", as_index=False)
        .agg(orders=("order_id", "count"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )

    return by_month, by_product, by_region


def main():
    df = load_sales("sample_data/sales.csv")
    by_month, by_product, by_region = summarise(df)

    print(f"Total revenue: £{df['revenue'].sum():,.2f} from {len(df)} orders")
    print("\nBy month")
    print(by_month)
    print("\nBy product")
    print(by_product)
    print("\nBy region")
    print(by_region)


if __name__ == "__main__":
    main()