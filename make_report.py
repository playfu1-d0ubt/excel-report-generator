"""Build a formatted Excel sales report from a sales CSV."""

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
MONEY = "£#,##0.00"

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


def style_table(ws, money_columns):
    """Format a sheet that holds one table starting at A1."""
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"  # keep the header visible when scrolling

    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        longest = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
        ws.column_dimensions[letter].width = longest + 4

        if column[0].value in money_columns:
            for cell in column[1:]:
                cell.number_format = MONEY


def write_report(path, by_month, by_product, by_region):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        by_month.to_excel(writer, sheet_name="By month", index=False)
        by_product.to_excel(writer, sheet_name="By product", index=False)
        by_region.to_excel(writer, sheet_name="By region", index=False)

        for ws in writer.book.worksheets:
            style_table(ws, money_columns=["revenue"])

def main():
    df = load_sales("sample_data/sales.csv")
    by_month, by_product, by_region = summarise(df)
    write_report("sales_report.xlsx", by_month, by_product, by_region)
    print(f"Saved sales_report.xlsx ({len(df)} orders, £{df['revenue'].sum():,.2f} revenue)")


if __name__ == "__main__":
    main()