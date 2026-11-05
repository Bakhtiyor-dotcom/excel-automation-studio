"""Combine, validate and summarize Excel/CSV sales files."""
from pathlib import Path
import pandas as pd

REQUIRED = ("order_id", "date", "customer", "product", "quantity", "unit_price")


def load_files(paths):
    frames = []
    for path in paths:
        path = Path(path)
        if path.suffix.lower() == ".csv":
            frame = pd.read_csv(path)
        elif path.suffix.lower() == ".xlsx":
            frame = pd.read_excel(path)
        else:
            raise ValueError(f"Unsupported file: {path.name} (use .csv or .xlsx)")
        frame.columns = frame.columns.astype(str).str.strip().str.lower().str.replace(" ", "_", regex=False)
        missing = set(REQUIRED) - set(frame.columns)
        if missing:
            raise ValueError(f"{path.name}: missing columns: {', '.join(sorted(missing))}")
        frame = frame.loc[:, REQUIRED].copy()
        frame["source_file"] = path.name
        frames.append(frame)
    if not frames:
        raise ValueError("No input files provided")
    return pd.concat(frames, ignore_index=True)


def clean_orders(df):
    """Return valid orders, rejected rows, and a count of duplicate IDs."""
    df = df.copy()
    for col in ("order_id", "customer", "product"):
        df[col] = df[col].astype("string").str.strip()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    for col in ("quantity", "unit_price"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    invalid = (df[["order_id", "customer", "product"]].isna().any(axis=1)
               | df[["order_id", "customer", "product"]].eq("").any(axis=1)
               | df["date"].isna() | df["quantity"].isna() | df["unit_price"].isna()
               | (df["quantity"] <= 0) | (df["unit_price"] < 0)
               | (df["quantity"] % 1 != 0))
    rejected = df.loc[invalid].copy()
    rejected["reason"] = "Missing or invalid required value"
    valid = df.loc[~invalid].copy()
    duplicates = valid.duplicated(subset=["order_id"], keep="first")
    duplicate_count = int(duplicates.sum())
    duplicate_rows = valid.loc[duplicates].copy()
    duplicate_rows["reason"] = "Duplicate order_id"
    rejected = pd.concat([rejected, duplicate_rows], ignore_index=True)
    valid = valid.loc[~duplicates].copy()
    valid["quantity"] = valid["quantity"].astype(int)
    valid["total"] = (valid["quantity"] * valid["unit_price"]).round(2)
    valid["date"] = valid["date"].dt.strftime("%Y-%m-%d")
    return valid, rejected, duplicate_count


def write_report(valid, rejected, output):
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    by_product = (valid.groupby("product", as_index=False, dropna=False)
                  .agg(orders=("order_id", "count"), units_sold=("quantity", "sum"), revenue=("total", "sum"))
                  .sort_values("revenue", ascending=False))
    by_customer = (valid.groupby("customer", as_index=False, dropna=False)
                   .agg(orders=("order_id", "count"), revenue=("total", "sum"))
                   .sort_values("revenue", ascending=False))
    overview = pd.DataFrame({"metric": ["Valid orders", "Rejected rows", "Total revenue", "Units sold"],
                             "value": [len(valid), len(rejected), round(float(valid["total"].sum()), 2), int(valid["quantity"].sum())]})
    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="YYYY-MM-DD") as writer:
        for name, data in [("Overview", overview), ("Cleaned Orders", valid),
                           ("By Product", by_product), ("By Customer", by_customer), ("Rejected Rows", rejected)]:
            data.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            from openpyxl.styles import Font, PatternFill, Alignment
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:
                letter = col[0].column_letter
                width = max(len(str(c.value or "")) for c in list(col)[:500]) + 3
                ws.column_dimensions[letter].width = min(max(width, 13), 38)
    return output
