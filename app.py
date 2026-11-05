"""Excel Automation Studio v2: flexible column mapping and sales dashboard."""
from pathlib import Path
from tempfile import TemporaryDirectory
import pandas as pd
import streamlit as st
from excel_tool.processor import REQUIRED, load_files, clean_orders, write_report

st.set_page_config(page_title="Excel Automation Studio v2", page_icon="📊", layout="wide")
st.markdown('''<style>
.block-container{padding-top:2rem;max-width:1250px}
[data-testid="stMetric"]{background:rgba(100,116,139,.12);padding:18px;border-radius:12px;border:1px solid rgba(148,163,184,.3)}
[data-testid="stMetricLabel"], [data-testid="stMetricValue"]{color:inherit!important}
[data-testid="stMetricValue"] {
    font-size: clamp(1.1rem, 1.6vw, 1.8rem);
    overflow-wrap: anywhere;
}

[data-testid="stMetric"] {
    min-width: 0;
}
h1{letter-spacing:-.035em}
</style>''', unsafe_allow_html=True)
st.title("Excel Automation Studio")
st.caption("Upload sales data, map columns, validate orders, and download a clean Excel report.")
st.divider()

ALIASES = {
    "order_id": ["order_id", "order id", "id", "номер заказа", "заказ", "номер_заказа"],
    "date": ["date", "order_date", "дата", "дата заказа", "дата_заказа"],
    "customer": ["customer", "client", "buyer", "клиент", "покупатель"],
    "product": ["product", "item", "товар", "продукт", "наименование"],
    "quantity": ["quantity", "qty", "units", "количество", "кол-во"],
    "unit_price": ["unit_price", "unit price", "price", "цена", "цена за единицу", "цена_за_единицу"],
}

def read_uploaded(upload):
    if upload.name.lower().endswith(".csv"):
        try:
            return pd.read_csv(upload, encoding="utf-8-sig")
        except UnicodeDecodeError:
            upload.seek(0)
            return pd.read_csv(upload, encoding="cp1251")
    return pd.read_excel(upload)

def guess_column(field, columns):
    for alias in ALIASES[field]:
        for col in columns:
            if str(col).strip().casefold().replace("_", " ") == alias.casefold().replace("_", " "):
                return col
    return None

with st.sidebar:
    st.header("Data source")
    uploads = st.file_uploader("Upload CSV or XLSX files", type=["csv", "xlsx"], accept_multiple_files=True)
    demo = st.checkbox("Use bundled sample files", value=not bool(uploads), disabled=bool(uploads), help="Turn off to use your own files")
    st.caption("Files are processed in this running session. Do not upload confidential data to an untrusted deployment.")

if not uploads and not demo:
    st.info("Upload files or enable the sample dataset to start.")
    st.stop()

try:
    with TemporaryDirectory() as tmp:
        directory = Path(tmp)
        if uploads:
            st.subheader("1. Match your columns")
            st.caption("Choose which column in each file represents each required field. Extra columns are ignored.")
            frames = []
            for i, upload in enumerate(uploads):
                with st.expander(f"{upload.name} — column mapping", expanded=True):
                    upload.seek(0)
                    df = read_uploaded(upload)
                    if df.empty:
                        st.error("This file has no rows.")
                        st.stop()
                    columns = list(df.columns)
                    choices = ["— Select column —"] + list(range(len(columns)))
                    mapped = {}
                    left, right = st.columns(2)
                    for j, field in enumerate(REQUIRED):
                        guess = guess_column(field, columns)
                        index = (columns.index(guess) + 1) if guess in columns else 0
                        with (left if j % 2 == 0 else right):
                            chosen = st.selectbox(field.replace("_", " ").title(), choices, index=index,
                                format_func=lambda v: "— Select column —" if isinstance(v, str) else str(columns[v]),
                                key=f"map_{i}_{field}")
                        mapped[field] = chosen
                    selected = [v for v in mapped.values() if isinstance(v, int)]
                    if len(selected) != len(REQUIRED):
                        st.warning("Select all six required columns to continue.")
                        st.stop()
                    if len(set(selected)) != len(selected):
                        st.error("Each field must use a different source column.")
                        st.stop()
                    normalized = df.iloc[:, [mapped[field] for field in REQUIRED]].copy()
                    normalized.columns = list(REQUIRED)
                    normalized["source_file"] = upload.name
                    frames.append(normalized)
            raw = pd.concat(frames, ignore_index=True)
        else:
            paths = sorted((Path(__file__).parent / "samples").glob("*.csv"))
            raw = load_files(paths)
        valid, rejected, duplicates = clean_orders(raw)
        report_path = write_report(valid, rejected, directory / "sales_report.xlsx")
        report_bytes = report_path.read_bytes()
except (ValueError, OSError, KeyError, pd.errors.ParserError, ImportError) as exc:
    st.error(f"Unable to process data: {exc}")
    st.stop()

st.subheader("2. Sales overview" if uploads else "Sales overview")
a,b,c,d = st.columns(4)
a.metric("Valid orders", f"{len(valid):,}")
b.metric("Rejected rows", f"{len(rejected):,}")
c.metric("Revenue", f"{valid['total'].sum():,.2f}")
d.metric("Units sold", f"{valid['quantity'].sum():,}")
st.caption(f"Removed duplicate order IDs: {duplicates} · Source rows: {len(raw)}")
product = (valid.groupby("product", as_index=False).agg(orders=("order_id","count"),units_sold=("quantity","sum"),revenue=("total","sum")).sort_values("revenue",ascending=False))
customer = (valid.groupby("customer", as_index=False).agg(orders=("order_id","count"),revenue=("total","sum")).sort_values("revenue",ascending=False))
left,right = st.columns(2)
with left:
    st.subheader("Revenue by product")
    if product.empty: st.info("No valid products")
    else: st.bar_chart(product.set_index("product")["revenue"], horizontal=True)
with right:
    st.subheader("Revenue by customer")
    if customer.empty: st.info("No valid customers")
    else: st.bar_chart(customer.set_index("customer")["revenue"], horizontal=True)
tabs = st.tabs(["Cleaned orders", "By product", "By customer", "Rejected rows"])
for tab, data in zip(tabs, [valid, product, customer, rejected]):
    with tab:
        st.dataframe(data, use_container_width=True, hide_index=True)
st.download_button("Download Excel report", report_bytes, file_name="sales_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
st.caption("Portfolio demo · Python / Streamlit / Pandas · v2")
