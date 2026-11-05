import pandas as pd
import pytest
from excel_tool.processor import load_files, clean_orders, write_report


def test_sample_cleaning():
    data = load_files(["samples/orders_january.csv", "samples/orders_february.csv"])
    valid, rejected, duplicate_count = clean_orders(data)
    assert len(valid) == 6
    assert len(rejected) == 3
    assert duplicate_count == 1
    assert valid["total"].sum() == pytest.approx(708)


def test_report_creation(tmp_path):
    data = load_files(["samples/orders_january.csv"])
    valid, rejected, _ = clean_orders(data)
    path = write_report(valid, rejected, tmp_path / "report.xlsx")
    assert path.exists()
    book = pd.ExcelFile(path)
    assert book.sheet_names == ["Overview", "Cleaned Orders", "By Product", "By Customer", "Rejected Rows"]


def test_missing_columns(tmp_path):
    path = tmp_path / "bad.csv"
    path.write_text("order_id,customer\n1,Alice\n", encoding="utf-8")
    with pytest.raises(ValueError, match="missing columns"):
        load_files([path])


def test_blank_order_id_rejected():
    df = pd.DataFrame([dict(order_id=" ", date="2026-01-01", customer="A", product="B", quantity=1, unit_price=5, source_file="x")])
    valid, rejected, _ = clean_orders(df)
    assert valid.empty and len(rejected) == 1
