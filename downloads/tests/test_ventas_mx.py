"""End-to-end checks for the ventas_mx spec against the real downloads.

Skipped automatically when no Ventas export is present, so the suite still runs on a
clean checkout.
"""

from __future__ import annotations

import pandas as pd
import pytest

from mx_sales import pipeline
from mx_sales.readers.excel import _deduplicate, find_header_row
from mx_sales.reports import ventas_mx


@pytest.fixture(scope="module")
def sales_files():
    """Snapshots of one store, so the dedup tests compare like with like."""
    items = [
        item
        for item in pipeline.discover()
        if item.spec.name == "ventas_mx" and "BOCINA_SM" in str(item.path)
    ]
    if not items:
        pytest.skip("no Ventas MX export found under the raw directory")
    return items


@pytest.fixture(scope="module")
def raw(sales_files):
    return ventas_mx.load(sales_files[-1].path)


@pytest.fixture(scope="module")
def clean(raw):
    return ventas_mx.transform(raw)


class TestGroupedHeader:
    def test_duplicate_names_are_disambiguated_by_group(self, raw):
        # "Unidades" appears under Ventas, Devoluciones and Reclamos.
        unidades = [c for c in raw.columns if c.endswith("__unidades")]
        assert set(unidades) == {"ventas__unidades", "devoluciones__unidades", "reclamos__unidades"}

    def test_columns_are_unique(self, raw):
        assert len(set(raw.columns)) == len(raw.columns)

    def test_group_labels_are_forward_filled(self, raw):
        assert "compradores__estado" in raw.columns
        assert "ventas__estado" in raw.columns


class TestTransform:
    def test_order_id_is_populated(self, clean):
        assert clean["order_id"].notna().all()

    def test_the_record_key_is_unique(self, clean):
        # order_id alone is not a key: a package parent shares nothing with its
        # children, but an exchange puts the money and the product on two rows that
        # carry the same order id.
        key = ["order_id", "sku", "row_role", "record_seq"]
        assert not clean[key].duplicated().any()

    def test_expected_dtypes(self, clean):
        assert clean["sold_at"].dtype == "datetime64[ns]"
        assert clean["units"].dtype == "Int64"
        assert clean["total_mxn"].dtype == "Float64"
        assert clean["is_advertising_sale"].dtype == "boolean"
        assert clean["sku"].dtype == "string"

    def test_blank_columns_are_null_not_whitespace(self, clean):
        # purchase_order is " " in every row of the raw export.
        assert clean["purchase_order"].isna().all()

    def test_no_string_column_holds_bare_whitespace(self, clean):
        for column in clean.select_dtypes(include="string").columns:
            values = clean[column].dropna()
            assert not values.str.fullmatch(r"\s*").any(), column

    def test_inferred_years_stay_close_to_the_sale(self, clean):
        # The year-less dates are resolved against the sale date. Landing more than a
        # few months away means the wrong year was chosen.
        for column in ["shipped_at", "delivered_at", "return_reviewed_at"]:
            both = clean.dropna(subset=["sold_at", column])
            distance = (both[column] - both["sold_at"]).abs()
            assert (distance <= pd.Timedelta(days=120)).all(), column

    def test_no_date_lands_in_a_year_absent_from_the_export(self, clean):
        years = set(clean["sold_at"].dt.year.dropna())
        for column in ["shipped_at", "delivered_at", "return_reviewed_at"]:
            inferred = set(clean[column].dt.year.dropna())
            assert inferred <= years | {min(years) - 1, max(years) + 1}, column

    def test_output_schema_is_fixed(self, clean):
        assert list(clean.columns) == ventas_mx.OUTPUT_ORDER

    def test_missing_required_column_raises(self):
        with pytest.raises(ValueError, match="expected columns missing"):
            ventas_mx.transform(pd.DataFrame({"ventas__de_venta": ["1"]}))


class TestDeduplication:
    def test_snapshots_collapse_to_distinct_records(self, sales_files):
        frames = [pipeline.process_file(item) for item in sales_files]
        combined = pd.concat(frames, ignore_index=True)
        result = pipeline.deduplicate(combined, ventas_mx.SPEC)

        key = list(ventas_mx.SPEC.primary_key)
        assert len(result) == len(combined[key].drop_duplicates())
        assert not result[key].duplicated().any()

    def test_no_order_is_lost_by_deduplication(self, sales_files):
        frames = [pipeline.process_file(item) for item in sales_files]
        combined = pd.concat(frames, ignore_index=True)
        result = pipeline.deduplicate(combined, ventas_mx.SPEC)
        assert set(result["order_id"]) == set(combined["order_id"])

    def test_newest_snapshot_wins(self, sales_files):
        if len(sales_files) < 2:
            pytest.skip("needs at least two snapshots")
        frames = [pipeline.process_file(item) for item in sales_files]
        combined = pd.concat(frames, ignore_index=True)
        result = pipeline.deduplicate(combined, ventas_mx.SPEC)

        newest = combined["source_modified_at"].max()
        shared = set(frames[0]["order_id"]) & set(frames[-1]["order_id"])
        kept = result[result["order_id"].isin(shared)]
        assert (kept["source_modified_at"] == newest).all()


class TestHeaderDetection:
    def test_picks_the_densest_row(self):
        frame = pd.DataFrame(
            [
                ["Title", None, None],
                [None, None, None],
                ["A", "B", "C"],
                [1, 2, 3],
            ]
        )
        assert find_header_row(frame) == 2

    def test_raises_when_no_header_found(self):
        with pytest.raises(ValueError, match="no header row"):
            find_header_row(pd.DataFrame([["Title", None, None], [None, None, None]]))

    def test_duplicate_names_get_suffixed(self):
        assert _deduplicate(["a", "a", "b", "a"]) == ["a", "a_1", "b", "a_2"]
