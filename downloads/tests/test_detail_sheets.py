"""Tests for the two charge-detail sheets that mirror the accountant's layout."""

from __future__ import annotations

import pandas as pd
import pytest

from mx_sales.accounting import detail_sheets
from mx_sales.reports.facturacion import CREDIT_NOTE_VOID


def make_charges(rows: list[dict]) -> pd.DataFrame:
    """Facturacion rows under ML's own Spanish headers."""
    base = {
        "Fecha del cargo": pd.Timestamp("2026-08-16"),
        "Detalle": "Cargo por venta",
        "Estado del cargo": None,
        "Valor del cargo": 100.0,
        "Número de venta": 2000017459121180,
        "Número de paquete": 2000014062603581,
        "Número de publicación": "MLM1",
        "Total de la venta": 399.0,
    }
    return pd.DataFrame([{**base, **r} for r in rows])


SKUS = pd.Series({"MLM1": "SKU-A"})


class TestChargeDetail:
    def test_the_three_prepended_columns_come_first(self):
        out = detail_sheets.charge_detail(make_charges([{}]), SKUS, "MX-TA02")
        assert list(out.columns[:3]) == ["年月", "店铺", "币别"]
        assert out.iloc[0]["年月"] == "202608"
        assert out.iloc[0]["店铺"] == "MX-TA02"

    def test_the_working_block_comes_last(self):
        out = detail_sheets.charge_detail(make_charges([{}]), SKUS)
        assert list(out.columns[-9:]) == detail_sheets.ADDED_COLUMNS

    def test_ml_columns_are_kept_verbatim(self):
        out = detail_sheets.charge_detail(make_charges([{}]), SKUS)
        for column in ["Detalle", "Valor del cargo", "Número de venta"]:
            assert column in out.columns

    @pytest.mark.parametrize(
        "detalle,category,flag",
        [
            ("Cargo por venta", "销售费", "YJ"),
            ("Anulación del cargo por venta", "取消销售费", "YJ"),
            ("Cargo por envíos de Mercado Libre", "Mercado Libre 运费", "YF"),
            ("Cargo por servicio de almacenamiento Full", "仓储服务费", "CC"),
            ("Cargo por servicio de colecta Full", "仓储服务费", "CC"),
            ("Cargo por campaña de publicidad de Product Ads", "广告费用", "GG"),
            ("Cargo por mantenimiento de Mi página", "页面维护费", "Y"),
        ],
    )
    def test_category_and_flag_are_derived_from_detalle(self, detalle, category, flag):
        out = detail_sheets.charge_detail(make_charges([{"Detalle": detalle}]), SKUS)
        assert out.iloc[0]["费用项目"] == category
        assert out.iloc[0]["标志"] == flag

    def test_ids_drop_the_2000_prefix_and_its_zero(self):
        out = detail_sheets.charge_detail(make_charges([{}]), SKUS)
        assert out.iloc[0]["引用R"] == 17459121180
        assert out.iloc[0]["引用AB"] == 14062603581

    def test_sku_comes_from_the_listing(self):
        out = detail_sheets.charge_detail(make_charges([{}]), SKUS)
        assert out.iloc[0]["SKU"] == "SKU-A"

    def test_an_account_level_charge_has_no_sku(self):
        out = detail_sheets.charge_detail(
            make_charges([{"Número de publicación": None, "Número de venta": None,
                           "Detalle": "Cargo por servicio de almacenamiento Full"}]),
            SKUS,
        )
        assert pd.isna(out.iloc[0]["SKU"])
        assert pd.isna(out.iloc[0]["重次"])

    def test_tax_is_the_sale_total_net_of_iva(self):
        out = detail_sheets.charge_detail(make_charges([{"Total de la venta": 399.0}]), SKUS)
        assert out.iloc[0]["税金"] == pytest.approx(36.12, abs=0.005)

    def test_sequence_counts_charges_within_an_order(self):
        out = detail_sheets.charge_detail(
            make_charges([{}, {"Detalle": "Cargo por envíos de Mercado Libre"}, {}]), SKUS
        )
        assert out["重次"].tolist() == [1, 2, 3]

    def test_sequence_restarts_for_a_different_order(self):
        out = detail_sheets.charge_detail(
            make_charges([{}, {"Número de venta": 2000017000000000}]), SKUS
        )
        assert out["重次"].tolist() == [1, 1]

    def test_credit_note_voids_are_flagged_and_dated(self):
        out = detail_sheets.charge_detail(
            make_charges([{"Estado del cargo": CREDIT_NOTE_VOID}]), SKUS
        )
        assert out.iloc[0]["调减费用"] == "Y"
        assert out.iloc[0]["抵扣年月"] == "202608"

    def test_an_invoice_void_is_not_deducted(self):
        out = detail_sheets.charge_detail(
            make_charges([{"Estado del cargo": "Anulado en factura"}]), SKUS
        )
        assert out.iloc[0]["调减费用"] == ""

    def test_a_void_without_a_sku_is_dated_but_not_deducted(self):
        # Their two Display Ads rows behave this way.
        out = detail_sheets.charge_detail(
            make_charges([{"Estado del cargo": CREDIT_NOTE_VOID,
                           "Número de publicación": None}]),
            SKUS,
        )
        assert out.iloc[0]["调减费用"] == ""
        assert out.iloc[0]["抵扣年月"] == "202608"

    def test_no_charges_gives_an_empty_sheet(self):
        assert detail_sheets.charge_detail(pd.DataFrame(), SKUS).empty


class TestCreditNoteDetail:
    def test_the_export_is_passed_through_untouched(self):
        notes = make_charges([{"Detalle": "Anulación del cargo por venta"}])
        out = detail_sheets.credit_note_detail(notes)
        assert list(out.columns) == list(notes.columns)
        assert len(out) == len(notes)

    def test_empty_input_gives_an_empty_sheet(self):
        assert detail_sheets.credit_note_detail(pd.DataFrame()).empty
