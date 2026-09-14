"""Tests for the Ventas MX validation rules."""

from __future__ import annotations

import pandas as pd
import pytest

from mx_sales import pipeline
from mx_sales.accounting import ventas_mx as rules
from mx_sales.reports import facturacion, ventas_mx


def make_row(**overrides) -> dict:
    """A row that reconciles: every component zero, total zero."""
    row = {column: 0.0 for column in rules.COMPONENT_COLUMNS}
    row[rules.TOTAL_COLUMN] = 0.0
    row.update(
        {
            "order_id": "1",
            "order_status": "Entregado",
            "sold_at": pd.Timestamp("2026-08-03 09:24"),
            "row_role": ventas_mx.STANDALONE,
            "package_id": None,
            "units": 1,
            "unit_price_mxn": 0.0,
            "sku": "SKU-1",
        }
    )
    row.update(overrides)
    return row


class TestRule1:
    def test_matching_row_reconciles(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=100.0, total_mxn=100.0)])
        result = rules.check_totals(frame)

        assert result["components_sum_mxn"][0] == 100.0
        assert result["total_diff_mxn"][0] == 0.0
        assert bool(result["total_matches"][0]) is True

    def test_every_component_is_summed(self):
        frame = pd.DataFrame(
            [make_row(**{c: 1.0 for c in rules.COMPONENT_COLUMNS},
                      total_mxn=float(len(rules.COMPONENT_COLUMNS)))]
        )
        result = rules.check_totals(frame)
        assert result["components_sum_mxn"][0] == float(len(rules.COMPONENT_COLUMNS))
        assert bool(result["total_matches"][0]) is True

    def test_hidden_charge_is_reported_as_a_difference(self):
        # The real case from the export: H + I + K exceeds P by 30 MXN that appears
        # in no column of the report.
        frame = pd.DataFrame(
            [
                make_row(
                    product_revenue_mxn=878.80,
                    sale_fee_and_taxes_mxn=-172.92,
                    shipping_cost_mxn=-88.50,
                    total_mxn=587.38,
                )
            ]
        )
        result = rules.check_totals(frame)

        assert result["components_sum_mxn"][0] == pytest.approx(617.38)
        assert result["total_diff_mxn"][0] == pytest.approx(30.00)
        assert bool(result["total_matches"][0]) is False

    def test_difference_is_signed(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=100.0, total_mxn=130.0)])
        result = rules.check_totals(frame)
        assert result["total_diff_mxn"][0] == pytest.approx(-30.0)

    def test_sub_cent_drift_still_counts_as_matching(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=100.001, total_mxn=100.0)])
        assert bool(rules.check_totals(frame)["total_matches"][0]) is True

    def test_a_full_cent_does_not(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=100.01, total_mxn=100.0)])
        assert bool(rules.check_totals(frame)["total_matches"][0]) is False

    def test_row_without_money_is_not_checked(self):
        # Package child rows carry no money at all; they must read as "nothing to
        # compare", not as a reconciling 0.00.
        frame = pd.DataFrame([{c: None for c in rules.COMPONENT_COLUMNS} | {"total_mxn": None}])
        result = rules.check_totals(frame)

        assert pd.isna(result["components_sum_mxn"][0])
        assert pd.isna(result["total_matches"][0])

    def test_partially_populated_row_is_still_checked(self):
        frame = pd.DataFrame(
            [{c: None for c in rules.COMPONENT_COLUMNS} | {"product_revenue_mxn": 50.0,
                                                           "total_mxn": 50.0}]
        )
        result = rules.check_totals(frame)
        assert bool(result["total_matches"][0]) is True

    def test_missing_column_raises(self):
        with pytest.raises(ValueError, match="columns missing"):
            rules.check_totals(pd.DataFrame({"product_revenue_mxn": [1.0]}))

    def test_original_columns_are_preserved(self):
        frame = pd.DataFrame([make_row(order_id="X")])
        result = rules.check_totals(frame)
        assert result["order_id"][0] == "X"
        assert set(frame.columns) <= set(result.columns)


def make_package(prices: list[float], units: list[int] | None = None, **parent_money) -> pd.DataFrame:
    """A package parent row followed by its product rows, shaped like the real export."""
    counts = units or [1] * len(prices)
    parent = make_row(
        order_id="PKG",
        order_status=f"Paquete de {len(prices)} productos",
        row_role=ventas_mx.PACKAGE_PARENT,
        package_id="PKG",
        units=None,
        unit_price_mxn=None,
        sku=None,
        **{c: None for c in rules.COMPONENT_COLUMNS},
        **{"total_mxn": None},
    )
    parent["product_revenue_mxn"] = sum(p * u for p, u in zip(prices, counts))
    parent.update(parent_money)

    children = [
        make_row(
            order_id=f"C{i}",
            order_status="Entregado",
            row_role=ventas_mx.PACKAGE_CHILD,
            package_id="PKG",
            units=count,
            unit_price_mxn=price,
            sku=f"SKU-{i}",
            **{c: None for c in rules.COMPONENT_COLUMNS},
            **{"total_mxn": None},
        )
        for i, (price, count) in enumerate(zip(prices, counts))
    ]
    return pd.DataFrame([parent, *children])


def products(frame: pd.DataFrame) -> pd.DataFrame:
    """The product rows of an allocated package; rule 2 leaves the package row in place."""
    return frame[frame["row_role"] == ventas_mx.PACKAGE_CHILD]


class TestRule2:
    def test_products_receive_the_revenue_they_contributed(self):
        result = rules.allocate_packages(make_package([193.51, 154.23]))
        assert products(result)["product_revenue_mxn"].tolist() == [193.51, 154.23]

    def test_every_allocated_column_sums_back_to_the_package(self):
        # No refund here: a package carrying one goes down rule 3's path instead.
        frame = make_package(
            [193.51, 154.23],
            sale_fee_and_taxes_mxn=-81.90,
            shipping_cost_mxn=-68.00,
            shipping_revenue_mxn=12.34,
            total_mxn=197.84,
        )
        result = products(rules.allocate_packages(frame))
        for column, expected in [
            ("product_revenue_mxn", 347.74),
            ("sale_fee_and_taxes_mxn", -81.90),
            ("shipping_cost_mxn", -68.00),
            ("shipping_revenue_mxn", 12.34),
            ("total_mxn", 197.84),
        ]:
            assert result[column].sum() == pytest.approx(expected, abs=0.005), column

    def test_split_is_proportional_not_equal(self):
        frame = make_package([300.0, 100.0], sale_fee_and_taxes_mxn=-80.0)
        result = products(rules.allocate_packages(frame))
        assert result["sale_fee_and_taxes_mxn"].tolist() == [-60.0, -20.0]

    def test_rounding_residue_is_not_lost(self):
        # 100 / 3 cannot be split into cents evenly; the pieces must still total 100.
        frame = make_package([1.0, 1.0, 1.0], total_mxn=100.0)
        result = products(rules.allocate_packages(frame))
        assert result["total_mxn"].sum() == pytest.approx(100.0, abs=0.005)
        assert sorted(result["total_mxn"].tolist()) == [33.33, 33.33, 33.34]

    def test_quantity_above_one_widens_the_share(self):
        frame = make_package([100.0, 100.0], units=[3, 1], sale_fee_and_taxes_mxn=-80.0)
        result = products(rules.allocate_packages(frame))
        assert result["sale_fee_and_taxes_mxn"].tolist() == [-60.0, -20.0]

    def test_null_parent_columns_stay_null(self):
        # The export leaves J, L, M and N empty on most packages.
        result = products(rules.allocate_packages(make_package([10.0, 10.0])))
        assert result["shipping_revenue_mxn"].isna().all()

    def test_allocated_rows_are_flagged(self):
        result = products(rules.allocate_packages(make_package([10.0, 10.0])))
        assert result["allocated_from_package"].all()

    def test_standalone_rows_are_untouched(self):
        frame = pd.DataFrame([make_row(order_id="S", product_revenue_mxn=50.0, total_mxn=50.0)])
        result = rules.allocate_packages(frame)
        assert len(result) == 1
        assert result["product_revenue_mxn"][0] == 50.0
        assert not result["allocated_from_package"][0]

    def test_package_without_unit_prices_is_left_alone(self):
        frame = make_package([0.0, 0.0], total_mxn=100.0)
        result = rules.allocate_packages(frame)
        # Nothing to divide on, so the package row survives for manual review.
        assert len(result) == 3
        assert not result["allocated_from_package"].any()

    def test_allocation_leaves_the_package_row_for_rule_4(self):
        result = rules.allocate_packages(make_package([10.0, 10.0]))
        assert len(result) == 3
        assert result["package_split"].sum() == 1

    def test_packages_helper_exposes_what_was_removed(self):
        frame = make_package([10.0, 10.0])
        assert len(rules.packages(frame)) == 1


def make_returned_package(
    prices: list[float], statuses: list[str], refund: float, **parent_money
) -> pd.DataFrame:
    """A package carrying a refund in O, with per-product statuses saying who came back."""
    frame = make_package(prices, **parent_money)
    frame.loc[0, rules.REFUND_COLUMN] = refund
    for position, status in enumerate(statuses, start=1):
        frame.loc[position, "order_status"] = status
    return frame


class TestRule3:
    def test_returned_product_is_excluded_from_the_basis(self):
        # A survives, B came back: A takes the whole remaining share rather than half.
        frame = make_returned_package(
            [200.0, 100.0],
            ["Entregado", "Devolución finalizada con reembolso al comprador"],
            refund=-100.0,
            sale_fee_and_taxes_mxn=-70.0,
            total_mxn=130.0,
        )
        result = rules.allocate_packages(frame)
        survivor = result[result["sku"] == "SKU-0"].iloc[0]
        returned = result[result["sku"] == "SKU-1"].iloc[0]

        assert survivor["sale_fee_and_taxes_mxn"] == pytest.approx(-70.0)
        assert survivor["total_mxn"] == pytest.approx(130.0)
        assert bool(returned["is_returned_item"]) is True
        assert pd.isna(returned["total_mxn"])

    def test_refund_is_folded_into_revenue(self):
        # H = 300, O = -100, so the revenue left to share is 200.
        frame = make_returned_package(
            [200.0, 100.0],
            ["Entregado", "Cancelada por el comprador"],
            refund=-100.0,
        )
        result = rules.allocate_packages(frame)
        survivor = result[result["sku"] == "SKU-0"].iloc[0]
        assert survivor["product_revenue_mxn"] == pytest.approx(200.0)

    def test_refund_is_not_deducted_twice(self):
        frame = make_returned_package(
            [200.0, 100.0], ["Entregado", "Cancelada por el comprador"], refund=-100.0
        )
        result = rules.allocate_packages(frame)
        survivor = result[result["sku"] == "SKU-0"].iloc[0]
        # O was absorbed into H, so it must not also appear on the surviving product.
        assert pd.isna(survivor[rules.REFUND_COLUMN])

    def test_rule_1_still_reconciles_after_a_refund_is_folded_in(self):
        frame = make_returned_package(
            [200.0, 100.0],
            ["Entregado", "Cancelada por el comprador"],
            refund=-100.0,
            sale_fee_and_taxes_mxn=-70.0,
            total_mxn=130.0,
        )
        result = rules.validate(frame)
        survivor = result[result["sku"] == "SKU-0"].iloc[0]
        assert bool(survivor["total_matches"]) is True

    def test_two_survivors_split_between_themselves(self):
        frame = make_returned_package(
            [200.0, 200.0, 100.0],
            ["Entregado", "Entregado", "Cancelada por el comprador"],
            refund=-100.0,
            sale_fee_and_taxes_mxn=-80.0,
        )
        result = rules.allocate_packages(frame)
        survivors = result[result["allocated_from_package"].fillna(False)]
        assert len(survivors) == 2
        assert survivors["sale_fee_and_taxes_mxn"].tolist() == [-40.0, -40.0]

    def test_fully_returned_package_keeps_its_row(self):
        frame = make_returned_package(
            [179.4, 179.4],
            ["Paquete cancelado por Mercado Libre", "Paquete cancelado por Mercado Libre"],
            refund=-274.30,
            total_mxn=-64.0,
        )
        result = rules.allocate_packages(frame)

        assert len(result) == 3
        assert (result["package_note"] == rules.FULLY_RETURNED).all()
        # The payout must survive into the accounts.
        parent = result[result["row_role"] == ventas_mx.PACKAGE_PARENT].iloc[0]
        assert parent["total_mxn"] == pytest.approx(-64.0)

    def test_money_kept_by_the_seller_is_not_a_return(self):
        # "Te dimos el dinero" books no refund, so the package allocates normally.
        frame = make_package([179.4, 179.4], sale_fee_and_taxes_mxn=-84.50)
        frame.loc[1, "order_status"] = "Mediación finalizada. Te dimos el dinero."
        result = rules.allocate_packages(frame)

        assert not result["is_returned_item"].any()
        assert products(result)["allocated_from_package"].all()

    def test_refund_with_no_identifiable_product_is_left_alone(self):
        frame = make_returned_package(
            [200.0, 100.0], ["Entregado", "Entregado"], refund=-100.0
        )
        result = rules.allocate_packages(frame)

        assert len(result) == 3
        assert (result["package_note"] == rules.REFUND_WITHOUT_IDENTIFIABLE_ITEM).all()
        assert not result["allocated_from_package"].any()

    def test_a_package_without_a_refund_is_unaffected_by_rule_3(self):
        result = rules.allocate_packages(make_package([157.08, 157.08]))
        assert not result["is_returned_item"].any()
        assert result["package_note"].isna().all()


class TestRule4:
    def test_split_package_row_is_removed(self):
        result = rules.drop_package_rows(rules.allocate_packages(make_package([157.08, 157.08])))
        assert len(result) == 2
        assert (result["row_role"] == ventas_mx.PACKAGE_CHILD).all()

    def test_unsplit_package_row_survives(self):
        # Rule 3 left this one intact, so it still holds the only copy of its money.
        frame = make_returned_package(
            [179.4, 179.4],
            ["Paquete cancelado por Mercado Libre", "Paquete cancelado por Mercado Libre"],
            refund=-274.30,
            total_mxn=-64.0,
        )
        result = rules.drop_package_rows(rules.allocate_packages(frame))
        assert len(result) == 3
        parent = result[result["row_role"] == ventas_mx.PACKAGE_PARENT].iloc[0]
        assert parent["total_mxn"] == pytest.approx(-64.0)

    def test_standalone_rows_are_untouched_by_rule_4(self):
        frame = pd.DataFrame([make_row(order_id="S", product_revenue_mxn=50.0, total_mxn=50.0)])
        assert len(rules.drop_package_rows(rules.allocate_packages(frame))) == 1

    def test_no_money_is_lost_when_the_package_row_goes(self):
        frame = make_package([193.51, 154.23], sale_fee_and_taxes_mxn=-81.90, total_mxn=197.84)
        before = float(frame.loc[0, "total_mxn"])
        after = rules.drop_package_rows(rules.allocate_packages(frame))
        assert after["total_mxn"].sum() == pytest.approx(before, abs=0.005)

    def test_is_a_no_op_on_a_frame_that_never_saw_rule_2(self):
        frame = pd.DataFrame([make_row()])
        assert len(rules.drop_package_rows(frame)) == 1


class TestRule5:
    def test_tax_is_105_percent_of_the_amount_net_of_iva(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=878.80)])
        result = rules.add_tax(frame)
        # 878.80 / 1.16 = 757.59 net; 10.5% of that is 79.55, booked as a cost.
        assert result[rules.TAX_COLUMN][0] == pytest.approx(-79.55, abs=0.005)

    def test_tax_is_negative_because_it_is_a_cost(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=633.36)])
        assert rules.add_tax(frame)[rules.TAX_COLUMN][0] < 0

    def test_tax_is_rounded_to_cents(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=3597.67)])
        value = rules.add_tax(frame)[rules.TAX_COLUMN][0]
        assert value == round(float(value), 2)

    def test_no_revenue_means_no_tax(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=None)])
        assert pd.isna(rules.add_tax(frame)[rules.TAX_COLUMN][0])

    def test_tax_follows_the_split_not_the_package(self):
        # A product out of a package is taxed on its own share, not the package total.
        frame = make_package([300.0, 100.0])
        result = rules.validate(frame)
        assert result[rules.TAX_COLUMN].tolist() == [
            pytest.approx(-27.16, abs=0.005),
            pytest.approx(-9.05, abs=0.005),
        ]

    def test_tax_sums_to_the_package_tax(self):
        frame = make_package([300.0, 100.0])
        result = rules.validate(frame)
        whole = round(400.0 / 1.16 * 0.105, 2)
        assert result[rules.TAX_COLUMN].sum() == pytest.approx(-whole, abs=0.02)


class TestRule6:
    def test_fee_is_column_i_less_the_tax(self):
        frame = rules.add_tax(
            pd.DataFrame([make_row(product_revenue_mxn=878.80, sale_fee_and_taxes_mxn=-172.92)])
        )
        result = rules.add_platform_fee(frame)
        # |−172.92| − 79.55 = 93.37, the "Cargo por venta" ML bills; booked as a cost.
        assert result[rules.PLATFORM_FEE_COLUMN][0] == pytest.approx(-93.37, abs=0.005)

    def test_fee_and_tax_rebuild_column_i(self):
        frame = rules.add_tax(
            pd.DataFrame([make_row(product_revenue_mxn=633.36, sale_fee_and_taxes_mxn=-127.00)])
        )
        result = rules.add_platform_fee(frame).iloc[0]
        rebuilt = result[rules.PLATFORM_FEE_COLUMN] + result[rules.TAX_COLUMN]
        assert rebuilt == pytest.approx(-abs(result[rules.FEE_COLUMN]), abs=0.005)

    def test_fee_is_negative_like_the_tax(self):
        frame = rules.add_tax(
            pd.DataFrame([make_row(product_revenue_mxn=633.36, sale_fee_and_taxes_mxn=-127.00)])
        )
        assert rules.add_platform_fee(frame)[rules.PLATFORM_FEE_COLUMN][0] < 0

    def test_no_fee_charged_means_the_tax_is_the_whole_of_it(self):
        frame = rules.add_tax(
            pd.DataFrame([make_row(product_revenue_mxn=100.0, sale_fee_and_taxes_mxn=None)])
        )
        assert pd.isna(rules.add_platform_fee(frame)[rules.PLATFORM_FEE_COLUMN][0])

    def test_running_before_rule_5_is_refused(self):
        frame = pd.DataFrame([make_row(product_revenue_mxn=100.0)])
        with pytest.raises(ValueError, match="run rule 5 first"):
            rules.add_platform_fee(frame)

    def test_fee_follows_the_package_split(self):
        frame = make_package([300.0, 100.0], sale_fee_and_taxes_mxn=-80.0)
        result = rules.validate(frame)
        # Each product's fee comes off its own share of I, less its own tax.
        assert result[rules.PLATFORM_FEE_COLUMN].sum() == pytest.approx(
            -80.0 - result[rules.TAX_COLUMN].sum(), abs=0.02
        )


class TestRule7:
    def test_period_is_yyyymm(self):
        frame = pd.DataFrame([make_row(sold_at=pd.Timestamp("2026-08-03 09:24"))])
        assert rules.add_period(frame)[rules.PERIOD_COLUMN][0] == "202608"

    def test_january_keeps_its_leading_zero(self):
        frame = pd.DataFrame([make_row(sold_at=pd.Timestamp("2026-01-31"))])
        assert rules.add_period(frame)[rules.PERIOD_COLUMN][0] == "202601"

    def test_period_sorts_chronologically_as_text(self):
        frame = pd.DataFrame(
            [
                make_row(sold_at=pd.Timestamp("2026-12-01")),
                make_row(sold_at=pd.Timestamp("2027-01-01")),
                make_row(sold_at=pd.Timestamp("2026-02-01")),
            ]
        )
        periods = rules.add_period(frame)[rules.PERIOD_COLUMN]
        assert sorted(periods) == ["202602", "202612", "202701"]

    def test_missing_sale_date_gives_no_period(self):
        frame = pd.DataFrame([make_row(sold_at=pd.NaT)])
        assert pd.isna(rules.add_period(frame)[rules.PERIOD_COLUMN][0])


class TestMonthlySummary:
    @pytest.fixture
    def summary(self):
        frame = rules.validate(
            pd.DataFrame(
                [
                    make_row(order_id="A", sold_at=pd.Timestamp("2026-07-05"),
                             product_revenue_mxn=100.0, sale_fee_and_taxes_mxn=-20.0,
                             total_mxn=80.0, units=1),
                    make_row(order_id="B", sold_at=pd.Timestamp("2026-07-20"),
                             product_revenue_mxn=200.0, sale_fee_and_taxes_mxn=-40.0,
                             total_mxn=160.0, units=2),
                    make_row(order_id="C", sold_at=pd.Timestamp("2026-08-02"),
                             product_revenue_mxn=50.0, sale_fee_and_taxes_mxn=-10.0,
                             total_mxn=40.0, units=1),
                ]
            )
        )
        return rules.monthly_summary(frame)

    def test_one_row_per_month(self, summary):
        assert summary[rules.PERIOD_COLUMN].tolist() == ["202607", "202608"]

    def test_counts_and_totals(self, summary):
        july = summary.iloc[0]
        assert july["orders"] == 2
        assert july["units"] == 3
        assert july["revenue_mxn"] == pytest.approx(300.0)
        assert july["payout_mxn"] == pytest.approx(240.0)

    def test_adds_the_period_itself_when_absent(self):
        frame = rules.validate(pd.DataFrame([make_row(sold_at=pd.Timestamp("2026-07-05"))]))
        without = frame.drop(columns=[rules.PERIOD_COLUMN])
        assert rules.monthly_summary(without)[rules.PERIOD_COLUMN].tolist() == ["202607"]


class TestRule8:
    def test_identity_columns_are_removed(self):
        frame = pd.DataFrame([{**make_row(), "buyer_name": "X", "invoice_tax_id": "Y",
                               "buyer_state": "Yucatán", "cfdi_code": "S01"}])
        result = rules.drop_buyer_identity(frame)
        for column in ["buyer_name", "invoice_tax_id", "buyer_state", "cfdi_code"]:
            assert column not in result.columns

    def test_money_columns_are_untouched(self):
        frame = pd.DataFrame([{**make_row(product_revenue_mxn=100.0), "buyer_name": "X"}])
        result = rules.drop_buyer_identity(frame)
        assert result["product_revenue_mxn"][0] == 100.0
        assert len(result) == 1

    def test_absent_columns_are_not_an_error(self):
        frame = pd.DataFrame([make_row()])
        assert len(rules.drop_buyer_identity(frame).columns) == len(frame.columns)

    def test_the_set_covers_both_sheet_groups(self):
        # Yellow "Facturación al comprador" and green "Compradores".
        assert "invoice_name" in ventas_mx.IDENTITY_COLUMNS
        assert "buyer_name" in ventas_mx.IDENTITY_COLUMNS

    def test_composite_derived_columns_are_included(self):
        # These are split out of identity source fields, so they must go too.
        for column in ["invoice_tax_id", "cfdi_code", "cfdi_description", "invoice_attached"]:
            assert column in ventas_mx.IDENTITY_COLUMNS

    def test_no_personal_data_survives_validate(self):
        frame = pd.DataFrame([{**make_row(), "buyer_name": "Sonia", "buyer_tax_id": "XAXX01",
                               "buyer_address": "Calle 1", "invoice_name": "Sonia"}])
        result = rules.validate(frame)
        assert not [c for c in result.columns if c in ventas_mx.IDENTITY_COLUMNS]


class TestRule9:
    def test_columns_e_and_f_are_removed(self):
        frame = pd.DataFrame([{**make_row(), "is_multi_product_package": True,
                               "belongs_to_kit": False}])
        result = rules.drop_unnecessary_columns(frame)
        assert "is_multi_product_package" not in result.columns
        assert "belongs_to_kit" not in result.columns

    def test_everything_else_survives(self):
        frame = pd.DataFrame([{**make_row(product_revenue_mxn=100.0),
                               "is_multi_product_package": True}])
        result = rules.drop_unnecessary_columns(frame)
        assert result["product_revenue_mxn"][0] == 100.0
        assert len(result) == 1

    def test_absent_columns_are_not_an_error(self):
        frame = pd.DataFrame([make_row()])
        assert len(rules.drop_unnecessary_columns(frame).columns) == len(frame.columns)

    def test_validate_drops_them(self):
        frame = pd.DataFrame([{**make_row(), "is_multi_product_package": True,
                               "belongs_to_kit": False}])
        result = rules.validate(frame)
        assert not [c for c in rules.UNNECESSARY_COLUMNS if c in result.columns]


class TestRule10:
    def test_zero_payout_row_is_excluded(self):
        frame = rules.validate(
            pd.DataFrame([make_row(product_revenue_mxn=659.0,
                                   cancellations_and_refunds_mxn=-419.29, total_mxn=0.0,
                                   order_status="Cancelada por el comprador")])
        )
        assert bool(frame[rules.COUNTS_COLUMN][0]) is False
        assert frame[rules.EXCLUSION_REASON_COLUMN][0] == rules.CANCELLED

    def test_reason_is_read_from_the_status(self):
        for status, expected in [
            ("Cancelada por el comprador", rules.CANCELLED),
            ("Devolución revisada. Solicita el retiro del producto", rules.RETURNED),
            ("Mediación finalizada con reembolso al comprador", rules.MEDIATION),
            ("Cancelaste la venta", rules.CANCELLED),
            ("Entregado", rules.ZERO_PAYOUT),
        ]:
            frame = rules.flag_revenue_exclusions(
                pd.DataFrame([make_row(total_mxn=0.0, order_status=status)])
            )
            assert frame[rules.EXCLUSION_REASON_COLUMN][0] == expected, status

    def test_partial_refund_that_still_paid_out_is_kept(self):
        # revenue 849, refund -113.75, payout 476.51 -- the sale did earn.
        frame = rules.flag_revenue_exclusions(
            pd.DataFrame([make_row(product_revenue_mxn=849.0,
                                   cancellations_and_refunds_mxn=-113.75, total_mxn=476.51,
                                   order_status="Mediación finalizada con reembolso al comprador")])
        )
        assert bool(frame[rules.COUNTS_COLUMN][0]) is True
        assert pd.isna(frame[rules.EXCLUSION_REASON_COLUMN][0])

    def test_null_payout_is_kept(self):
        # Null means nothing to judge, which is not the same as a real zero.
        frame = rules.flag_revenue_exclusions(pd.DataFrame([make_row(total_mxn=None)]))
        assert bool(frame[rules.COUNTS_COLUMN][0]) is True

    def test_a_loss_making_sale_is_kept(self):
        frame = rules.flag_revenue_exclusions(pd.DataFrame([make_row(total_mxn=-667.50)]))
        assert bool(frame[rules.COUNTS_COLUMN][0]) is True

    def test_excluded_rows_lists_only_the_excluded(self):
        frame = rules.validate(
            pd.DataFrame([
                make_row(order_id="A", product_revenue_mxn=100.0, total_mxn=100.0),
                make_row(order_id="B", product_revenue_mxn=200.0, total_mxn=0.0),
            ])
        )
        assert rules.excluded_rows(frame)["order_id"].tolist() == ["B"]

    def test_monthly_summary_reconciles_counted_and_excluded(self):
        frame = rules.validate(
            pd.DataFrame([
                make_row(order_id="A", product_revenue_mxn=100.0, total_mxn=100.0),
                make_row(order_id="B", product_revenue_mxn=200.0, total_mxn=0.0),
            ])
        )
        row = rules.monthly_summary(frame).iloc[0]
        assert row["revenue_mxn"] == pytest.approx(100.0)
        assert row["excluded_orders"] == 1
        assert row["excluded_revenue_mxn"] == pytest.approx(200.0)
        assert row["gross_revenue_mxn"] == pytest.approx(300.0)

    def test_a_period_with_nothing_excluded_reports_zero(self):
        frame = rules.validate(
            pd.DataFrame([make_row(product_revenue_mxn=100.0, total_mxn=100.0)])
        )
        row = rules.monthly_summary(frame).iloc[0]
        assert row["excluded_orders"] == 0
        assert row["excluded_revenue_mxn"] == pytest.approx(0.0)


class TestSignConvention:
    def test_every_cost_column_is_negative(self):
        frame = rules.validate(
            pd.DataFrame([make_row(product_revenue_mxn=878.80, sale_fee_and_taxes_mxn=-172.92,
                                   shipping_cost_mxn=-88.50,
                                   cancellations_and_refunds_mxn=-10.0, total_mxn=607.38)])
        )
        row = rules.monthly_summary(frame).iloc[0]
        for column in ["platform_fee_mxn", "tax_withheld_mxn", "shipping_cost_mxn", "refunds_mxn"]:
            assert row[column] < 0, column

    def test_the_monthly_columns_add_up_to_the_payout(self):
        frame = rules.validate(
            pd.DataFrame([make_row(product_revenue_mxn=878.80, sale_fee_and_taxes_mxn=-172.92,
                                   shipping_revenue_mxn=20.0, shipping_cost_mxn=-88.50,
                                   cancellations_and_refunds_mxn=-10.0, total_mxn=627.38)])
        )
        row = rules.monthly_summary(frame).iloc[0]
        rebuilt = (row["revenue_mxn"] + row["platform_fee_mxn"] + row["tax_withheld_mxn"]
                   + row["shipping_revenue_mxn"] + row["shipping_cost_mxn"] + row["refunds_mxn"])
        assert rebuilt == pytest.approx(row["payout_mxn"], abs=0.02)


class TestSkuSummary:
    def test_one_row_per_period_and_sku(self):
        frame = rules.validate(
            pd.DataFrame([
                make_row(order_id="A", sku="X", sold_at=pd.Timestamp("2026-07-05"),
                         product_revenue_mxn=100.0, total_mxn=100.0),
                make_row(order_id="B", sku="X", sold_at=pd.Timestamp("2026-08-05"),
                         product_revenue_mxn=200.0, total_mxn=200.0),
                make_row(order_id="C", sku="Y", sold_at=pd.Timestamp("2026-08-06"),
                         product_revenue_mxn=50.0, total_mxn=50.0),
            ])
        )
        out = rules.sku_summary(frame)
        assert len(out) == 3
        assert set(zip(out[rules.PERIOD_COLUMN], out["sku"])) == {
            ("202607", "X"), ("202608", "X"), ("202608", "Y"),
        }

    def test_excluded_rows_never_reach_the_tab(self):
        frame = rules.validate(
            pd.DataFrame([
                make_row(order_id="A", sku="X", product_revenue_mxn=100.0, total_mxn=100.0),
                make_row(order_id="B", sku="X", product_revenue_mxn=900.0, total_mxn=0.0),
            ])
        )
        row = rules.sku_summary(frame).iloc[0]
        assert row["revenue_mxn"] == pytest.approx(100.0)
        assert row["orders"] == 1

    def test_settlement_unit_price_is_payout_over_units(self):
        frame = rules.validate(
            pd.DataFrame([
                make_row(order_id="A", sku="X", units=2, product_revenue_mxn=200.0,
                         total_mxn=150.0),
            ])
        )
        row = rules.sku_summary(frame).iloc[0]
        assert row["settlement_unit_price_mxn"] == pytest.approx(75.0)

    def test_tax_is_rounded_once_on_the_group(self):
        # Three rows of 0.05 each round to 0.00 individually but 0.02 in aggregate.
        frame = rules.validate(
            pd.DataFrame([
                make_row(order_id=f"O{i}", sku="X", product_revenue_mxn=0.05, total_mxn=0.05)
                for i in range(3)
            ])
        )
        per_row = frame[rules.TAX_COLUMN].sum()
        grouped = rules.sku_summary(frame).iloc[0]["tax_withheld_mxn"]
        assert grouped == pytest.approx(-round(0.15 / 1.16 * 0.105, 2), abs=0.005)
        assert grouped != per_row

    def test_a_sku_with_no_units_does_not_divide_by_zero(self):
        frame = rules.validate(
            pd.DataFrame([make_row(sku="X", units=0, product_revenue_mxn=10.0, total_mxn=10.0)])
        )
        assert pd.isna(rules.sku_summary(frame).iloc[0]["settlement_unit_price_mxn"])


def make_charges(rows: list[dict]) -> pd.DataFrame:
    """Facturacion rows shaped like the cleaned output."""
    base = {"charge_id": None, "listing_id": "MLM1", "sold_at": pd.Timestamp("2026-08-05"),
            "fee_category": "commission", "charge_amount_mxn": 0.0, "charge_status": pd.NA,
            "order_id": "1"}
    frame = pd.DataFrame([{**base, **r} for r in rows])
    frame["charge_status"] = frame["charge_status"].astype("string")
    return frame


class TestFeesBySku:
    def test_charges_reach_a_sku_through_the_listing(self):
        sales = pd.DataFrame([make_row(sku="X", listing_id="MLM1", units=1)])
        fees = facturacion.fees_by_sku(
            make_charges([{"charge_amount_mxn": 100.0}]), rules.listing_sku_map(sales)
        )
        assert fees.iloc[0]["sku"] == "X"
        # ML publishes charges positive; ours are costs.
        assert fees.iloc[0]["fee_commission_mxn"] == pytest.approx(-100.0)

    def test_ordinary_charges_survive_the_void_filter(self):
        # Regression: a normal charge has a null status, and `.ne()` on a nullable
        # column yields NA there -- filtering on it dropped every ordinary charge.
        fees = facturacion.fees_by_sku(
            make_charges([
                {"charge_amount_mxn": 100.0, "charge_status": pd.NA},
                {"charge_amount_mxn": 50.0, "charge_status": "Anulado en factura"},
            ]),
            pd.Series({"MLM1": "X"}),
        )
        assert fees.iloc[0]["fee_commission_mxn"] == pytest.approx(-150.0)

    def test_credit_note_voids_are_dropped(self):
        fees = facturacion.fees_by_sku(
            make_charges([
                {"charge_amount_mxn": 100.0},
                {"charge_amount_mxn": 40.0, "charge_status": facturacion.CREDIT_NOTE_VOID},
            ]),
            pd.Series({"MLM1": "X"}),
        )
        assert fees.iloc[0]["fee_commission_mxn"] == pytest.approx(-100.0)

    def test_account_level_charges_are_left_out(self):
        # Storage and advertising name no listing, so they belong to no product.
        fees = facturacion.fees_by_sku(
            make_charges([{"fee_category": "storage", "listing_id": None},
                          {"fee_category": "advertising", "listing_id": None}]),
            pd.Series({"MLM1": "X"}),
        )
        assert fees.empty

    def test_no_charges_returns_an_empty_frame(self):
        assert facturacion.fees_by_sku(pd.DataFrame(), pd.Series(dtype="object")).empty


class TestListingSkuMap:
    def test_busiest_sku_wins_a_shared_listing(self):
        sales = pd.DataFrame([
            make_row(order_id="A", sku="GY", listing_id="MLM1", units=6),
            make_row(order_id="B", sku="PK", listing_id="MLM1", units=1),
        ])
        assert rules.listing_sku_map(sales)["MLM1"] == "GY"

    def test_missing_columns_give_an_empty_map(self):
        assert rules.listing_sku_map(pd.DataFrame({"sku": ["X"]})).empty


class TestAccountLevelCosts:
    def make(self, rows):
        base = {"listing_id": None, "charge_month": "202608", "charge_status": pd.NA,
                "fee_category": "advertising", "charge_amount_mxn": 0.0}
        frame = pd.DataFrame([{**base, **r} for r in rows])
        frame["charge_status"] = frame["charge_status"].astype("string")
        return frame

    def test_advertising_and_storage_are_reported_negative(self):
        out = facturacion.account_level_costs(self.make([
            {"fee_category": "advertising", "charge_amount_mxn": 100.0},
            {"fee_category": "storage", "charge_amount_mxn": 40.0},
        ]))
        row = out.iloc[0]
        assert row["advertising_mxn"] == pytest.approx(-100.0)
        assert row["storage_mxn"] == pytest.approx(-40.0)

    def test_pickup_counts_as_storage(self):
        out = facturacion.account_level_costs(self.make([
            {"fee_category": "storage", "charge_amount_mxn": 663.0},
            {"fee_category": "pickup", "charge_amount_mxn": 4497.66},
        ]))
        assert out.iloc[0]["storage_mxn"] == pytest.approx(-5160.66)

    def test_credit_note_voids_are_dropped(self):
        # The real shape: advertising bills 6,394.60 in total, of which 117.26 is voided
        # by a credit note, leaving the 6,277.34 the accountant publishes.
        out = facturacion.account_level_costs(self.make([
            {"charge_amount_mxn": 6277.34},
            {"charge_amount_mxn": 117.26, "charge_status": facturacion.CREDIT_NOTE_VOID},
        ]))
        assert out.iloc[0]["advertising_mxn"] == pytest.approx(-6277.34)

    def test_charges_tied_to_a_listing_are_not_account_level(self):
        out = facturacion.account_level_costs(self.make([
            {"listing_id": "MLM1", "charge_amount_mxn": 500.0},
        ]))
        assert out.empty or pd.isna(out.iloc[0]["advertising_mxn"])

    def test_no_charges_gives_an_empty_frame(self):
        assert facturacion.account_level_costs(pd.DataFrame()).empty


class TestMonthlyAccountCosts:
    def frame(self):
        return rules.validate(pd.DataFrame([
            make_row(order_id="A", sold_at=pd.Timestamp("2026-08-05"),
                     product_revenue_mxn=1000.0, total_mxn=800.0),
        ]))

    def test_costs_join_on_the_charge_month(self):
        charges = pd.DataFrame([{"listing_id": None, "charge_month": "202608",
                                 "charge_status": pd.NA, "fee_category": "advertising",
                                 "charge_amount_mxn": 100.0}])
        charges["charge_status"] = charges["charge_status"].astype("string")
        row = rules.monthly_summary(self.frame(), charges).iloc[0]
        assert row["advertising_mxn"] == pytest.approx(-100.0)
        assert row["net_after_account_costs_mxn"] == pytest.approx(700.0)

    def test_a_month_with_no_such_charges_is_null_not_zero(self):
        row = rules.monthly_summary(self.frame(), pd.DataFrame()).iloc[0]
        assert pd.isna(row["advertising_mxn"])
        # With nothing to deduct the net equals the payout.
        assert row["net_after_account_costs_mxn"] == pytest.approx(row["payout_mxn"])

    def test_the_summary_still_works_without_charges(self):
        row = rules.monthly_summary(self.frame()).iloc[0]
        assert row["payout_mxn"] == pytest.approx(800.0)


class TestRuleOrder:
    def test_validate_removes_packages_then_checks_totals(self):
        frame = make_package(
            [50.0, 50.0], sale_fee_and_taxes_mxn=-20.0, total_mxn=80.0
        )
        result = rules.validate(frame)

        assert len(result) == 2
        # Each product now carries its own money, so rule 1 can check it.
        assert result["total_matches"].notna().all()
        assert result["total_matches"].all()

    def test_a_package_hiding_a_gap_is_caught_after_allocation(self):
        frame = make_package(
            [50.0, 50.0], sale_fee_and_taxes_mxn=-20.0, total_mxn=50.0
        )
        result = rules.validate(frame)
        assert (~result["total_matches"]).all()
        assert result["total_diff_mxn"].sum() == pytest.approx(30.0, abs=0.005)


class TestReporting:
    @pytest.fixture
    def checked(self):
        return rules.check_totals(
            pd.DataFrame(
                [
                    make_row(product_revenue_mxn=100.0, total_mxn=100.0),
                    make_row(product_revenue_mxn=100.0, total_mxn=70.0),
                    make_row(product_revenue_mxn=100.0, total_mxn=95.0),
                    {c: None for c in rules.COMPONENT_COLUMNS} | {"total_mxn": None},
                ]
            )
        )

    def test_exceptions_lists_only_failures_worst_first(self, checked):
        failed = rules.exceptions(checked)
        assert len(failed) == 2
        assert failed["total_diff_mxn"].tolist() == [30.0, 5.0]

    def test_summary_counts(self, checked):
        summary = rules.summarize(checked).iloc[0]
        assert summary["rows"] == 4
        assert summary["checked"] == 3
        assert summary["not_checked"] == 1
        assert summary["matching"] == 1
        assert summary["differing"] == 2
        assert summary["net_difference_mxn"] == pytest.approx(35.0)
        assert summary["largest_difference_mxn"] == pytest.approx(30.0)

    def test_summary_is_clean_when_everything_reconciles(self):
        checked = rules.check_totals(pd.DataFrame([make_row(product_revenue_mxn=1.0, total_mxn=1.0)]))
        summary = rules.summarize(checked).iloc[0]
        assert summary["differing"] == 0
        assert summary["net_difference_mxn"] == 0.0


@pytest.fixture(scope="module")
def real_export():
    items = [
        item
        for item in pipeline.discover()
        if item.spec.name == "ventas_mx" and "BOCINA_SM" in str(item.path)
    ]
    if not items:
        pytest.skip("no BOCINA_SM sales export available")
    return rules.check_totals(pipeline.process_file(items[-1]))


class TestAgainstTheRealExport:
    def test_every_money_row_is_checked(self, real_export):
        # A row with a total must always be comparable.
        with_total = real_export[real_export[rules.TOTAL_COLUMN].notna()]
        assert with_total["total_matches"].notna().all()

    def test_rows_without_money_are_skipped_not_failed(self, real_export):
        without = real_export[real_export[rules.TOTAL_COLUMN].isna()]
        assert not without["total_matches"].eq(False).any()

    def test_the_known_thirty_peso_gap_is_found(self, real_export):
        failed = rules.exceptions(real_export)
        assert not failed.empty
        # Every difference in this export is the same fixed 30.00 charge.
        assert failed["total_diff_mxn"].abs().max() == pytest.approx(30.0)
