"""Validation rules for the Ventas MX report.

Rules are added one at a time, as each is confirmed.

Rule 1 -- the components must add up
    ML publishes eight money columns, H through O, and a ninth column P it calls the
    total. Those eight should sum to P. Where they do not, ML has moved money that
    appears in no column of this report, and the difference is the amount to chase.

Rule 2 -- push package money down onto the products
    A multi-product sale is written as a "Paquete de N productos" row carrying all the
    money, followed by N product rows carrying the SKU and unit price but no money at
    all. That makes per-product figures impossible. Rule 2 shares H..P out across the
    products in proportion to what each contributes to the package.

Rule 3 -- a returned product must not distort the others
    The refund in column O is folded into H and only the products that did not come
    back divide the result, so a return cannot skew what the surviving products show.

Rule 4 -- the package row goes once it has been split
    Its money now sits on the products, so keeping it would count everything twice. A
    package rule 3 could not split keeps its row, because it still holds the only copy
    of its money.

Rule 5 -- the tax withheld on the sale
    ``H / 1.16 * 0.105``: strip the 16% IVA out of the product revenue, then take 10.5%
    of what is left.

Rule 6 -- the platform's own fee
    Column I is "Cargo por venta e impuestos" -- commission *and* tax together. Taking
    the tax from rule 5 back out leaves the commission ML actually charged.

Rule 7 -- the accounting period
    ``YYYYMM`` taken from the sale date, so rows group into months. It is the month the
    sale happened, not the month ML billed for it; those agree 98.9% of the time in this
    data but not always.

Rule 8 -- drop the buyer identity blocks
    "Facturación al comprador" (yellow in the sheet) and "Compradores" (green) describe
    who bought, not what the sale earned. None of the other rules read them, and they
    carry names, tax ids and home addresses, so they come out.

Rule 9 -- remove the unnecessary columns
    Columns that carry no accounting signal, listed in :data:`UNNECESSARY_COLUMNS`.
    Currently the two source flags E ``Paquete de varios productos`` and F ``Pertenece a
    un kit``; add to that list as more are agreed. They are dropped from the accounting
    output only -- the cleaned table still mirrors the source sheet.

Rule 10 -- a fully reversed sale does not count toward revenue
    A cancelled or returned sale keeps its revenue in the export while its payout falls
    to zero. Those rows are flagged rather than deleted, and reported beside the counted
    figures so the two reconcile to the export's gross total.

Rule 11 -- what each product sold and settled for
    The order-level rows rolled up to one row per period and SKU, with ML's own billed
    fees joined in from Reporte_Facturacion. Facturacion names no SKU, but it names the
    listing, and the sales export maps every listing to a SKU -- so no hand-maintained
    product master is needed.

They run 2, 3, 4, 1, 10, 5, 6, 7, 8, 9: the packages are split and their rows removed
first, so by the time the H..O check runs every remaining row carries its own money, the
derived columns are computed from each product's own share, and the two column-removal
rules go last so nothing earlier is deprived of a column it needed. Rule 11 is a
summary built on top of the finished rows rather than a step in that chain.

Sign convention: revenue is positive and **every cost is negative** -- the platform fee,
the withheld tax, shipping and refunds alike. That matches the accountant's workbook and
makes the money columns add up to the payout.
"""

from __future__ import annotations

import pandas as pd

from mx_sales.reports.facturacion import (
    ACCOUNT_LEVEL_COSTS,
    SKU_FEE_COLUMNS,
    account_level_costs,
    fees_by_sku,
)
from mx_sales.reports.ventas_mx import IDENTITY_COLUMNS, PACKAGE_CHILD, PACKAGE_PARENT

#: The money components the total is built from, in sheet order.
#:
#: The sheet letters differ per account. ``Costo de envío por cambio de producto`` is
#: only issued to accounts that have had product exchanges, and when it appears it sits
#: at L and pushes everything after it along -- so BOCINA runs H..O with the total at P,
#: while EWTTO runs H..P with the total at Q. Columns are matched by name, never by
#: position, so both layouts read correctly; the letters below are BOCINA's.
COMPONENT_COLUMNS = [
    "product_revenue_mxn",                # H  Ingresos por productos
    "sale_fee_and_taxes_mxn",             # I  Cargo por venta e impuestos
    "shipping_revenue_mxn",               # J  Ingresos por envío
    "shipping_cost_mxn",                  # K  Costos de envío
    "exchange_shipping_cost_mxn",         # (L on accounts that have it) Costo de envío
                                          #     por cambio de producto
    "shipping_cost_declared_mxn",         # L  Costo de envío basado en medidas y peso
    "shipping_dimension_adjustment_mxn",  # M  Cargo por diferencias en medidas y peso
    "discounts_and_bonuses_mxn",          # N  Descuentos y bonificaciones
    "cancellations_and_refunds_mxn",      # O  Anulaciones y reembolsos
]

#: The total column (P on BOCINA, Q on EWTTO).
TOTAL_COLUMN = "total_mxn"

#: Everything rule 2 shares out: the components plus the total.
ALLOCATED_COLUMNS = COMPONENT_COLUMNS + [TOTAL_COLUMN]

#: Amounts are pesos to two decimals, so anything under half a cent is equality.
TOLERANCE = 0.005

#: Rule 3. A refund booked in column O is what marks a return; the status wording is
#: only used afterwards, to work out *which* product in a package came back, because
#: product rows carry no money of their own.
REFUND_COLUMN = "cancellations_and_refunds_mxn"

#: Product statuses that mean the buyer got their money back. "Te dimos el dinero" and
#: "Descartamos el producto" are deliberately absent: the seller kept the money there,
#: and those rows book no refund at all.
RETURN_STATUS = r"devoluci|reembolso|cancelad"
KEPT_MONEY_STATUS = r"te dimos el dinero|descartamos el producto"

#: Rule 5. H is gross of IVA; the withholding is 10.5% of the amount net of it.
IVA_RATE = 0.16
WITHHOLDING_RATE = 0.105
TAX_COLUMN = "tax_withholding_mxn"

#: Rule 6. Column I, and the platform fee left once rule 5's tax comes out of it.
FEE_COLUMN = "sale_fee_and_taxes_mxn"
PLATFORM_FEE_COLUMN = "platform_fee_mxn"

#: Rule 7. The accounting period, taken from the sale date.
PERIOD_COLUMN = "sale_period"
PERIOD_SOURCE = "sold_at"

#: Rule 10. Rows whose sale was fully reversed carry the original revenue but pay out
#: nothing, so counting them overstates sales. ``total_mxn == 0`` is what the accountant
#: filters on (their 零金额 column) and it reproduces their published figure exactly.
COUNTS_COLUMN = "counts_toward_revenue"
EXCLUSION_REASON_COLUMN = "revenue_exclusion_reason"

#: Why a row was excluded, read from its status. ``ZERO_PAYOUT`` is the fallback for a
#: row that nets to nothing without saying why.
CANCELLED = "cancelled"
RETURNED = "returned"
MEDIATION = "mediation"
ZERO_PAYOUT = "zero_payout"

#: Values of ``package_note`` explaining why a package was not allocated.
FULLY_RETURNED = "fully_returned"
REFUND_WITHOUT_IDENTIFIABLE_ITEM = "refund_without_identifiable_item"
NO_BASIS_TO_SPLIT_ON = "no_unit_prices_to_split_on"


def check_totals(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 1: sum H..O, compare against P, and flag the rows that disagree.

    Adds three columns:

    ``components_sum_mxn``
        the sum of H..O, so the comparison can be audited rather than trusted;
    ``total_diff_mxn``
        ``components_sum_mxn - total_mxn`` -- the unexplained amount, signed;
    ``total_matches``
        ``True`` where the row reconciles, ``False`` where it does not, and null where
        there is nothing to compare.
    """
    missing = [c for c in COMPONENT_COLUMNS + [TOTAL_COLUMN] if c not in frame.columns]
    if missing:
        raise ValueError(f"ventas_mx rule 1: columns missing from the cleaned frame: {missing}")

    out = frame.copy()

    # A column that is null for every row arrives as object dtype, which cannot be
    # summed, so coerce before adding rather than assuming the cleaner's dtypes.
    money = out[COMPONENT_COLUMNS].apply(pd.to_numeric, errors="coerce").astype("Float64")
    total = pd.to_numeric(out[TOTAL_COLUMN], errors="coerce").astype("Float64")

    # min_count=1 keeps an all-null row null rather than collapsing it to zero, which
    # would otherwise read as a real 0.00 difference against a null total.
    components = money.sum(axis=1, min_count=1).round(2)
    difference = (components - total).round(2)

    out["components_sum_mxn"] = components.astype("Float64")
    out["total_diff_mxn"] = difference.astype("Float64")
    out["total_matches"] = (
        (difference.abs() <= TOLERANCE).where(difference.notna()).astype("boolean")
    )
    return out


def allocate_packages(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 2: share a package's money across the products inside it.

    Each product's share is its own contribution over the package's:
    ``(units * unit_price) / sum(units * unit_price)``. On these exports every product
    row has ``units = 1`` and the unit prices sum exactly to the package's H, so this is
    the ``W / H`` ratio -- expressed against the summed basis so it stays correct if a
    product line ever carries a quantity above one.

    Every column in :data:`ALLOCATED_COLUMNS` is shared out on that same ratio. Rounding
    residue is given to the largest share, so the products always sum back to the
    package to the cent.

    Rule 3: a returned product must not distort the others, so it is left out of the
    basis. The refund in column O is folded into H first -- the revenue to share is
    ``H + O`` -- and the surviving products divide that between them. Returned products
    keep no money and are flagged with ``is_returned_item``.

    The package row itself is left in place and marked ``package_split``; rule 4
    (:func:`drop_package_rows`) is what removes it.

    Adds ``allocated_from_package``, ``is_returned_item``, ``package_note`` and
    ``package_split``.
    """
    out = frame.copy()
    out["allocated_from_package"] = pd.Series(False, index=out.index, dtype="boolean")
    out["is_returned_item"] = pd.Series(False, index=out.index, dtype="boolean")
    out["package_split"] = pd.Series(False, index=out.index, dtype="boolean")
    out["package_note"] = pd.Series(pd.NA, index=out.index, dtype="string")

    if "row_role" not in out.columns:
        return out

    for _, group in out[out["package_id"].notna()].groupby("package_id", sort=False):
        parents = group[group["row_role"] == PACKAGE_PARENT]
        children = group[group["row_role"] == PACKAGE_CHILD]
        if len(parents) != 1 or children.empty:
            continue

        parent_index = parents.index[0]
        refund = out.at[parent_index, REFUND_COLUMN]
        has_refund = pd.notna(refund) and abs(float(refund)) > TOLERANCE

        returned = _returned_children(out.loc[children.index]) if has_refund else []
        out.loc[returned, "is_returned_item"] = True

        if has_refund and not returned:
            # Money came back but no product row says which one. Guessing would put the
            # refund on the wrong SKU, so leave the package for a human.
            out.loc[group.index, "package_note"] = REFUND_WITHOUT_IDENTIFIABLE_ITEM
            continue

        keep = children.index.difference(pd.Index(returned))
        if keep.empty:
            # Nothing survived to receive the money; keeping the package row is the only
            # way the payout stays in the accounts.
            out.loc[group.index, "package_note"] = FULLY_RETURNED
            continue

        basis = _contribution(out.loc[keep])
        if basis.sum() <= 0:
            out.loc[group.index, "package_note"] = NO_BASIS_TO_SPLIT_ON
            continue

        shares = basis / basis.sum()
        for column, amount in _amounts_to_share(out.loc[parent_index], has_refund).items():
            out.loc[keep, column] = _split(amount, shares)

        out.loc[keep, "allocated_from_package"] = True
        out.at[parent_index, "package_split"] = True

    return out.reset_index(drop=True)


def add_tax(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 5: the tax withheld on the sale, ``H / 1.16 * 0.105``.

    H is quoted gross of the 16% IVA, so dividing by 1.16 gives the amount net of it and
    10.5% of that is the withholding. Reported **negative**, because it is a cost and
    every cost in this report carries the same sign -- which also makes the money columns
    add up to the payout.

    It is computed after the packages are split, so a product that came out of a package
    is taxed on its own share of the revenue rather than on the package's.
    """
    out = frame.copy()
    revenue = pd.to_numeric(out["product_revenue_mxn"], errors="coerce").astype("Float64")
    withheld = (revenue / (1 + IVA_RATE) * WITHHOLDING_RATE).round(2)
    out[TAX_COLUMN] = (-withheld).astype("Float64")
    return out


def add_platform_fee(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 6: the platform's fee, column I with rule 5's tax taken back out.

    Column I is "Cargo por venta e impuestos" -- commission and tax in one figure -- so
    removing the tax leaves the commission.

    On the sign: the arithmetic is done on magnitudes and the result is then negated, so
    the fee is a negative cost like rule 5's tax. ``|I| - |tax|`` reproduces the "Cargo
    por venta" that Reporte_Facturacion bills, to the cent, on 98 of the 104 orders that
    appear in both reports -- Facturacion publishes its charges positive, so that
    comparison is made on magnitudes. Subtracting the signed values instead would give
    -184.33 where ML bills 69.67, which matches nothing.

    The two parts rebuild the source column: ``platform_fee + tax == -|I|``.
    """
    if TAX_COLUMN not in frame.columns:
        raise ValueError(f"ventas_mx rule 6 needs {TAX_COLUMN!r}; run rule 5 first")

    out = frame.copy()
    charged = pd.to_numeric(out[FEE_COLUMN], errors="coerce").astype("Float64").abs()
    tax = out[TAX_COLUMN].astype("Float64").abs()
    out[PLATFORM_FEE_COLUMN] = (-(charged - tax)).round(2).astype("Float64")
    return out


def attach_sales_channel(frame: pd.DataFrame, charges: pd.DataFrame) -> pd.DataFrame:
    """Fill ``sales_channel`` from Reporte_Facturacion.

    "Canal de Venta" (Mercado Libre vs Mercado Shops) is not in the sales export at all
    -- it is only published on the charge rows -- so it has to be joined in by order.
    An order with no charge row keeps a null channel rather than a guessed one.
    """
    out = frame.copy()
    if charges is None or charges.empty or "sales_channel" not in charges.columns:
        if "sales_channel" not in out.columns:
            out["sales_channel"] = pd.Series(pd.NA, index=out.index, dtype="string")
        return out

    channel = (
        charges.dropna(subset=["order_id", "sales_channel"])
        .drop_duplicates(subset=["order_id"])
        .set_index("order_id")["sales_channel"]
    )
    out["sales_channel"] = out["order_id"].map(channel).astype("string")
    return out


def add_period(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 7: a ``YYYYMM`` period from the sale date.

    Strictly this is derived data -- ``sold_at`` is already a timestamp, so any tool that
    can group by month can produce it on demand. It is stored anyway because the people
    reading the workbook are grouping in Excel, where a plain text key is far easier to
    pivot on than a date expression, and because a stored key cannot drift between one
    person's query and another's.
    """
    if PERIOD_SOURCE not in frame.columns:
        raise ValueError(f"ventas_mx rule 7 needs {PERIOD_SOURCE!r} to derive the period")

    out = frame.copy()
    sold = pd.to_datetime(out[PERIOD_SOURCE], errors="coerce")
    out[PERIOD_COLUMN] = sold.dt.strftime("%Y%m").astype("string")
    return out


def flag_revenue_exclusions(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 10: mark the rows that must not count toward revenue.

    A cancelled or fully-returned sale keeps its original ``product_revenue_mxn`` while
    ``total_mxn`` falls to zero, so summing revenue over every row counts sales that
    earned nothing. Excluding the zero-payout rows reproduces the accountant's published
    figure to the cent.

    A *partial* refund is deliberately kept: the sale still paid out, so the revenue was
    real. A null payout is also kept -- only a genuine zero is an exclusion, since a null
    means there was nothing to judge.

    Adds ``counts_toward_revenue`` and, where it is False, ``revenue_exclusion_reason``.
    """
    out = frame.copy()
    payout = pd.to_numeric(out[TOTAL_COLUMN], errors="coerce").astype("Float64")

    excluded = (payout.abs() <= TOLERANCE).fillna(False)
    out[COUNTS_COLUMN] = (~excluded).astype("boolean")
    out[EXCLUSION_REASON_COLUMN] = pd.Series(pd.NA, index=out.index, dtype="string")
    out.loc[excluded, EXCLUSION_REASON_COLUMN] = _exclusion_reason(out.loc[excluded])
    return out


def _exclusion_reason(rows: pd.DataFrame) -> pd.Series:
    """Why a zero-payout row was excluded, read from its status."""
    status = rows["order_status"].fillna("") if "order_status" in rows.columns else ""
    if isinstance(status, str):
        return pd.Series(ZERO_PAYOUT, index=rows.index, dtype="string")

    # "cancel" rather than rule 3's "cancelad": ML writes both "Cancelada por el
    # comprador" and "Cancelaste la venta", and the second is just as much a cancellation.
    reason = pd.Series(ZERO_PAYOUT, index=rows.index, dtype="string")
    reason = reason.mask(status.str.contains(r"devoluci", case=False), RETURNED)
    reason = reason.mask(status.str.contains(r"cancel", case=False), CANCELLED)
    reason = reason.mask(status.str.contains(r"mediaci", case=False), MEDIATION)
    return reason


def excluded_rows(frame: pd.DataFrame) -> pd.DataFrame:
    """The rows rule 10 kept out of revenue, newest sale first."""
    if COUNTS_COLUMN not in frame.columns:
        frame = flag_revenue_exclusions(frame)
    excluded = frame[~frame[COUNTS_COLUMN].fillna(True)]
    if "sold_at" in excluded.columns:
        excluded = excluded.sort_values("sold_at", ascending=False)
    return excluded


def monthly_summary(
    frame: pd.DataFrame, charges: pd.DataFrame | None = None
) -> pd.DataFrame:
    """The monthly picture the period column exists to support.

    Money is summed over the rows rule 10 counts, so ``revenue_mxn`` is sales that
    actually earned. What was left out is shown beside it -- ``excluded_revenue_mxn``
    plus ``revenue_mxn`` equals ``gross_revenue_mxn``, the raw total from the export --
    so nothing disappears without being visible.

    Every cost is negative, which makes the row add up::

        revenue + platform_fee + tax + shipping_revenue + shipping_cost + refunds = payout

    ``charges`` adds the two account-level costs the accountant carries as 广告费 and
    平台仓租. They sit outside that identity on purpose: ML bills them to the account
    rather than to a sale, so they reduce the month's profit without touching any order's
    payout. They are keyed on the month ML billed them, which is the only date they have.
    """
    if PERIOD_COLUMN not in frame.columns:
        frame = add_period(frame)
    if COUNTS_COLUMN not in frame.columns:
        frame = flag_revenue_exclusions(frame)

    counted = frame[frame[COUNTS_COLUMN].fillna(True)]
    aggregations = {
        "orders": ("order_id", "nunique"),
        "units": ("units", "sum"),
        "revenue_mxn": ("product_revenue_mxn", "sum"),
        "platform_fee_mxn": (PLATFORM_FEE_COLUMN, "sum"),
        "tax_withheld_mxn": (TAX_COLUMN, "sum"),
        "shipping_revenue_mxn": ("shipping_revenue_mxn", "sum"),
        "shipping_cost_mxn": ("shipping_cost_mxn", "sum"),
        "refunds_mxn": (REFUND_COLUMN, "sum"),
        "payout_mxn": (TOTAL_COLUMN, "sum"),
        "unexplained_mxn": ("total_diff_mxn", "sum"),
    }
    available = {k: v for k, v in aggregations.items() if v[0] in counted.columns}
    summary = counted.groupby(PERIOD_COLUMN, dropna=False).agg(**available)

    left_out = frame[~frame[COUNTS_COLUMN].fillna(True)]
    dropped = left_out.groupby(PERIOD_COLUMN, dropna=False).agg(
        excluded_orders=("order_id", "nunique"),
        excluded_revenue_mxn=("product_revenue_mxn", "sum"),
    )

    summary = summary.join(dropped, how="outer").reset_index()
    summary["excluded_orders"] = summary["excluded_orders"].fillna(0).astype("Int64")
    summary["excluded_revenue_mxn"] = summary["excluded_revenue_mxn"].astype("Float64").fillna(0.0)
    summary["gross_revenue_mxn"] = (
        summary["revenue_mxn"].astype("Float64").fillna(0.0) + summary["excluded_revenue_mxn"]
    )

    summary = _attach_account_costs(summary, charges)

    for column in summary.columns:
        if column not in (PERIOD_COLUMN, "orders", "units", "excluded_orders"):
            summary[column] = summary[column].astype("Float64").round(2)
    return summary.sort_values(PERIOD_COLUMN).reset_index(drop=True)


def _attach_account_costs(
    summary: pd.DataFrame, charges: pd.DataFrame | None
) -> pd.DataFrame:
    """Add advertising and storage, and the profit left once they are taken off.

    These are the accountant's 广告费 and 平台仓租. They are keyed on the month ML billed
    them, because that is the only date such a charge carries -- no order, no listing, no
    sale date. So on a month where a sale and its billing straddle a boundary these will
    not line up with the sales columns, and that is a property of ML's data rather than
    something to smooth over.
    """
    out = summary.copy()
    costs = account_level_costs(charges)
    for column in ACCOUNT_LEVEL_COSTS:
        out[column] = pd.Series(pd.NA, index=out.index, dtype="Float64")

    if not costs.empty:
        indexed = costs.set_index("charge_month")
        for column in ACCOUNT_LEVEL_COSTS:
            out[column] = out[PERIOD_COLUMN].map(indexed[column]).astype("Float64")

    out["net_after_account_costs_mxn"] = (
        out["payout_mxn"].astype("Float64").fillna(0.0)
        + out["advertising_mxn"].astype("Float64").fillna(0.0)
        + out["storage_mxn"].astype("Float64").fillna(0.0)
    ).round(2)
    return out


def listing_sku_map(frame: pd.DataFrame) -> pd.Series:
    """Listing id -> the SKU that listing mostly sells.

    Facturacion names the listing but never the SKU, so this is what lets a charge reach
    a product. A listing almost always carries one SKU; where colour variants share a
    listing the busiest one wins, which is how the accountant's own product master reads
    it -- their ``BH00335GY`` covers the single ``BH00335PK`` sale on the same listing.
    """
    if not {"listing_id", "sku"} <= set(frame.columns):
        return pd.Series(dtype="object")

    pairs = frame.dropna(subset=["listing_id", "sku"])
    if pairs.empty:
        return pd.Series(dtype="object")

    ranked = (
        pairs.assign(_units=pd.to_numeric(pairs["units"], errors="coerce").fillna(0))
        .groupby(["listing_id", "sku"], as_index=False)["_units"]
        .sum()
        .sort_values(["listing_id", "_units", "sku"], ascending=[True, False, True])
    )
    return ranked.drop_duplicates("listing_id").set_index("listing_id")["sku"]


def sku_summary(frame: pd.DataFrame, charges: pd.DataFrame | None = None) -> pd.DataFrame:
    """One row per period and SKU -- what each product sold and what it settled for.

    The same shape as the accountant's 销售 tab, and it uses the same rule 10 filter, so
    a cancelled sale is left out here exactly as it is on the monthly sheet.

    ``tax_withheld_mxn`` is computed on the **grouped** revenue rather than summed from
    the per-row column. Rule 5 rounds each row to the cent, and adding hundreds of those
    drifts from the figure you get by rounding once -- on this data by up to 0.16 per
    SKU. The accountant rounds once, so this matches them.
    """
    if PERIOD_COLUMN not in frame.columns:
        frame = add_period(frame)
    if COUNTS_COLUMN not in frame.columns:
        frame = flag_revenue_exclusions(frame)

    counted = frame[frame[COUNTS_COLUMN].fillna(True)]
    aggregations = {
        "orders": ("order_id", "nunique"),
        "units": ("units", "sum"),
        "revenue_mxn": ("product_revenue_mxn", "sum"),
        "platform_fee_mxn": (PLATFORM_FEE_COLUMN, "sum"),
        "shipping_revenue_mxn": ("shipping_revenue_mxn", "sum"),
        "shipping_cost_mxn": ("shipping_cost_mxn", "sum"),
        "refunds_mxn": (REFUND_COLUMN, "sum"),
        "payout_mxn": (TOTAL_COLUMN, "sum"),
    }
    available = {k: v for k, v in aggregations.items() if v[0] in counted.columns}
    summary = (
        counted.groupby([PERIOD_COLUMN, "sku"], dropna=False).agg(**available).reset_index()
    )

    revenue = summary["revenue_mxn"].astype("Float64")
    summary["tax_withheld_mxn"] = (
        -(revenue / (1 + IVA_RATE) * WITHHOLDING_RATE)
    ).round(2).astype("Float64")

    units = summary["units"].astype("Float64")
    payout = summary["payout_mxn"].astype("Float64")
    summary["settlement_unit_price_mxn"] = (
        payout / units.where(units != 0)
    ).round(2).astype("Float64")

    for column in summary.columns:
        if column not in (PERIOD_COLUMN, "sku", "orders", "units"):
            summary[column] = summary[column].astype("Float64").round(2)

    summary = _attach_sku_fees(summary, frame, charges)

    order = [
        PERIOD_COLUMN, "sku", "orders", "units", "revenue_mxn", "tax_withheld_mxn",
        "shipping_revenue_mxn", *SKU_FEE_COLUMNS.values(), "fee_total_mxn",
        "settlement_amount_mxn", "platform_fee_mxn", "shipping_cost_mxn", "refunds_mxn",
        "payout_mxn", "settlement_unit_price_mxn", "k3_material",
    ]
    summary = summary[[c for c in order if c in summary.columns]]
    return summary.sort_values([PERIOD_COLUMN, "revenue_mxn"], ascending=[True, False]).reset_index(
        drop=True
    )


def _attach_sku_fees(
    summary: pd.DataFrame, sales: pd.DataFrame, charges: pd.DataFrame | None
) -> pd.DataFrame:
    """Join what ML actually billed per SKU, and settle the row from it.

    ``settlement_amount_mxn`` is built the accountant's way -- revenue, the withheld tax,
    the shipping the buyer paid, and the billed fees -- rather than from ML's own total,
    so it can be compared against ``payout_mxn`` instead of restating it.
    """
    out = summary.copy()
    out["k3_material"] = pd.Series(pd.NA, index=out.index, dtype="string")

    fees = fees_by_sku(charges, listing_sku_map(sales)) if charges is not None else None
    if fees is None or fees.empty:
        return out

    out = out.merge(fees, on=[PERIOD_COLUMN, "sku"], how="left", validate="one_to_one")
    present = [c for c in SKU_FEE_COLUMNS.values() if c in out.columns]
    out["fee_total_mxn"] = out[present].sum(axis=1, min_count=1).round(2).astype("Float64")
    out["settlement_amount_mxn"] = (
        out["revenue_mxn"].astype("Float64").fillna(0.0)
        + out["tax_withheld_mxn"].astype("Float64").fillna(0.0)
        + out["shipping_revenue_mxn"].astype("Float64").fillna(0.0)
        + out["fee_total_mxn"].astype("Float64").fillna(0.0)
    ).round(2)
    return out


#: Rule 9. Columns removed from the accounting output for carrying no accounting signal.
#: Extend this list as more are agreed; each entry should say why it earned its place.
UNNECESSARY_COLUMNS = (
    # E "Paquete de varios productos" -- not the package marker it looks like. It reads
    # Sí on 147 standalone rows in BOCINA and 7,466 in EWTTO, far more than the actual
    # package products, and nothing else in the report explains which rows get it. The
    # package logic uses ``row_role`` instead, derived from the status and row position.
    "is_multi_product_package",
    # F "Pertenece a un kit" -- False on every product row in both stores. No ML kits
    # have been sold, so the column is constant and says nothing.
    "belongs_to_kit",
    # The parsed form of V "Variante", which restates ``listing_variant`` as JSON:
    # "Color : Negro | Voltaje : 127V" against {"Color": "Negro", "Voltaje": "127V"}.
    # The two carry the same content -- no row has one without the other -- so the raw
    # text is kept and the reformatted copy goes.
    "listing_variant_attributes",
    # The year is already the first four characters of rule 7's ``sale_period``.
    "sale_year",
    # Lineage stamped by the pipeline. One export belongs to one seller account, so
    # within a single run these are constant. They are dropped from the review output
    # only: deduplication and Parquet partitioning key on ``seller_id`` and read the
    # cleaned table, which keeps both. Note that order ids are unique per account and
    # not across them, so validated outputs from two stores cannot be concatenated
    # once these are gone.
    "store",
    "seller_id",
)


def drop_unnecessary_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 9: remove the columns that carry no accounting signal.

    Only the accounting output loses them. The cleaned ``ventas_mx`` table keeps every
    source column, so a column dropped here can be reinstated without re-reading the
    workbook -- and a blank E, which marks package parents and exchange rows, stays
    available there as a cross-check on the row-structure detection.
    """
    return frame.drop(columns=[c for c in UNNECESSARY_COLUMNS if c in frame.columns])


def drop_buyer_identity(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 8: remove the "Facturación al comprador" and "Compradores" column blocks.

    The set comes from :data:`~mx_sales.reports.ventas_mx.IDENTITY_COLUMNS`, which is
    derived from the sheet's own group headers, so it follows the cleaner's mapping
    instead of being restated here.

    Note this also removes ``buyer_state`` and ``buyer_municipality``, the only geography
    in the report -- worth keeping (pass this step over) if sales by region is ever
    wanted.
    """
    return frame.drop(columns=[c for c in IDENTITY_COLUMNS if c in frame.columns])


def drop_package_rows(frame: pd.DataFrame) -> pd.DataFrame:
    """Rule 4: remove the package rows whose money has been split onto their products.

    Only rows marked ``package_split`` go. A package that rule 3 left intact -- every
    product returned, a refund nobody can attribute, no unit prices to divide on -- still
    holds the only copy of its money, so removing it would drop that money from the
    accounts.
    """
    if "package_split" not in frame.columns:
        return frame
    return frame[~frame["package_split"].fillna(False)].reset_index(drop=True)


def _returned_children(children: pd.DataFrame) -> list:
    """Which products in a package came back.

    Product rows carry no money, so the refund in the package's column O cannot say
    which item it belongs to; the status wording is the only per-product signal there is.
    """
    status = children["order_status"].fillna("")
    returned = status.str.contains(RETURN_STATUS, case=False, regex=True)
    kept = status.str.contains(KEPT_MONEY_STATUS, case=False, regex=True)
    return list(children.index[returned & ~kept])


def _amounts_to_share(parent: pd.Series, has_refund: bool) -> dict[str, float]:
    """The package amounts to divide, with the refund folded into revenue.

    Column O reverses the returned portion, so adding it to H leaves the revenue that
    actually stands. O is then not shared again -- doing so would deduct it twice --
    which also keeps the H..O total equal to P, so rule 1 still holds after allocation.

    A caveat worth knowing: ML sets O to the reversal of the whole net (revenue less
    fees and shipping), not to the returned product's list price. ``H + O`` is therefore
    revenue-net-of-reversal, not the surviving products' combined price.
    """
    amounts: dict[str, float] = {}
    for column in ALLOCATED_COLUMNS:
        value = parent[column]
        if column == REFUND_COLUMN and has_refund:
            continue
        if pd.isna(value):
            continue
        amounts[column] = float(value)

    if has_refund:
        revenue = float(parent["product_revenue_mxn"] or 0.0)
        amounts["product_revenue_mxn"] = round(revenue + float(parent[REFUND_COLUMN]), 2)
    return amounts


def _contribution(children: pd.DataFrame) -> pd.Series:
    """What each product contributes to the package: ``units * unit_price``."""
    units = pd.to_numeric(children["units"], errors="coerce").fillna(1)
    price = pd.to_numeric(children["unit_price_mxn"], errors="coerce").fillna(0.0)
    return (units * price).astype(float)


def _split(amount: float, shares: pd.Series) -> pd.Series:
    """Divide ``amount`` across ``shares``, rounded to cents and summing back exactly.

    Rounding each share on its own can leave the pieces a cent short of the whole, which
    would quietly create or destroy money. The residue goes to the largest share.
    """
    allocated = (shares * amount).round(2)
    residue = round(amount - allocated.sum(), 2)
    if residue:
        largest = shares.to_numpy().argmax()
        # Round again after adding: float addition reintroduces noise (33.33 + 0.01
        # lands on 33.339999999999996), and these are cent amounts.
        allocated.iloc[largest] = round(allocated.iloc[largest] + residue, 2)
    return allocated


def packages(frame: pd.DataFrame) -> pd.DataFrame:
    """The package rows of a frame, for auditing what rule 2 removed."""
    if "row_role" not in frame.columns:
        return frame.iloc[0:0]
    return frame[frame["row_role"] == PACKAGE_PARENT]


def validate(frame: pd.DataFrame, charges: pd.DataFrame | None = None) -> pd.DataFrame:
    """Apply every rule, in the order they have to run.

    Rules 2 and 3 split the packages, rule 4 removes the rows that have been split, and
    rule 1 checks what is left -- by then every remaining row carries its own money.
    """
    allocated = allocate_packages(frame)      # rules 2 and 3
    remaining = drop_package_rows(allocated)  # rule 4
    checked = flag_revenue_exclusions(check_totals(remaining))  # rules 1 and 10
    derived = add_platform_fee(add_tax(checked))  # rules 5 and 6
    channelled = attach_sales_channel(derived, charges)
    tidied = drop_buyer_identity(add_period(channelled))  # rules 7 and 8
    return drop_unnecessary_columns(tidied)   # rule 9


def control_totals(source: pd.DataFrame, result: pd.DataFrame) -> pd.DataFrame:
    """Tie the rule output back to the sheet it came from.

    The rules move money between rows; they must never create or destroy any. This is
    the check that says so: every money column, totalled before and after, with the
    difference. Anything other than 0.00 is a bug in the rules, not a finding about ML.

    One exception is legitimate. Rule 3 folds a refund into revenue on a package it
    splits, so on such a package H rises and O falls by the same amount; the ``total_mxn``
    line stays exact regardless and is the one to trust.
    """
    rows = [
        {
            "metric": "rows",
            "source": float(len(source)),
            "output": float(len(result)),
            "difference": float(len(result) - len(source)),
        }
    ]
    for column in ALLOCATED_COLUMNS:
        before = _total(source, column)
        after = _total(result, column)
        rows.append(
            {
                "metric": column,
                "source": before,
                "output": after,
                "difference": round(after - before, 2),
            }
        )
    return pd.DataFrame(rows)


def _total(frame: pd.DataFrame, column: str) -> float:
    if column not in frame.columns:
        return 0.0
    return round(float(pd.to_numeric(frame[column], errors="coerce").sum()), 2)


def row_accounting(source: pd.DataFrame, result: pd.DataFrame) -> pd.DataFrame:
    """Explain the row count: what left the sheet, and why."""
    allocated = allocate_packages(source) if "package_split" not in source.columns else source
    return pd.DataFrame(
        [
            {"stage": "rows in the sheet", "rows": len(source)},
            {
                "stage": "package rows split and removed (rule 4)",
                "rows": -int(allocated["package_split"].fillna(False).sum()),
            },
            {
                "stage": "package rows kept (rule 3 could not split)",
                "rows": int(
                    (
                        (result.get("row_role") == PACKAGE_PARENT)
                        if "row_role" in result.columns
                        else pd.Series(dtype=bool)
                    ).sum()
                ),
            },
            {"stage": "rows in the output", "rows": len(result)},
        ]
    )


def exceptions(frame: pd.DataFrame) -> pd.DataFrame:
    """The rows that failed rule 1, largest difference first."""
    failed = frame[frame["total_matches"].eq(False)]
    return failed.reindex(
        failed["total_diff_mxn"].abs().sort_values(ascending=False).index
    )


def summarize(frame: pd.DataFrame) -> pd.DataFrame:
    """How many rows reconciled, how many did not, and by how much."""
    checked = frame[frame["total_matches"].notna()]
    failed = checked[checked["total_matches"].eq(False)]

    return pd.DataFrame(
        [
            {
                "rows": len(frame),
                "checked": len(checked),
                "not_checked": len(frame) - len(checked),
                "matching": int(checked["total_matches"].sum()),
                "differing": len(failed),
                "net_difference_mxn": _scalar(failed["total_diff_mxn"].sum()),
                "largest_difference_mxn": _scalar(failed["total_diff_mxn"].abs().max()),
            }
        ]
    )


def _scalar(value: object) -> float:
    """Aggregates over an empty or all-null selection come back as NA; report 0.00."""
    return 0.0 if pd.isna(value) else round(float(value), 2)
