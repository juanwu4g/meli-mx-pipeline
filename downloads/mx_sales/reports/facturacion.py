"""``Reporte_Facturacion_MercadoLibre_<Mes><Año>`` -- the itemized charge report.

This is the file that explains what ML actually billed. Each row is one charge tied
(usually) to one sale, so pivoting ``Detalle`` by sale number turns the charges into
per-order fee columns -- which is what rule 9 needs.

Charge amounts are published as positive numbers here even though they are costs; the
sign convention is left as ML publishes it and documented on each column.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from mx_sales.clean.numbers import to_numeric
from mx_sales.clean.text import blank_to_na
from mx_sales.readers.excel import read_sheet
from mx_sales.reports.base import ReportSpec, require_columns

SHEET_NAME = "REPORT"
HEADER_ROW = 7

#: ``Detalle`` value -> canonical fee key. Matched case-insensitively on a prefix, so
#: ML's wording drift ("Cargo por envíos de Mercado Libre" vs "... de Mercado Envíos")
#: does not silently drop a charge.
FEE_CATEGORIES: dict[str, str] = {
    "cargo por venta": "commission",
    "anulacion del cargo por venta": "commission_reversal",
    "cargo por envios": "shipping",
    "anulacion del cargo por envios": "shipping_reversal",
    "cargo por campana de publicidad": "advertising",
    "cargo por servicio de almacenamiento": "storage",
    "cargo por servicio de colecta": "pickup",
    "cargo por retiro de stock": "stock_withdrawal",
    "cargo por devolucion": "return_fee",
    "cargo por mantenimiento": "maintenance",
    "anulacion del cargo por mantenimiento": "maintenance_reversal",
    "dinero en garantia del programa beneficio de reputacion": "reputation_guarantee",
}

#: Fee categories that ML ties to an individual sale. Everything else is an account-level
#: cost -- storage, advertising, stock withdrawal -- and belongs to no particular order.
PER_SALE_CATEGORIES = frozenset(
    {"commission", "commission_reversal", "shipping", "shipping_reversal", "return_fee"}
)

COLUMNS = {
    "n_de_factura_fiscal": "invoice_number",
    "fecha_del_cargo": "charged_at",
    "numero_del_cargo": "charge_id",
    "detalle": "charge_detail",
    "estado_del_cargo": "charge_status",
    "cargo_que_bonifica": "reverses_charge_id",
    "valor_del_cargo": "charge_amount_mxn",
    "porcentaje_por_categoria": "category_rate",
    "numero_de_venta": "order_id",
    "pago": "payment_id",
    "fecha_de_venta": "sold_at",
    "canal_de_venta": "sales_channel",
    "cantidad_vendida": "quantity_sold",
    "precio_unitario": "unit_price_mxn",
    "total_de_la_venta": "order_total_mxn",
    "numero_de_envio": "shipment_id",
    "numero_de_publicacion": "listing_id",
    "titulo_de_publicacion": "listing_title",
    "tipo_de_publicacion": "listing_type",
    "codigo_ml": "ml_code",
}

ID_COLUMNS = ("order_id", "charge_id", "payment_id", "shipment_id", "reverses_charge_id")
NUMERIC_COLUMNS = (
    "charge_amount_mxn", "category_rate", "unit_price_mxn", "order_total_mxn",
)


def load(path: Path) -> pd.DataFrame:
    return read_sheet(path, SHEET_NAME, header_row=HEADER_ROW)


def load_display(path: Path) -> pd.DataFrame:
    """The charge rows under ML's own Spanish column names.

    :func:`load` slugifies the header so the cleaner can address columns by a stable key.
    Sheets that mirror the export back to the reader need the wording ML actually printed,
    so this keeps it. Padding columns Excel leaves behind are dropped.
    """
    frame = pd.read_excel(path, sheet_name=SHEET_NAME, header=HEADER_ROW)
    keep = [c for c in frame.columns if not str(c).startswith("Unnamed")]
    return frame[keep].dropna(how="all").reset_index(drop=True)


def transform(frame: pd.DataFrame) -> pd.DataFrame:
    """Clean the charge rows and tag each with a canonical fee category."""
    require_columns(frame, ["detalle", "valor_del_cargo"], report="facturacion")

    out = pd.DataFrame(index=frame.index)
    for source, target in COLUMNS.items():
        out[target] = blank_to_na(frame[source]) if source in frame.columns else pd.NA

    for column in NUMERIC_COLUMNS:
        out[column] = to_numeric(out[column])

    # Excel hands back 16-digit ids as float64 ("2000017697366516.0"); going through
    # Int64 keeps them exact and makes them joinable against the sales report.
    for column in ID_COLUMNS:
        out[column] = _as_id(frame_column(frame, COLUMNS, column))

    out["quantity_sold"] = to_numeric(out["quantity_sold"], dtype="Int64")
    out["charged_at"] = pd.to_datetime(out["charged_at"], errors="coerce", format="mixed")
    out["sold_at"] = pd.to_datetime(out["sold_at"], errors="coerce", format="mixed")
    out["fee_category"] = out["charge_detail"].map(categorize_fee)
    out["charge_month"] = out["charged_at"].dt.strftime("%Y%m").astype("string")
    return out


def frame_column(frame: pd.DataFrame, mapping: dict[str, str], target: str) -> pd.Series:
    """Fetch the raw column feeding ``target``, or an all-null series if it is absent."""
    for source, name in mapping.items():
        if name == target:
            if source in frame.columns:
                return frame[source]
            break
    return pd.Series(pd.NA, index=frame.index, dtype="object")


def _as_id(series: pd.Series) -> pd.Series:
    """Render a numeric-looking identifier as exact digits, never scientific notation."""
    numeric = pd.to_numeric(series, errors="coerce")
    as_text = numeric.astype("Int64").astype("string")
    # Fall back to the original text where the value was never numeric.
    return as_text.fillna(blank_to_na(series))


def categorize_fee(detail: object) -> object:
    """Map a Spanish ``Detalle`` string to a canonical fee key."""
    if detail is None or detail is pd.NA or not isinstance(detail, str):
        return pd.NA
    normalized = _normalize(detail)
    # Longest key first so "anulacion del cargo por venta" wins over "cargo por venta".
    for key in sorted(FEE_CATEGORIES, key=len, reverse=True):
        if key in normalized:
            return FEE_CATEGORIES[key]
    return "other"


def _normalize(text: str) -> str:
    import unicodedata

    stripped = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", stripped.lower()).strip()


def cost_breakdown(charges: pd.DataFrame) -> pd.DataFrame:
    """Classify what ML charged, one row per ``Detalle``.

    ``Detalle`` is ML's own wording and the only description a charge carries, so it is
    kept verbatim alongside the canonical ``fee_category``. ``tied_to_sale`` counts the
    charges that name an order: commission and shipping always do, while storage,
    advertising and stock withdrawal never do -- those are account-level costs that
    belong to no particular sale, which is worth seeing rather than assuming.

    Reversals stay on their own row rather than being netted off, so both the charge and
    its cancellation are visible.
    """
    if charges.empty:
        return pd.DataFrame(
            columns=["charge_month", "charge_detail", "fee_category", "charges",
                     "tied_to_sale", "total_mxn"]
        )

    working = charges.copy()
    working["tied"] = working["order_id"].notna()

    breakdown = (
        working.groupby(["charge_month", "charge_detail", "fee_category"], dropna=False)
        .agg(
            charges=("charge_amount_mxn", "size"),
            tied_to_sale=("tied", "sum"),
            total_mxn=("charge_amount_mxn", "sum"),
        )
        .reset_index()
    )
    breakdown["total_mxn"] = breakdown["total_mxn"].astype("Float64").round(2)
    breakdown["tied_to_sale"] = breakdown["tied_to_sale"].astype("Int64")

    breakdown = breakdown.sort_values(
        ["charge_month", "total_mxn"], ascending=[True, False]
    ).reset_index(drop=True)

    totals = pd.DataFrame(
        [
            {
                "charge_month": "TOTAL",
                "charge_detail": "",
                "fee_category": "",
                "charges": int(breakdown["charges"].sum()),
                "tied_to_sale": int(breakdown["tied_to_sale"].sum()),
                "total_mxn": round(float(breakdown["total_mxn"].sum()), 2),
            }
        ]
    )
    return pd.concat([breakdown, totals], ignore_index=True)


#: Charges voided by a credit note are refunded to the seller, so counting them would
#: overstate the fees actually borne. Charges voided *on the invoice* are left in: ML
#: has already netted those, and the accountant does not deduct them either.
CREDIT_NOTE_VOID = "Anulado en nota de crédito"

#: Fee categories that reach an individual product, and the column each becomes.
SKU_FEE_COLUMNS = {
    "commission": "fee_commission_mxn",
    "commission_reversal": "fee_commission_reversal_mxn",
    "shipping": "fee_shipping_mxn",
    "shipping_reversal": "fee_shipping_reversal_mxn",
    "return_fee": "fee_return_mxn",
}


def fees_by_sku(charges: pd.DataFrame, listing_to_sku: pd.Series) -> pd.DataFrame:
    """Charges per sale period and SKU, one column per fee category.

    Facturacion names no SKU. It does name the listing (``Número de publicación``), and
    the sales export maps every listing to a SKU, so the two join on that -- no
    hand-maintained product master is needed.

    Charges with no listing are account-level -- storage, advertising, pickup, page
    maintenance -- and belong to no product, so they are left out rather than spread
    across SKUs on a guess. Charges voided by a credit note are dropped, and amounts are
    negated: ML publishes its charges positive, while every cost in our output is negative.
    """
    empty = pd.DataFrame(columns=["sale_period", "sku", *SKU_FEE_COLUMNS.values()])
    if charges is None or charges.empty or "listing_id" not in charges.columns:
        return empty

    working = charges.dropna(subset=["listing_id", "sold_at"]).copy()
    if "charge_status" in working.columns:
        # A normal charge has no status at all. ``ne`` on a nullable column yields NA
        # there, and boolean indexing drops NA rows -- which would silently discard
        # every ordinary charge and keep only the voided ones.
        voided = working["charge_status"].eq(CREDIT_NOTE_VOID).fillna(False)
        working = working[~voided]

    working["sku"] = working["listing_id"].map(listing_to_sku)
    working = working.dropna(subset=["sku"])
    working = working[working["fee_category"].isin(SKU_FEE_COLUMNS)]
    if working.empty:
        return empty

    working["sale_period"] = working["sold_at"].dt.strftime("%Y%m")
    pivot = (
        working.pivot_table(
            index=["sale_period", "sku"],
            columns="fee_category",
            values="charge_amount_mxn",
            aggfunc="sum",
        )
        .rename(columns=SKU_FEE_COLUMNS)
        .reset_index()
    )
    pivot.columns.name = None

    for column in SKU_FEE_COLUMNS.values():
        if column not in pivot.columns:
            pivot[column] = pd.NA
        pivot[column] = (-pivot[column].astype("Float64")).round(2)
    return pivot


#: Costs that belong to the account rather than to any sale, and the column each becomes.
#: ML never names an order or a listing on these, so they can only be reported per month.
#: ``pickup`` (colecta) sits with storage, which is how the accountant groups it.
ACCOUNT_LEVEL_COSTS = {
    "advertising_mxn": ("advertising",),
    "storage_mxn": ("storage", "pickup"),
}


def account_level_costs(charges: pd.DataFrame) -> pd.DataFrame:
    """Advertising and Full storage per charge month.

    These are the two costs the accountant carries on their settlement sheet as 广告费
    and 平台仓租. They name no order and no listing, so unlike commission and shipping
    they cannot be pushed down to a sale -- the month ML billed them is as far as they go.

    Credit-note voids are dropped, the same rule the per-SKU fees use: on this data that
    is what turns advertising of 6,394.60 into the 6,277.34 they publish. Amounts come
    back negative, matching the sign every other cost in the output carries.
    """
    columns = ["charge_month", *ACCOUNT_LEVEL_COSTS]
    if charges is None or charges.empty or "fee_category" not in charges.columns:
        return pd.DataFrame(columns=columns)

    working = charges[charges["listing_id"].isna()] if "listing_id" in charges else charges
    if "charge_status" in working.columns:
        voided = working["charge_status"].eq(CREDIT_NOTE_VOID).fillna(False)
        working = working[~voided]
    if working.empty:
        return pd.DataFrame(columns=columns)

    out = pd.DataFrame({"charge_month": sorted(working["charge_month"].dropna().unique())})
    for column, categories in ACCOUNT_LEVEL_COSTS.items():
        totals = (
            working[working["fee_category"].isin(categories)]
            .groupby("charge_month")["charge_amount_mxn"]
            .sum()
        )
        out[column] = (-out["charge_month"].map(totals)).round(2).astype("Float64")
    return out


def fees_by_order(charges: pd.DataFrame) -> pd.DataFrame:
    """Pivot the charge rows into one row per order, one column per fee category.

    Reversals are folded into the charge they cancel, so ``fee_commission_mxn`` is the
    net commission actually borne by the order. Keyed by ``(store, order_id)`` because
    order ids are only unique within a seller account.
    """
    tagged = charges.dropna(subset=["order_id"]).copy()
    if tagged.empty:
        return pd.DataFrame(columns=["store", "order_id"])
    if "store" not in tagged.columns:
        tagged["store"] = ""

    signed = tagged["charge_amount_mxn"].astype("Float64")
    tagged["net_amount"] = signed
    tagged["base_category"] = tagged["fee_category"].str.replace("_reversal", "", regex=False)

    pivot = (
        tagged.pivot_table(
            index=["store", "order_id"],
            columns="base_category",
            values="net_amount",
            aggfunc="sum",
        )
        .rename(columns=lambda c: f"fee_{c}_mxn")
        .reset_index()
    )
    pivot.columns.name = None

    fee_columns = [c for c in pivot.columns if c.startswith("fee_")]
    pivot["fee_total_mxn"] = pivot[fee_columns].sum(axis=1, min_count=1).round(2)
    for column in fee_columns + ["fee_total_mxn"]:
        pivot[column] = pivot[column].round(2).astype("Float64")
    return pivot


SPEC = ReportSpec(
    name="facturacion",
    filename_pattern=r"Reporte_Facturacion_MercadoLibre_.*\.xlsx$",
    loader=load,
    transformer=transform,
    description="Itemized ML charges (commission, shipping, ads, Full storage) per sale.",
    primary_key=("store", "charge_id"),
    partition_by=("store", "charge_month"),
)
