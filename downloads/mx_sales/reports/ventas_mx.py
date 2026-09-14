"""``Ventas_MX_Mercado_Libre_y_Mercado_Shops`` -- the order-level sales report.

Worked example of a full report spec. What this export needs cleaned:

* a two-row header whose lower names repeat (``Unidades`` under Ventas, Devoluciones
  and Reclamos), disambiguated by the group row;
* blanks written as a single space, so every column looks 100% populated;
* Spanish long-form dates, some without a year (``"11 de agosto"``), resolved against
  the sale date on the same row;
* Si/No flags, and money columns that need a consistent nullable numeric dtype;
* composite fields: ``"RFC: XAXX010101000"``, ``"S01 Sin efectos fiscales."`` and
  ``"Color : Negro | Voltaje : 127V"``.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from mx_sales.clean.booleans import to_boolean
from mx_sales.clean.dates import to_datetime
from mx_sales.clean.numbers import to_numeric
from mx_sales.clean.text import blank_to_na, split_code_label, strip_prefix
from mx_sales.readers.excel import read_grouped_sheet
from mx_sales.reports.base import ReportSpec, require_columns

SHEET_NAME = "Ventas MX"
GROUP_ROW = 4
HEADER_ROW = 5

#: A package parent row announces how many child rows follow it. In the workbook these
#: are the dark grey rows; the children below them are the light grey ones. Detection is
#: structural rather than colour-based so it survives a restyled export -- the colours
#: are used only to cross-check it (see ``tests/test_ventas_mx.py``).
PACKAGE_PARENT_STATUS = re.compile(r"^Paquete de (\d+) productos?$", re.IGNORECASE)

PACKAGE_PARENT = "package_parent"
PACKAGE_CHILD = "package_child"
STANDALONE = "standalone"

#: Source column -> output column. Anything not listed is dropped, which keeps the
#: output stable when ML adds columns to the export.
TEXT_COLUMNS = {
    "ventas__de_venta": "order_id",
    "ventas__estado": "order_status",
    "ventas__descripcion_del_estado": "order_status_detail",
    "ventas__orden_de_compra": "purchase_order",
    "publicaciones__sku": "sku",
    "publicaciones__de_publicacion": "listing_id",
    # Present only for accounts that run an official store.
    "publicaciones__tienda_oficial": "official_store",
    "publicaciones__titulo_de_la_publicacion": "listing_title",
    "publicaciones__variante": "listing_variant",
    "publicaciones__tipo_de_publicacion": "listing_type",
    "facturacion_al_comprador__factura_adjunta": "invoice_attached_status",
    "facturacion_al_comprador__datos_personales_o_de_empresa": "invoice_name",
    "facturacion_al_comprador__direccion": "invoice_address",
    "facturacion_al_comprador__tipo_de_contribuyente": "taxpayer_type_code",
    "facturacion_al_comprador__tipo_de_usuario": "taxpayer_person_type",
    "facturacion_al_comprador__regimen_fiscal": "tax_regime",
    "compradores__comprador": "buyer_name",
    "compradores__ife": "buyer_tax_id",
    "compradores__domicilio": "buyer_address",
    "compradores__municipio_alcaldia": "buyer_municipality",
    "compradores__estado": "buyer_state",
    "compradores__codigo_postal": "buyer_postal_code",
    "compradores__pais": "buyer_country",
    "envios__forma_de_entrega": "shipping_method",
    "envios__transportista": "shipping_carrier",
    "envios__numero_de_seguimiento": "shipping_tracking_number",
    "envios__url_de_seguimiento": "shipping_tracking_url",
    "devoluciones__forma_de_entrega": "return_shipping_method",
    "devoluciones__transportista": "return_carrier",
    "devoluciones__numero_de_seguimiento": "return_tracking_number",
    "devoluciones__url_de_seguimiento": "return_tracking_url",
    "devoluciones__dinero_a_favor": "return_money_favours",
    "devoluciones__resultado": "return_inspection_result",
    "devoluciones__destino": "return_destination",
    "devoluciones__motivo_del_resultado": "return_result_reason",
}

#: Money columns, coerced to a nullable float.
NUMERIC_COLUMNS = {
    "ventas__ingresos_por_productos_mxn": "product_revenue_mxn",
    "ventas__cargo_por_venta_e_impuestos_mxn": "sale_fee_and_taxes_mxn",
    "ventas__ingresos_por_envio_mxn": "shipping_revenue_mxn",
    "ventas__costos_de_envio_mxn": "shipping_cost_mxn",
    # Only some accounts get this column, and it shifts every money column after it.
    # Mapping is by name rather than position, so its presence does not disturb the
    # others -- but it is a real money column and belongs in the H..O sum.
    "ventas__costo_de_envio_por_cambio_de_producto": "exchange_shipping_cost_mxn",
    "ventas__costo_de_envio_basado_en_medidas_y_peso_declarados": "shipping_cost_declared_mxn",
    "ventas__cargo_por_diferencias_en_medidas_y_peso_del_paquete": "shipping_dimension_adjustment_mxn",
    "ventas__descuentos_y_bonificaciones": "discounts_and_bonuses_mxn",
    "ventas__anulaciones_y_reembolsos_mxn": "cancellations_and_refunds_mxn",
    "ventas__total_mxn": "total_mxn",
    "publicaciones__precio_unitario_de_venta_de_la_publicacion_mxn": "unit_price_mxn",
}

#: Count columns, coerced to a nullable integer.
COUNT_COLUMNS = {
    "ventas__unidades": "units",
    "devoluciones__unidades": "returned_units",
    "reclamos__unidades": "claim_units",
    "reclamos__reclamo_cerrado": "claims_closed",
}

BOOLEAN_COLUMNS = {
    "ventas__paquete_de_varios_productos": "is_multi_product_package",
    "ventas__pertenece_a_un_kit": "belongs_to_kit",
    "publicidad__venta_por_publicidad": "is_advertising_sale",
    "compradores__negocio": "buyer_is_business",
    "devoluciones__revisado_por_mercado_libre": "return_reviewed_by_ml",
    "reclamos__reclamo_abierto": "has_open_claim",
    "reclamos__con_mediacion": "has_mediation",
}

#: Dates the export writes without a year; resolved against ``sold_at``.
RELATIVE_DATE_COLUMNS = {
    "envios__fecha_en_camino": "shipped_at",
    "envios__fecha_entregado": "delivered_at",
    "devoluciones__fecha_en_camino": "return_shipped_at",
    "devoluciones__fecha_entregado": "return_delivered_at",
    "devoluciones__fecha_de_revision": "return_reviewed_at",
}

#: The two buyer-identity groups in the sheet: "Facturación al comprador" (yellow,
#: columns Y..AF) and "Compradores" (green, AG..AN). Named by source group so the set
#: follows the mapping above rather than being restated by hand.
IDENTITY_GROUPS = ("facturacion_al_comprador", "compradores")

#: Output columns fed by those groups, including the ones split out of composite source
#: fields (`Tipo y número de documento`, `CFDI`, `Factura adjunta`).
IDENTITY_COLUMNS = tuple(
    target for source, target in TEXT_COLUMNS.items() if source.startswith(IDENTITY_GROUPS)
) + tuple(
    target for source, target in BOOLEAN_COLUMNS.items() if source.startswith(IDENTITY_GROUPS)
) + ("invoice_attached", "invoice_tax_id", "cfdi_code", "cfdi_description")

#: Final column order of the output table.
OUTPUT_ORDER = [
    "order_id", "sold_at", "order_status", "order_status_detail",
    "is_multi_product_package", "belongs_to_kit", "units",
    "product_revenue_mxn", "sale_fee_and_taxes_mxn", "shipping_revenue_mxn",
    "shipping_cost_mxn", "exchange_shipping_cost_mxn", "shipping_cost_declared_mxn",
    "shipping_dimension_adjustment_mxn",
    "discounts_and_bonuses_mxn", "cancellations_and_refunds_mxn", "total_mxn",
    "net_margin_mxn", "purchase_order", "is_advertising_sale",
    "sku", "listing_id", "official_store", "sales_channel",
    "listing_title", "listing_variant", "listing_variant_attributes",
    "unit_price_mxn", "listing_type",
    "invoice_attached", "invoice_attached_status", "invoice_name", "invoice_tax_id",
    "invoice_address", "taxpayer_type_code", "cfdi_code", "cfdi_description",
    "taxpayer_person_type", "tax_regime",
    "buyer_name", "buyer_is_business", "buyer_tax_id", "buyer_address",
    "buyer_municipality", "buyer_state", "buyer_postal_code", "buyer_country",
    "shipping_method", "shipped_at", "delivered_at", "shipping_carrier",
    "shipping_tracking_number", "shipping_tracking_url", "days_to_deliver",
    "returned_units", "return_shipping_method", "return_shipped_at", "return_delivered_at",
    "return_carrier", "return_tracking_number", "return_tracking_url",
    "return_reviewed_by_ml", "return_reviewed_at", "return_money_favours",
    "return_inspection_result", "return_destination", "return_result_reason",
    "claim_units", "has_open_claim", "claims_closed", "has_mediation",
    "sale_year", "sale_month",
    "row_role", "package_id", "package_size", "row_index", "record_seq",
]


def load(path: Path) -> pd.DataFrame:
    """Read the ``Ventas MX`` sheet using its two-row grouped header."""
    return read_grouped_sheet(path, SHEET_NAME, group_row=GROUP_ROW, header_row=HEADER_ROW)


def transform(frame: pd.DataFrame) -> pd.DataFrame:
    """Clean, type and enrich the raw sales frame."""
    require_columns(frame, ["ventas__de_venta", "ventas__fecha_de_venta"], report="ventas_mx")

    out = pd.DataFrame(index=frame.index)

    for source, target in TEXT_COLUMNS.items():
        out[target] = blank_to_na(frame[source]) if source in frame else pd.NA

    for source, target in NUMERIC_COLUMNS.items():
        out[target] = to_numeric(frame[source]) if source in frame else pd.NA

    for source, target in COUNT_COLUMNS.items():
        out[target] = to_numeric(frame[source], dtype="Int64") if source in frame else pd.NA

    for source, target in BOOLEAN_COLUMNS.items():
        out[target] = to_boolean(frame[source]) if source in frame else pd.NA

    out["sold_at"] = to_datetime(frame["ventas__fecha_de_venta"])
    for source, target in RELATIVE_DATE_COLUMNS.items():
        out[target] = to_datetime(frame[source], reference=out["sold_at"])

    out = _split_composites(frame, out)
    out = _assign_package_roles(out)
    out = _derive(out)

    for column in OUTPUT_ORDER:
        if column not in out:
            out[column] = pd.NA
    return out[OUTPUT_ORDER]


def _split_composites(frame: pd.DataFrame, out: pd.DataFrame) -> pd.DataFrame:
    """Break combined fields into columns that can be filtered and joined on."""
    # "RFC: XAXX010101000" -> "XAXX010101000"
    out["invoice_tax_id"] = strip_prefix(
        frame["facturacion_al_comprador__tipo_y_numero_de_documento"], "RFC:"
    )

    # "S01 Sin efectos fiscales." -> ("S01", "Sin efectos fiscales")
    out["cfdi_code"], out["cfdi_description"] = split_code_label(
        frame["facturacion_al_comprador__cfdi"], r"^([A-Z]\d{2})\s+(.*)$"
    )

    # "Color : Negro | Voltaje : 127V" -> {"Color": "Negro", "Voltaje": "127V"}
    out["listing_variant_attributes"] = blank_to_na(frame["publicaciones__variante"]).map(
        _parse_variant, na_action="ignore"
    )

    # "Factura adjunta" / "Factura no adjunta" -> boolean
    attached = blank_to_na(frame["facturacion_al_comprador__factura_adjunta"]).str.lower()
    out["invoice_attached"] = attached.map(
        lambda v: False if "no adjunta" in v else True if "adjunta" in v else pd.NA,
        na_action="ignore",
    ).astype("boolean")
    return out


def _assign_package_roles(out: pd.DataFrame) -> pd.DataFrame:
    """Label each row as a package parent, a package child, or a standalone sale.

    A parent row carries the package-level money but no SKU or unit price; the ``n``
    rows directly beneath it carry the products and no money at all. Rows are therefore
    matched by position, which is why the reader must preserve the sheet order.
    """
    out = out.reset_index(drop=True)
    out["row_index"] = out.index + 1

    role = pd.Series(STANDALONE, index=out.index, dtype="string")
    package_id = pd.Series(pd.NA, index=out.index, dtype="string")
    package_size = pd.Series(pd.NA, index=out.index, dtype="Int64")

    status = out["order_status"].fillna("")
    for position, value in status.items():
        match = PACKAGE_PARENT_STATUS.match(str(value).strip())
        if match is None:
            continue

        size = int(match.group(1))
        children = [p for p in range(position + 1, min(position + 1 + size, len(out)))]
        if len(children) != size:
            continue

        role.iloc[position] = PACKAGE_PARENT
        package_id.iloc[position] = out.at[position, "order_id"]
        package_size.iloc[position] = size
        for child in children:
            role.iloc[child] = PACKAGE_CHILD
            package_id.iloc[child] = out.at[position, "order_id"]
            package_size.iloc[child] = size

    out["row_role"] = role
    out["package_id"] = package_id
    out["package_size"] = package_size

    # One order can legitimately occupy several rows: a package parent plus its
    # children, or an exchange ("Venta con solicitud de cambio") whose money and
    # product sit on separate rows. Numbering the repeats gives every row a key that
    # is stable across snapshots, so deduplication cannot silently drop one of them.
    out["record_seq"] = (
        out.groupby(["order_id", "sku", "row_role"], dropna=False).cumcount().astype("Int64")
    )
    return out


def _parse_variant(value: str) -> dict[str, str]:
    """Turn ``"Color : Negro | Voltaje : 127V"`` into a dict of attributes."""
    attributes: dict[str, str] = {}
    for part in value.split("|"):
        key, separator, val = part.partition(":")
        if separator:
            attributes[key.strip()] = val.strip()
    return attributes


def _derive(out: pd.DataFrame) -> pd.DataFrame:
    """Add the columns the raw export does not carry but most questions need."""
    out["sale_year"] = out["sold_at"].dt.year.astype("Int64")
    out["sale_month"] = out["sold_at"].dt.strftime("%Y-%m").astype("string")

    # ML's "Total (MXN)" is the payout it computed. Restating it from the components
    # gives an independent figure, so a divergence shows up instead of being assumed away.
    components = [
        "product_revenue_mxn", "sale_fee_and_taxes_mxn", "shipping_revenue_mxn",
        "shipping_cost_mxn", "discounts_and_bonuses_mxn", "cancellations_and_refunds_mxn",
    ]
    # Rounded to cents: summing floats leaves noise like 1.4e-14 where the true value is 0.
    out["net_margin_mxn"] = out[components].sum(axis=1, min_count=1).round(2).astype("Float64")

    delta = out["delivered_at"] - out["sold_at"]
    out["days_to_deliver"] = (delta.dt.total_seconds() / 86400).round(2).astype("Float64")
    return out


SPEC = ReportSpec(
    name="ventas_mx",
    filename_pattern=r"Ventas_MX_Mercado_Libre_y_Mercado_Shops.*\.xlsx$",
    loader=load,
    transformer=transform,
    description="Order-level sales, shipping, returns and claims from the ML/Shops sales report.",
    primary_key=("seller_id", "order_id", "sku", "row_role", "record_seq"),
    partition_by=("seller_id", "sale_month"),
)
