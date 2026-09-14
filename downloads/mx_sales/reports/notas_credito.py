"""``Reporte_Notas_Credito_MercadoLibre_<Mes><Año>`` -- the credit-note report.

ML publishes this on exactly the same 33-column schema as ``Reporte_Facturacion`` -- a
credit note is a Facturacion row, just split into its own file -- so the loader and the
fee categorisation are reused wholesale rather than restated.

Every row is an ``Anulación…``: these are the same events the Facturacion rows carry as
``charge_status = "Anulado en nota de crédito"``, seen from the other side.
"""

from __future__ import annotations

from mx_sales.reports.facturacion import load, load_display, transform
from mx_sales.reports.base import ReportSpec

SPEC = ReportSpec(
    name="notas_credito",
    filename_pattern=r"Reporte_Notas_Credito_MercadoLibre_.*\.xlsx$",
    loader=load,
    transformer=transform,
    description="ML credit notes -- charge reversals, same schema as Reporte_Facturacion.",
    primary_key=("store", "charge_id"),
    partition_by=("store", "charge_month"),
)

__all__ = ["SPEC", "load", "load_display", "transform"]
