"""The catalogue of known report types.

Registering a spec here is the only wiring a new report needs; discovery, cleaning,
Parquet output and the DuckDB view all follow from it.
"""

from __future__ import annotations

from pathlib import Path

from mx_sales.reports import facturacion, notas_credito, ventas_mx
from mx_sales.reports.base import ReportSpec

SPECS: tuple[ReportSpec, ...] = (
    ventas_mx.SPEC,
    facturacion.SPEC,
    notas_credito.SPEC,
    # Remaining report types found in the downloads, not yet specified:
    #   settlement_v2 / reserve-release / account_statement / collection / withdraw  (CSV, ";")
    #   Cargos_Full / Pagos_Facturas                                                 (XLSX)
    #   Costos_por_servicio_almacenamiento / stock_general_full / conciliation / Returns
)

SPECS_BY_NAME = {spec.name: spec for spec in SPECS}


def spec_for(path: Path) -> ReportSpec | None:
    """Return the spec that claims ``path``, or ``None`` if the file is unrecognized."""
    for spec in SPECS:
        if spec.matches(path):
            return spec
    return None


def get(name: str) -> ReportSpec:
    """Look a spec up by table name."""
    try:
        return SPECS_BY_NAME[name]
    except KeyError:
        known = ", ".join(sorted(SPECS_BY_NAME)) or "(none registered)"
        raise KeyError(f"unknown report {name!r}; registered: {known}") from None
