"""Run the Ventas MX validation rules over one cleaned export."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font

from mx_sales import pipeline
from mx_sales.accounting import ventas_mx as rules
from mx_sales.accounting.labels import chinese
from mx_sales.accounting.detail_sheets import charge_detail, credit_note_detail
from mx_sales.reports.facturacion import cost_breakdown, load_display
from mx_sales.config import SETTINGS, Settings

REPORT_NAME = "ventas_mx_validated"


def prepare(
    path: Path | str, settings: Settings = SETTINGS
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Return the cleaned sheet, the rule output and the charges they were checked against.

    The charges are loaded here rather than by the caller because rule 9's
    ``sales_channel`` needs them: that column exists only on the Facturacion rows.
    """
    items = [
        item
        for item in pipeline.discover(settings)
        if item.spec.name == "ventas_mx" and pipeline.matches_file(item.path, path, settings)
    ]
    if not items:
        raise FileNotFoundError(f"no ventas_mx export matched {path!r}")

    source = pipeline.process_file(items[-1], settings).sort_values("row_index")
    source = source.reset_index(drop=True)
    charges = charges_for(path, settings)
    return source, rules.validate(source, charges), charges


def build_from_file(path: Path | str, settings: Settings = SETTINGS) -> pd.DataFrame:
    """Clean a single sales export and apply the validation rules to it.

    Returns the rows after rule 2 has removed the package rows, so the frame is shorter
    than the sheet by the number of packages it contained.
    """
    return prepare(path, settings)[1]


def write(frame: pd.DataFrame, settings: Settings = SETTINGS, store: str | None = None) -> Path:
    """Persist the validated rows, one file per store.

    Each store gets its own report, so each needs its own output path -- a single fixed
    filename would mean the second store's run silently replaced the first's. The store
    name is not read from ``frame``: rule 9 removes it, since a report covers one store
    and repeating the name on every row says nothing.
    """
    target = settings.processed_dir / REPORT_NAME
    if store:
        target = target / f"store={store}"
    target.mkdir(parents=True, exist_ok=True)

    path = target / f"{REPORT_NAME}.parquet"
    pipeline._prepare_for_parquet(frame).to_parquet(path, engine="pyarrow", index=False)
    return path


def sibling_files(path: Path | str, report: str, settings: Settings = SETTINGS) -> list[Path]:
    """Paths of one report type downloaded into the same snapshot folder as ``path``."""
    items = [
        item
        for item in pipeline.discover(settings)
        if item.spec.name == "ventas_mx" and pipeline.matches_file(item.path, path, settings)
    ]
    if not items:
        return []

    folder = items[-1].path.parent
    return [
        item.path
        for item in pipeline.discover(settings)
        if item.spec.name == report and item.path.parent == folder
    ]


def display_rows(paths: list[Path]) -> pd.DataFrame:
    """Concatenate several charge exports under ML's own Spanish headers."""
    frames = [load_display(p) for p in paths]
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def charges_for(path: Path | str, settings: Settings = SETTINGS) -> pd.DataFrame:
    """The Reporte_Facturacion charges downloaded alongside a sales export.

    Only the files in the same snapshot folder are read: those are the charges that
    belong with that export, rather than whatever happens to exist elsewhere in the tree.
    """
    items = [
        item
        for item in pipeline.discover(settings)
        if item.spec.name == "ventas_mx" and pipeline.matches_file(item.path, path, settings)
    ]
    if not items:
        return pd.DataFrame()

    folder = items[-1].path.parent
    siblings = [
        pipeline.process_file(item, settings)
        for item in pipeline.discover(settings)
        if item.spec.name == "facturacion" and item.path.parent == folder
    ]
    return pd.concat(siblings, ignore_index=True) if siblings else pd.DataFrame()


def store_of(source: pd.DataFrame) -> str | None:
    """The store a cleaned export belongs to, taken before rule 9 drops the column."""
    if "store" not in source.columns or source.empty:
        return None
    stores = source["store"].dropna().unique()
    return str(stores[0]) if len(stores) == 1 else None


def to_excel(
    frame: pd.DataFrame,
    path: Path,
    source: pd.DataFrame | None = None,
    charges: pd.DataFrame | None = None,
    export: Path | str | None = None,
    store: str | None = None,
) -> Path:
    """Write the review workbook.

    ``control`` first, because it is what says the transformation can be trusted; the
    detail only means something once the totals tie out.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    sheets = {
        "resumen": rules.summarize(frame),
        "mensual": rules.monthly_summary(frame, charges),
        "sku": rules.sku_summary(frame, charges),
        "diferencias": _for_excel(rules.exceptions(frame)),
        "excluidos": _for_excel(rules.excluded_rows(frame)),
        "paquetes": _for_excel(_packages(frame)),
        "ventas": _for_excel(frame),
    }
    if charges is not None and not charges.empty:
        sheets["costos"] = cost_breakdown(charges)
    if export is not None:
        detail = charge_detail(
            display_rows(sibling_files(export, "facturacion")),
            rules.listing_sku_map(frame),
            store,
        )
        if not detail.empty:
            sheets["报告明细"] = _for_excel(detail)
        notes = credit_note_detail(display_rows(sibling_files(export, "notas_credito")))
        if not notes.empty:
            sheets["贷记明细"] = _for_excel(notes)
    if source is not None:
        sheets = {"control": rules.control_totals(source, frame), **sheets}

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, sheet in sheets.items():
            sheet.to_excel(writer, sheet_name=name, index=False)
        for name in sheets:
            _add_chinese_header(writer.book[name])
    return path


def _add_chinese_header(worksheet) -> None:
    """Stack the Chinese label under the English one inside the same header cell.

    Keeping both names in one cell leaves the sheet a normal one-header-row table, so
    filters, sorting and pivots all still work on it.
    """
    for cell in worksheet[1]:
        if cell.value is None:
            continue
        english = str(cell.value)
        cell.value = f"{english}\n{chinese(english)}"
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
        cell.font = Font(bold=True)

    worksheet.row_dimensions[1].height = 30
    worksheet.freeze_panes = "A2"


ALLOCATED = rules.ALLOCATED_COLUMNS


def _packages(frame: pd.DataFrame) -> pd.DataFrame:
    """Every row rules 2-4 touched, so the split can be eyeballed package by package."""
    if "package_id" not in frame.columns:
        return frame.iloc[0:0]
    touched = frame[frame["package_id"].notna()]
    columns = [
        c
        for c in [
            "row_index", "package_id", "row_role", "order_id", "sku", "units",
            "unit_price_mxn", *ALLOCATED, "allocated_from_package", "is_returned_item",
            "package_note",
        ]
        if c in touched.columns
    ]
    return touched[columns]


def _for_excel(frame: pd.DataFrame) -> pd.DataFrame:
    """Render values Excel cannot hold: drop tz info, serialize dicts as JSON."""
    out = frame.copy()
    for column in out.columns:
        if pd.api.types.is_datetime64tz_dtype(out[column]):
            out[column] = out[column].dt.tz_localize(None)
        elif out[column].dtype == object:
            out[column] = out[column].map(_as_cell)
    return out


def _as_cell(value: object) -> object:
    if value is None or value is pd.NA:
        return None
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    return str(value)
