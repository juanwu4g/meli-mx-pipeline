"""Per-report specs. One module per ML report type."""

from mx_sales.reports.base import ReportSpec, add_lineage, require_columns

__all__ = ["ReportSpec", "add_lineage", "require_columns"]
