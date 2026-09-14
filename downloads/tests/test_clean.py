"""Tests for the cleaning primitives, using the real shapes found in the exports."""

from __future__ import annotations

import pandas as pd
import pytest

from mx_sales.clean.booleans import to_boolean
from mx_sales.clean.dates import parse_spanish_datetime, to_datetime
from mx_sales.clean.numbers import to_numeric
from mx_sales.clean.text import blank_to_na, slugify, split_code_label, strip_prefix


class TestBlankToNa:
    def test_single_space_becomes_na(self):
        # ML writes " " rather than an empty cell, which hides the real null rate.
        result = blank_to_na(pd.Series([" ", "Entregado", "   ", ""]))
        assert result.isna().tolist() == [True, False, True, True]

    def test_collapses_internal_whitespace(self):
        assert blank_to_na(pd.Series(["  Mercado   Envios  "]))[0] == "Mercado Envios"

    def test_non_breaking_space_is_blank(self):
        assert blank_to_na(pd.Series([" "])).isna().all()


class TestSlugify:
    @pytest.mark.parametrize(
        "value,expected",
        [
            ("Ingresos por envío (MXN)", "ingresos_por_envio_mxn"),
            ("# de venta", "de_venta"),
            ("Municipio/Alcaldía", "municipio_alcaldia"),
            ("Régimen Fiscal", "regimen_fiscal"),
        ],
    )
    def test_slugify(self, value, expected):
        assert slugify(value) == expected


class TestNumbers:
    def test_parses_csv_thousands_format(self):
        # CSV exports render money as "31,350.43"; XLSX exports arrive already typed.
        result = to_numeric(pd.Series(["31,350.43", "-105,502.20", " "]))
        assert result[0] == pytest.approx(31350.43)
        assert result[1] == pytest.approx(-105502.20)
        assert pd.isna(result[2])

    def test_passes_through_native_numerics(self):
        result = to_numeric(pd.Series([633.36, -127, None]))
        assert result[0] == pytest.approx(633.36)
        assert result.dtype == "Float64"

    def test_integer_dtype(self):
        assert to_numeric(pd.Series(["1", "2"]), dtype="Int64").dtype == "Int64"


class TestBooleans:
    def test_si_no_mapping(self):
        result = to_boolean(pd.Series(["Sí", "Si", "No", " ", "No aplica"]))
        assert result.tolist()[:3] == [True, True, False]
        assert pd.isna(result[3])
        assert result[4] is False or result[4] == False  # noqa: E712

    def test_dtype_is_nullable_boolean(self):
        assert to_boolean(pd.Series(["Si", " "])).dtype == "boolean"


class TestSpanishDates:
    def test_full_date_with_time(self):
        assert parse_spanish_datetime("3 de agosto de 2026 09:24 hs.") == pd.Timestamp("2026-08-03 09:24")

    def test_pipe_separated_time_without_year(self):
        reference = pd.Timestamp("2026-08-03 09:24")
        assert parse_spanish_datetime("4 de agosto | 11:36", reference) == pd.Timestamp("2026-08-04 11:36")

    def test_date_only_without_year(self):
        reference = pd.Timestamp("2026-07-27 16:14")
        assert parse_spanish_datetime("27 de agosto", reference) == pd.Timestamp("2026-08-27")

    def test_year_rolls_forward_across_december(self):
        # A sale on 30 December delivered "2 de enero" belongs to the next year.
        reference = pd.Timestamp("2026-12-30 10:00")
        assert parse_spanish_datetime("2 de enero | 08:00", reference) == pd.Timestamp("2027-01-02 08:00")

    def test_year_less_date_without_reference_is_nat(self):
        assert pd.isna(parse_spanish_datetime("2 de enero"))

    def test_blank_and_garbage_are_nat(self):
        assert pd.isna(parse_spanish_datetime(" "))
        assert pd.isna(parse_spanish_datetime("no aplica"))

    def test_vectorized_uses_row_reference(self):
        values = pd.Series(["4 de agosto | 11:36", "2 de enero | 08:00"])
        reference = pd.Series([pd.Timestamp("2026-08-03"), pd.Timestamp("2026-12-30")])
        result = to_datetime(values, reference=reference)
        assert result.tolist() == [pd.Timestamp("2026-08-04 11:36"), pd.Timestamp("2027-01-02 08:00")]


class TestComposites:
    def test_strip_rfc_prefix(self):
        assert strip_prefix(pd.Series(["RFC: XAXX010101000"]), "RFC:")[0] == "XAXX010101000"

    def test_split_cfdi_code_and_label(self):
        # The same code appears with and without a trailing period in one export.
        code, label = split_code_label(
            pd.Series(["S01 Sin efectos fiscales", "S01 Sin efectos fiscales.", " "]),
            r"^([A-Z]\d{2})\s+(.*)$",
        )
        assert code.tolist()[:2] == ["S01", "S01"]
        assert label.tolist()[:2] == ["Sin efectos fiscales", "Sin efectos fiscales"]
        assert pd.isna(code[2])
