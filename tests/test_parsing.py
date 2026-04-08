"""Tests voor verlofwaarde.parsing — header-matching, coercion, validatie."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from openpyxl import Workbook

from verlofwaarde.parsing import read_aanvulling, read_verlofsaldi

FIXTURES = Path(__file__).parent / "fixtures"


def test_read_verlofsaldi_joris_fixture():
    df = read_verlofsaldi(str(FIXTURES / "verlofsaldi_joris.xlsx"))
    assert list(df.columns) == ["mdw", "naam", "dv", "type_verlof", "bkjr", "beginsaldo"]
    assert len(df) == 5
    assert df["mdw"].iloc[0] == "10004"
    assert df["beginsaldo"].sum() == pytest.approx(0.0 + 5.68 + 160.0 + 28.0 + 56.0)
    # Type verlof komt er onveranderd in (we strippen jaartallen pas in calculator).
    assert "Bovenwettelijke rechten 2025" in df["type_verlof"].tolist()


def test_read_aanvulling_joris_fixture():
    df = read_aanvulling(str(FIXTURES / "aanvulling_joris.xlsx"))
    assert {"mdw", "bkjr", "aanvulling_pu"}.issubset(df.columns)
    assert len(df) == 5
    # Lookup 2026 → €6,57
    row_2026 = df[(df["mdw"] == "10004") & (df["bkjr"] == 2026)].iloc[0]
    assert row_2026["aanvulling_pu"] == pytest.approx(6.57)


def test_read_verlofsaldi_missing_required_column_raises():
    wb = Workbook()
    ws = wb.active
    ws.append(["Mdw.", "Naam", "Type verlof"])  # geen Beginsaldo
    ws.append(["10004", "Test", "Wettelijke rechten 2025"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    with pytest.raises(ValueError, match="beginsaldo"):
        read_verlofsaldi(buf)


def test_read_aanvulling_accepts_alternate_header_spellings():
    wb = Workbook()
    ws = wb.active
    ws.append(["Personeelsnr", "Boekjaar", "Totaalbedrag"])
    ws.append(["10004", 2025, 4.54])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    df = read_aanvulling(buf)
    assert df["mdw"].iloc[0] == "10004"
    assert df["bkjr"].iloc[0] == 2025
    assert df["aanvulling_pu"].iloc[0] == pytest.approx(4.54)
