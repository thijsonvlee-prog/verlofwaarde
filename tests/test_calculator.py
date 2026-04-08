"""Tests voor verlofwaarde.calculator — extract-helpers en end-to-end-berekening."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from verlofwaarde.calculator import (
    bereken,
    extract_basis_type,
    extract_opbouwjaar,
    unieke_basis_types,
)
from verlofwaarde.instellingen import apply_defaults
from verlofwaarde.parsing import read_aanvulling, read_verlofsaldi

FIXTURES = Path(__file__).parent / "fixtures"


@pytest.mark.parametrize(
    "naam, jaar",
    [
        ("Bovenwettelijke rechten 2025", 2025),
        ("Bovenwettelijke rechten 2025 (aanv.)", 2025),
        ("Wettelijke rechten 2026", 2026),
        ("Negatief verlof", None),
        ("", None),
    ],
)
def test_extract_opbouwjaar(naam, jaar):
    assert extract_opbouwjaar(naam) == jaar


@pytest.mark.parametrize(
    "naam, basis",
    [
        ("Bovenwettelijke rechten 2025", "Bovenwettelijke rechten"),
        ("Bovenwettelijke rechten 2025 (aanv.)", "Bovenwettelijke rechten (aanv.)"),
        ("Wettelijke rechten 2026", "Wettelijke rechten"),
        ("Negatief verlof", "Negatief verlof"),
    ],
)
def test_extract_basis_type(naam, basis):
    assert extract_basis_type(naam) == basis


def test_joris_voorbeeld_end_to_end():
    """Verifieer de handberekening uit het plan: gewogen gemiddelde ≈ €6,52."""
    verlof = read_verlofsaldi(str(FIXTURES / "verlofsaldi_joris.xlsx"))
    aanvulling = read_aanvulling(str(FIXTURES / "aanvulling_joris.xlsx"))

    basis_types = unieke_basis_types(verlof)
    instellingen = apply_defaults(basis_types)

    result = bereken(verlof, aanvulling, instellingen)

    assert len(result.resultaat) == 1
    row = result.resultaat.iloc[0]
    assert row["Mdw."] == "10004"
    assert row["Naam"] == "Joris Alofs"
    assert row["Totaal uren"] == pytest.approx(249.68)
    assert row["Totaal aanvulling (€)"] == pytest.approx(1628.8676, abs=1e-3)
    assert row["Gewogen aanvulling p/u (€)"] == pytest.approx(6.5238, abs=1e-3)


def test_business_voorbeeld_3eu_4eu():
    """Synthetische TEST-medewerker uit de specificatie: gewogen p/u ≈ €3,87."""
    verlof = pd.DataFrame(
        {
            "mdw": ["TEST"] * 3,
            "naam": ["TEST"] * 3,
            "dv": ["1"] * 3,
            "type_verlof": [
                "Bovenwettelijke rechten 2025",
                "Wettelijke rechten 2026",
                "Bovenwettelijke rechten 2026",
            ],
            "bkjr": pd.Series([2026, 2026, 2026], dtype="Int64"),
            "beginsaldo": [32.0, 160.0, 60.0],
        }
    )
    aanvulling = pd.DataFrame(
        {
            "mdw": ["TEST", "TEST"],
            "naam": ["TEST", "TEST"],
            "dv": ["1", "1"],
            "bkjr": pd.Series([2025, 2026], dtype="Int64"),
            "aanvulling_pu": [3.0, 4.0],
        }
    )
    instellingen = {"Bovenwettelijke rechten": True, "Wettelijke rechten": True}

    result = bereken(verlof, aanvulling, instellingen)

    row = result.resultaat.iloc[0]
    assert row["Totaal uren"] == pytest.approx(252.0)
    assert row["Totaal aanvulling (€)"] == pytest.approx(976.0)
    assert row["Gewogen aanvulling p/u (€)"] == pytest.approx(3.873, abs=1e-3)


def test_niet_aanvullingsrechtig_drukt_gemiddelde():
    """Niet-aanvullingsrechtige uren tellen wel in de noemer maar met €0 in de teller."""
    verlof = pd.DataFrame(
        {
            "mdw": ["1", "1"],
            "naam": ["A", "A"],
            "dv": ["1", "1"],
            "type_verlof": [
                "Wettelijke rechten 2026",
                "Bovenwettelijke rechten 2026",
            ],
            "bkjr": pd.Series([2026, 2026], dtype="Int64"),
            "beginsaldo": [100.0, 100.0],
        }
    )
    aanvulling = pd.DataFrame(
        {
            "mdw": ["1"],
            "naam": ["A"],
            "dv": ["1"],
            "bkjr": pd.Series([2026], dtype="Int64"),
            "aanvulling_pu": [4.0],
        }
    )
    # Bovenwettelijk uitgezet → telt mee in noemer (100u) met €0 in teller.
    instellingen = {"Wettelijke rechten": True, "Bovenwettelijke rechten": False}

    result = bereken(verlof, aanvulling, instellingen)
    row = result.resultaat.iloc[0]
    assert row["Totaal uren"] == pytest.approx(200.0)
    assert row["Totaal aanvulling (€)"] == pytest.approx(400.0)
    assert row["Gewogen aanvulling p/u (€)"] == pytest.approx(2.0)


def test_geen_jaartal_wordt_volledig_genegeerd():
    verlof = pd.DataFrame(
        {
            "mdw": ["1", "1"],
            "naam": ["A", "A"],
            "dv": ["1", "1"],
            "type_verlof": ["Negatief verlof", "Wettelijke rechten 2026"],
            "bkjr": pd.Series([2026, 2026], dtype="Int64"),
            "beginsaldo": [40.0, 100.0],
        }
    )
    aanvulling = pd.DataFrame(
        {
            "mdw": ["1"],
            "naam": ["A"],
            "dv": ["1"],
            "bkjr": pd.Series([2026], dtype="Int64"),
            "aanvulling_pu": [5.0],
        }
    )
    result = bereken(verlof, aanvulling, {"Wettelijke rechten": True, "Negatief verlof": True})
    row = result.resultaat.iloc[0]
    # Negatief-verlof-rij is volledig genegeerd, niet in noemer.
    assert row["Totaal uren"] == pytest.approx(100.0)
    assert row["Gewogen aanvulling p/u (€)"] == pytest.approx(5.0)


def test_ontbrekende_aanvulling_wordt_gemeld():
    verlof = pd.DataFrame(
        {
            "mdw": ["1"],
            "naam": ["A"],
            "dv": ["1"],
            "type_verlof": ["Wettelijke rechten 2024"],
            "bkjr": pd.Series([2026], dtype="Int64"),
            "beginsaldo": [40.0],
        }
    )
    aanvulling = pd.DataFrame(
        {
            "mdw": ["1"],
            "naam": ["A"],
            "dv": ["1"],
            "bkjr": pd.Series([2026], dtype="Int64"),
            "aanvulling_pu": [5.0],
        }
    )
    result = bereken(verlof, aanvulling, {"Wettelijke rechten": True})
    assert len(result.ontbrekende_aanvullingen) == 1
    # Geen match → effectieve aanvulling 0, gewogen p/u dus 0.
    assert result.resultaat.iloc[0]["Gewogen aanvulling p/u (€)"] == pytest.approx(0.0)


def test_onbekend_basis_type_wordt_gemeld():
    verlof = pd.DataFrame(
        {
            "mdw": ["1"],
            "naam": ["A"],
            "dv": ["1"],
            "type_verlof": ["Iets exotisch 2026"],
            "bkjr": pd.Series([2026], dtype="Int64"),
            "beginsaldo": [40.0],
        }
    )
    aanvulling = pd.DataFrame(
        {
            "mdw": ["1"],
            "naam": ["A"],
            "dv": ["1"],
            "bkjr": pd.Series([2026], dtype="Int64"),
            "aanvulling_pu": [5.0],
        }
    )
    result = bereken(verlof, aanvulling, {})
    assert "Iets exotisch" in result.onbekende_basis_types
