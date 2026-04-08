"""Kernberekening: per medewerker een gewogen gemiddelde aanvulling p/u.

De rekenregel staat in detail in het plan onder 'Berekeningslogica'. Kort:

1. Skip rijen met Beginsaldo == 0.
2. Extract opbouwjaar uit `Type verlof` (4-cijferig jaartal). Geen jaar → skip.
3. Strip jaartal om het basis-type te krijgen ("Bovenwettelijke rechten 2025
   (aanv.)" → "Bovenwettelijke rechten (aanv.)").
4. Lookup of het basis-type aanvullingsrechtig is. Niet → effectieve aanvulling
   wordt 0,00 (telt wel mee in de noemer).
5. Lookup aanvulling p/u op (Mdw, opbouwjaar). Niet gevonden → 0,00.
6. Gewicht = Beginsaldo × effectieve aanvulling p/u.

Aggregatie per Mdw: gewogen gemiddelde = som(gewicht) / som(beginsaldo).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

_JAAR_RE = re.compile(r"\b(20\d{2})\b")


@dataclass
class BerekeningResultaat:
    """Resultaat van een berekening met optionele waarschuwingen voor de UI."""

    resultaat: pd.DataFrame
    detail: pd.DataFrame
    ontbrekende_aanvullingen: pd.DataFrame
    onbekende_basis_types: list[str]


def extract_opbouwjaar(type_verlof: str) -> int | None:
    """Geef het 4-cijferige opbouwjaar terug of None als er geen jaartal in de naam zit."""
    if type_verlof is None:
        return None
    match = _JAAR_RE.search(str(type_verlof))
    return int(match.group(1)) if match else None


def extract_basis_type(type_verlof: str) -> str:
    """Strip het 4-cijferige jaartal en normaliseer spaties.

    Voorbeelden:
        "Bovenwettelijke rechten 2025"          → "Bovenwettelijke rechten"
        "Bovenwettelijke rechten 2025 (aanv.)"  → "Bovenwettelijke rechten (aanv.)"
        "Negatief verlof"                       → "Negatief verlof"
    """
    if type_verlof is None:
        return ""
    s = _JAAR_RE.sub("", str(type_verlof))
    return re.sub(r"\s+", " ", s).strip()


def unieke_basis_types(verlof_df: pd.DataFrame) -> list[str]:
    """Geef een gesorteerde lijst van unieke basis-types in de verlofsaldi-upload."""
    if verlof_df.empty:
        return []
    types = verlof_df["type_verlof"].dropna().map(extract_basis_type)
    return sorted({t for t in types if t})


def _aggregate_aanvulling(aanvulling_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregeer aanvulling p/u per (mdw, bkjr); meerdere DV's → eerste hit."""
    if aanvulling_df.empty:
        return aanvulling_df.assign(aanvulling_pu=pd.Series(dtype=float))
    return (
        aanvulling_df.dropna(subset=["bkjr"])
        .groupby(["mdw", "bkjr"], as_index=False, sort=False)["aanvulling_pu"]
        .first()
    )


def bereken(
    verlof_df: pd.DataFrame,
    aanvulling_df: pd.DataFrame,
    instellingen: dict[str, bool],
) -> BerekeningResultaat:
    """Voer de gewogen-gemiddelde-berekening uit.

    Returns:
        BerekeningResultaat met de eindtabel, een detail-DataFrame voor debug,
        een lijst (mdw, opbouwjaar) zonder match in de aanvullingstabel, en
        een lijst basis-types die niet in de instellingen voorkomen.
    """
    df = verlof_df.copy()
    df["beginsaldo"] = pd.to_numeric(df["beginsaldo"], errors="coerce").fillna(0.0)
    df["opbouwjaar"] = df["type_verlof"].map(extract_opbouwjaar).astype("Int64")
    df["basis_type"] = df["type_verlof"].map(extract_basis_type)

    # Skip wat niet in de berekening hoort: geen jaartal, of beginsaldo 0.
    actief = df[df["opbouwjaar"].notna() & (df["beginsaldo"] != 0)].copy()

    # Onbekende basis-types worden als niet-aanvullingsrechtig behandeld; we
    # geven de UI een lijst zodat de gebruiker bewust kan kiezen.
    aanwezige_types = sorted({t for t in actief["basis_type"] if t})
    onbekende = [t for t in aanwezige_types if t not in instellingen]
    actief["aanvullingsrechtig"] = actief["basis_type"].map(
        lambda t: bool(instellingen.get(t, False))
    )

    # Lookup aanvulling p/u op (mdw, opbouwjaar). DV wordt genegeerd zoals
    # afgesproken.
    aanvulling_unique = _aggregate_aanvulling(aanvulling_df)
    actief["opbouwjaar_int"] = actief["opbouwjaar"].astype("Int64")
    actief = actief.merge(
        aanvulling_unique[["mdw", "bkjr", "aanvulling_pu"]].rename(
            columns={"bkjr": "opbouwjaar_int"}
        ),
        on=["mdw", "opbouwjaar_int"],
        how="left",
    )

    ontbrekend_mask = actief["aanvulling_pu"].isna()
    ontbrekende_aanvullingen = (
        actief.loc[ontbrekend_mask, ["mdw", "naam", "type_verlof", "opbouwjaar"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    actief["aanvulling_pu"] = actief["aanvulling_pu"].fillna(0.0)
    actief["effectieve_aanvulling_pu"] = actief["aanvulling_pu"].where(
        actief["aanvullingsrechtig"], 0.0
    )
    actief["gewicht"] = actief["beginsaldo"] * actief["effectieve_aanvulling_pu"]

    # Aggregeer per medewerker. Naam = eerste niet-lege naam voor die mdw.
    naam_per_mdw = (
        df[df["naam"].astype(str).str.strip() != ""]
        .groupby("mdw")["naam"]
        .first()
    )

    grouped = (
        actief.groupby("mdw", as_index=False, sort=False)
        .agg(
            totaal_uren=("beginsaldo", "sum"),
            totaal_aanvulling=("gewicht", "sum"),
        )
    )
    grouped["naam"] = grouped["mdw"].map(naam_per_mdw).fillna("")
    grouped["gewogen_aanvulling_pu"] = grouped.apply(
        lambda r: (r["totaal_aanvulling"] / r["totaal_uren"]) if r["totaal_uren"] else 0.0,
        axis=1,
    )

    resultaat = grouped.rename(
        columns={
            "mdw": "Mdw.",
            "naam": "Naam",
            "totaal_uren": "Totaal uren",
            "totaal_aanvulling": "Totaal aanvulling (€)",
            "gewogen_aanvulling_pu": "Gewogen aanvulling p/u (€)",
        }
    )[
        [
            "Mdw.",
            "Naam",
            "Totaal uren",
            "Totaal aanvulling (€)",
            "Gewogen aanvulling p/u (€)",
        ]
    ].sort_values("Mdw.").reset_index(drop=True)

    detail = actief[
        [
            "mdw",
            "naam",
            "type_verlof",
            "opbouwjaar",
            "basis_type",
            "aanvullingsrechtig",
            "beginsaldo",
            "aanvulling_pu",
            "effectieve_aanvulling_pu",
            "gewicht",
        ]
    ].rename(
        columns={
            "mdw": "Mdw.",
            "naam": "Naam",
            "type_verlof": "Type verlof",
            "opbouwjaar": "Opbouwjaar",
            "basis_type": "Basis-type",
            "aanvullingsrechtig": "Aanvullingsrechtig",
            "beginsaldo": "Beginsaldo (uren)",
            "aanvulling_pu": "Aanvulling p/u (€)",
            "effectieve_aanvulling_pu": "Effectieve aanvulling p/u (€)",
            "gewicht": "Gewicht (€)",
        }
    ).reset_index(drop=True)

    return BerekeningResultaat(
        resultaat=resultaat,
        detail=detail,
        ontbrekende_aanvullingen=ontbrekende_aanvullingen,
        onbekende_basis_types=onbekende,
    )
