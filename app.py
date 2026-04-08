"""Streamlit-webapp 'Waarde vakantiedag'.

Upload een verlofsaldi-export en een aanvulling-per-uur-export, controleer per
basis-verloftype of het aanvullingsrechtig is, en bekijk per medewerker het
gewogen-gemiddelde aanvullingstarief — op het scherm én als xlsx-download.
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from verlofwaarde.calculator import bereken, unieke_basis_types
from verlofwaarde.export import build_resultaat_xlsx
from verlofwaarde.instellingen import apply_defaults
from verlofwaarde.parsing import read_aanvulling, read_verlofsaldi

st.set_page_config(
    page_title="Waarde vakantiedag",
    page_icon=":palm_tree:",
    layout="wide",
)


def _check_password() -> bool:
    """Eenvoudige password-gate via `st.secrets["APP_PASSWORD"]`.

    Geen secret geconfigureerd → app is open (handig voor lokale dev).
    """
    expected = st.secrets.get("APP_PASSWORD") if hasattr(st, "secrets") else None
    if not expected:
        return True

    if st.session_state.get("authenticated"):
        return True

    st.title("Waarde vakantiedag")
    st.write("Voer het wachtwoord in om verder te gaan.")
    pwd = st.text_input("Wachtwoord", type="password")
    if st.button("Inloggen"):
        if pwd == expected:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Onjuist wachtwoord.")
    return False


def _format_resultaat(df: pd.DataFrame) -> pd.DataFrame:
    """Geef een kopie van het resultaat terug met afgeronde waarden voor weergave."""
    out = df.copy()
    out["Totaal uren"] = out["Totaal uren"].round(2)
    out["Totaal aanvulling (€)"] = out["Totaal aanvulling (€)"].round(2)
    out["Gewogen aanvulling p/u (€)"] = out["Gewogen aanvulling p/u (€)"].round(4)
    return out


def main() -> None:
    if not _check_password():
        return

    st.title("Waarde vakantiedag — gewogen aanvulling p/u")
    st.markdown(
        """
Bereken per medewerker één gewogen-gemiddelde **aanvulling tijdens verlof p/u**
op basis van het beginsaldo per opbouwjaar en de aanvulling p/u die in dát
opbouwjaar gold. Het resulterende tarief geldt vervolgens voor élk verlofuur
dat dit jaar wordt opgenomen.
"""
    )

    st.header("1. Verlofsaldi-export")
    verlof_file = st.file_uploader(
        "Upload het verlofsaldi-export (.xlsx)",
        type=["xlsx"],
        key="verlof_uploader",
    )

    st.header("2. Aanvulling per uur")
    aanvulling_file = st.file_uploader(
        "Upload het aanvulling-per-uur-export (.xlsx)",
        type=["xlsx"],
        key="aanvulling_uploader",
    )

    if not (verlof_file and aanvulling_file):
        st.info("Upload beide bestanden om door te gaan.")
        return

    try:
        verlof_df = read_verlofsaldi(verlof_file)
    except ValueError as exc:
        st.error(f"Verlofsaldi-bestand kan niet worden gelezen: {exc}")
        return

    try:
        aanvulling_df = read_aanvulling(aanvulling_file)
    except ValueError as exc:
        st.error(f"Aanvulling-bestand kan niet worden gelezen: {exc}")
        return

    st.success(
        f"{len(verlof_df)} verlofsaldi-regels en {len(aanvulling_df)} "
        f"aanvulling-regels ingelezen."
    )

    st.header("3. Instellingen — wel of niet aanvullingsrechtig")
    st.markdown(
        "Per **basis-verloftype** (jaartal weggehaald) bepaalt u of het "
        "aanvullingsrechtig is. Niet-aanvullingsrechtige uren tellen mee in "
        "de noemer maar met €0 in de teller — dat verlaagt het gewogen "
        "gemiddelde."
    )

    basis_types = unieke_basis_types(verlof_df)
    if not basis_types:
        st.warning("Geen basis-verloftypes gevonden in de upload.")
        return

    defaults = apply_defaults(basis_types)
    if "instellingen_df" not in st.session_state or set(
        st.session_state["instellingen_df"]["Basis-type"]
    ) != set(basis_types):
        st.session_state["instellingen_df"] = pd.DataFrame(
            [{"Basis-type": bt, "Aanvullingsrechtig": defaults[bt]} for bt in basis_types]
        )

    edited = st.data_editor(
        st.session_state["instellingen_df"],
        column_config={
            "Basis-type": st.column_config.TextColumn("Basis-type", disabled=True),
            "Aanvullingsrechtig": st.column_config.CheckboxColumn(
                "Aanvullingsrechtig", default=False
            ),
        },
        hide_index=True,
        use_container_width=True,
        key="instellingen_editor",
    )
    st.session_state["instellingen_df"] = edited

    instellingen = {
        row["Basis-type"]: bool(row["Aanvullingsrechtig"]) for _, row in edited.iterrows()
    }

    st.header("4. Resultaat")
    if st.button("Bereken", type="primary"):
        result = bereken(verlof_df, aanvulling_df, instellingen)
        st.session_state["result"] = result

    result = st.session_state.get("result")
    if result is None:
        st.info("Klik op 'Bereken' om het gewogen gemiddelde te berekenen.")
        return

    if result.onbekende_basis_types:
        st.warning(
            "Onbekende basis-types staan op niet-aanvullingsrechtig: "
            + ", ".join(result.onbekende_basis_types)
            + ". Pas de instellingen aan en herbereken."
        )
    if not result.ontbrekende_aanvullingen.empty:
        st.warning(
            f"{len(result.ontbrekende_aanvullingen)} regel(s) hadden geen "
            "matchende aanvulling p/u in de aanvullingstabel — daar is met "
            "€0 gerekend. Zie 'Detail per regel' hieronder."
        )

    weergave = _format_resultaat(result.resultaat)
    st.dataframe(
        weergave,
        hide_index=True,
        use_container_width=True,
        column_config={
            "Totaal uren": st.column_config.NumberColumn(format="%.2f"),
            "Totaal aanvulling (€)": st.column_config.NumberColumn(format="€ %.2f"),
            "Gewogen aanvulling p/u (€)": st.column_config.NumberColumn(format="€ %.4f"),
        },
    )

    st.download_button(
        label="Download resultaat als Excel",
        data=build_resultaat_xlsx(result.resultaat),
        file_name="waarde_vakantiedag_resultaat.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Detail per regel"):
        st.dataframe(
            result.detail,
            hide_index=True,
            use_container_width=True,
        )

    if not result.ontbrekende_aanvullingen.empty:
        with st.expander("Regels zonder aanvulling p/u in de aanvullingstabel"):
            st.dataframe(
                result.ontbrekende_aanvullingen,
                hide_index=True,
                use_container_width=True,
            )


if __name__ == "__main__":
    main()
