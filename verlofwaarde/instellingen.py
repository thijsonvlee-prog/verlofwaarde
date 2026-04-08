"""Default-instellingen voor welke basis-verloftypes aanvullingsrechtig zijn.

De gebruiker kan deze defaults in de UI overrulen voor de huidige sessie.
Onbekende basis-types krijgen `False` als default — de gebruiker moet bewust
beslissen of een nieuw type aanvullingsrechtig is.
"""

from __future__ import annotations

DEFAULT_AANVULLINGSRECHTIG: dict[str, bool] = {
    "Wettelijke rechten": True,
    "Bovenwettelijke rechten": True,
    "Bovenwettelijke rechten (aanv.)": True,
}


def apply_defaults(basis_types: list[str]) -> dict[str, bool]:
    """Map elke gevonden basis-type naar zijn default-aanvullingsrechtigheid.

    Bekende types krijgen hun default uit `DEFAULT_AANVULLINGSRECHTIG`.
    Onbekende types krijgen `False` zodat ze in de UI duidelijk zichtbaar
    zijn als 'nieuw, te beoordelen'.
    """
    return {bt: DEFAULT_AANVULLINGSRECHTIG.get(bt, False) for bt in basis_types}
