"""SP geographic classification: polo / município → zone.

Used by the dashboard to group multi-polo batch uploads into a Zone →
Município → Polo drill-down. Maps are static (no DB) because Sabesp's
operational structure changes slowly enough that a code change is the
right scope of friction.

Lookup priority — polo first, município second — because the município
"São Paulo" alone is too coarse to distinguish Norte (Pirituba, Santana)
from Sul (Santo Amaro) from Oeste (Lapa). Polo names map to a single
zone unambiguously. Outside the capital the polo dictionary tends to be
empty, so município fills in (Guarulhos, the ABC, the Serra municípios).

Unknown inputs return ``ZONA_NAO_CLASSIFICADA`` so the dashboard still
renders something the user can correct by adding to the dict and shipping
a new commit.
"""

from __future__ import annotations

ZONA_NAO_CLASSIFICADA = "Não Classificada"

# Polo → Zone. Polo names live in xlsx data sheet names as uppercase
# ("DADOS - PIRITUBA"); the lookup lower-cases on the way in.
POLO_TO_ZONE: dict[str, str] = {
    # São Paulo capital — Norte
    "pirituba": "Zona Norte",
    "santana": "Zona Norte",
    "freguesia do ó": "Zona Norte",
    "tucuruvi": "Zona Norte",
    "vila maria": "Zona Norte",
    "casa verde": "Zona Norte",
    "jaçanã": "Zona Norte",
    "perus": "Zona Norte",
    # São Paulo capital — Sul
    "santo amaro": "Zona Sul",
    "capela do socorro": "Zona Sul",
    "cidade ademar": "Zona Sul",
    "campo limpo": "Zona Sul",
    "m'boi mirim": "Zona Sul",
    "ipiranga": "Zona Sul",
    "jabaquara": "Zona Sul",
    "vila mariana": "Zona Sul",
    "parelheiros": "Zona Sul",
    # São Paulo capital — Leste
    "penha": "Zona Leste",
    "itaquera": "Zona Leste",
    "são miguel paulista": "Zona Leste",
    "itaim paulista": "Zona Leste",
    "ermelino matarazzo": "Zona Leste",
    "vila prudente": "Zona Leste",
    "aricanduva": "Zona Leste",
    "moóca": "Zona Leste",
    "mooca": "Zona Leste",
    "guaianases": "Zona Leste",
    "cidade tiradentes": "Zona Leste",
    "são mateus": "Zona Leste",
    "sapopemba": "Zona Leste",
    # São Paulo capital — Oeste
    "lapa": "Zona Oeste",
    "butantã": "Zona Oeste",
    "pinheiros": "Zona Oeste",
    # São Paulo capital — Centro
    "sé": "Zona Centro",
    "se": "Zona Centro",
    "república": "Zona Centro",
    # Greater SP — Guarulhos polos
    "pimentas": "Zona Leste Metropolitana",
    "gopoúva": "Zona Leste Metropolitana",
    "gopouva": "Zona Leste Metropolitana",
    # Greater SP — Serra / Extremo Norte
    "extremo norte": "Extremo Norte Metropolitana",
}

# Município → Zone. Fallback when the polo isn't in POLO_TO_ZONE — covers
# the long tail of metro-area suburbs whose neighbourhood-polo split isn't
# catalogued. Don't put "São Paulo" here; the capital is too coarse.
MUNICIPALITY_TO_ZONE: dict[str, str] = {
    # Guarulhos region (east of capital)
    "guarulhos": "Zona Leste Metropolitana",
    # Serra municípios (north of capital)
    "caieiras": "Extremo Norte Metropolitana",
    "cajamar": "Extremo Norte Metropolitana",
    "francisco morato": "Extremo Norte Metropolitana",
    "franco da rocha": "Extremo Norte Metropolitana",
    "mairiporã": "Extremo Norte Metropolitana",
    "mairipora": "Extremo Norte Metropolitana",
    # Grande ABC
    "santo andré": "Grande ABC",
    "santo andre": "Grande ABC",
    "são bernardo do campo": "Grande ABC",
    "sao bernardo do campo": "Grande ABC",
    "são caetano do sul": "Grande ABC",
    "sao caetano do sul": "Grande ABC",
    "diadema": "Grande ABC",
    "mauá": "Grande ABC",
    "maua": "Grande ABC",
    "ribeirão pires": "Grande ABC",
    "ribeirao pires": "Grande ABC",
    "rio grande da serra": "Grande ABC",
    # Oeste metropolitana
    "osasco": "Zona Oeste Metropolitana",
    "barueri": "Zona Oeste Metropolitana",
    "carapicuíba": "Zona Oeste Metropolitana",
    "carapicuiba": "Zona Oeste Metropolitana",
    "santana de parnaíba": "Zona Oeste Metropolitana",
    "santana de parnaiba": "Zona Oeste Metropolitana",
    "jandira": "Zona Oeste Metropolitana",
    "itapevi": "Zona Oeste Metropolitana",
    "cotia": "Zona Oeste Metropolitana",
    "vargem grande paulista": "Zona Oeste Metropolitana",
    # Alto Tietê / leste estendido
    "ferraz de vasconcelos": "Vale do Tietê",
    "itaquaquecetuba": "Vale do Tietê",
    "poá": "Vale do Tietê",
    "poa": "Vale do Tietê",
    "suzano": "Vale do Tietê",
    "mogi das cruzes": "Vale do Tietê",
    "arujá": "Vale do Tietê",
    "aruja": "Vale do Tietê",
    "biritiba mirim": "Vale do Tietê",
    "salesópolis": "Vale do Tietê",
    "salesopolis": "Vale do Tietê",
    # Sul metropolitana
    "embu das artes": "Zona Sul Metropolitana",
    "taboão da serra": "Zona Sul Metropolitana",
    "taboao da serra": "Zona Sul Metropolitana",
    "itapecerica da serra": "Zona Sul Metropolitana",
    "são lourenço da serra": "Zona Sul Metropolitana",
    "sao lourenco da serra": "Zona Sul Metropolitana",
    "embu-guaçu": "Zona Sul Metropolitana",
    "embu-guacu": "Zona Sul Metropolitana",
}


def zone_for(municipality: str = "", polo: str = "") -> str:
    """Resolve a (município, polo) pair to a zone label.

    Polo wins when both are supplied — see module docstring for why.
    Inputs are stripped and lower-cased before lookup. Unknown inputs
    return ``ZONA_NAO_CLASSIFICADA`` so dashboard rendering never crashes
    on a polo we haven't onboarded yet.
    """
    polo_key = (polo or "").strip().lower()
    if polo_key and polo_key in POLO_TO_ZONE:
        return POLO_TO_ZONE[polo_key]
    muni_key = (municipality or "").strip().lower()
    if muni_key and muni_key in MUNICIPALITY_TO_ZONE:
        return MUNICIPALITY_TO_ZONE[muni_key]
    return ZONA_NAO_CLASSIFICADA
