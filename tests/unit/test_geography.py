from app.core.geography import (
    MUNICIPALITY_TO_ZONE,
    POLO_TO_ZONE,
    ZONA_NAO_CLASSIFICADA,
    zone_for,
)


class TestZoneForPoloPriority:
    def test_polo_match_wins_over_municipality(self):
        # São Paulo município covers many zones — the polo (e.g. Pirituba) is the
        # only field granular enough to disambiguate Norte from Sul from Oeste.
        assert zone_for(municipality="São Paulo", polo="Pirituba") == "Zona Norte"
        assert zone_for(municipality="São Paulo", polo="Santo Amaro") == "Zona Sul"

    def test_polo_lookup_is_case_insensitive(self):
        # Polo names live in the xlsx data sheet name as uppercase ("DADOS - PIRITUBA")
        # and title-case after extraction; both must resolve to the same zone.
        assert zone_for(polo="PIRITUBA") == zone_for(polo="pirituba") == zone_for(polo="Pirituba")

    def test_fixture_polos_resolve_to_their_known_zones(self):
        # Pin the six polos that appear in Model/ fixtures so a careless edit to
        # POLO_TO_ZONE can't silently re-classify a polo that already ships in
        # client reports.
        assert zone_for(polo="Pirituba") == "Zona Norte"
        assert zone_for(polo="Santana") == "Zona Norte"
        assert zone_for(polo="Freguesia do Ó") == "Zona Norte"
        assert zone_for(polo="Pimentas") == "Zona Leste Metropolitana"
        assert zone_for(polo="Gopoúva") == "Zona Leste Metropolitana"
        assert zone_for(polo="Extremo Norte") == "Extremo Norte Metropolitana"


class TestZoneForMunicipalityFallback:
    def test_municipality_fallback_used_when_polo_unknown(self):
        # A polo we haven't catalogued in a known município still resolves —
        # município-level mapping handles the long tail.
        assert (
            zone_for(municipality="Guarulhos", polo="Novo Polo XYZ") == "Zona Leste Metropolitana"
        )
        assert zone_for(municipality="Santo André", polo="Outro") == "Grande ABC"

    def test_municipality_lookup_is_case_insensitive_and_strips(self):
        assert zone_for(municipality=" GUARULHOS ") == "Zona Leste Metropolitana"

    def test_municipality_handles_accent_variants_for_mairipora(self):
        # The xlsx files in Model/ sometimes drop the tilde on "Mairipora" vs
        # "Mairiporã"; both must resolve to the same zone.
        assert zone_for(municipality="Mairiporã") == zone_for(municipality="Mairipora")


class TestZoneForUnknown:
    def test_returns_sentinel_when_nothing_matches(self):
        assert zone_for(municipality="Marte", polo="Cratera") == ZONA_NAO_CLASSIFICADA

    def test_empty_inputs_return_sentinel(self):
        assert zone_for(municipality="", polo="") == ZONA_NAO_CLASSIFICADA
        assert zone_for() == ZONA_NAO_CLASSIFICADA


class TestMapShape:
    def test_polo_keys_are_lowercase(self):
        # Lookup normalises to lower; if a constant has an uppercase key it
        # would never match and become dead code.
        for key in POLO_TO_ZONE:
            assert key == key.lower(), f"polo key {key!r} must be lowercase"

    def test_municipality_keys_are_lowercase(self):
        for key in MUNICIPALITY_TO_ZONE:
            assert key == key.lower(), f"municipality key {key!r} must be lowercase"

    def test_all_known_sp_capital_zones_have_at_least_one_polo(self):
        # User requirement: "all Zones from SP" — Norte/Sul/Leste/Oeste/Centro
        # must each have ≥ 1 polo so a freshly-onboarded user immediately sees
        # the zone label even before they upload data covering that zone.
        zones_present = set(POLO_TO_ZONE.values())
        for required_zone in ("Zona Norte", "Zona Sul", "Zona Leste", "Zona Oeste", "Zona Centro"):
            assert required_zone in zones_present, f"{required_zone} missing from POLO_TO_ZONE"
