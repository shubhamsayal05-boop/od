"""Pure unit tests for the DriveScope / ODRIV scoring engine (no Mongo)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from engine import config_loader, scoring, importer
from engine.classifier import classify_event, get_channel, check_condition


def _cfg():
    return config_loader.load_seed_sections()


def test_seed_sections_complete():
    cfg = _cfg()
    for key in config_loader.SECTION_FILES:
        assert key in cfg
    assert len(cfg["catalog"]) >= 60
    assert len(cfg["targets"]) > 1000


def test_build_targets_lookup_fallback_for_bad_priority():
    cfg = _cfg()
    project = {"version": "4.6", "priority": "P1", "mode": "AUTO"}
    lookup = config_loader.build_targets_lookup(cfg["targets"], project)
    assert len(lookup) >= 50
    # case-insensitive criterion resolve
    sample_sdv = next(iter(lookup))
    first_crit = next(iter(lookup[sample_sdv]))
    row = config_loader.resolve_criteria_row(lookup[sample_sdv], first_crit.upper())
    assert row is not None


def test_resolve_priorisation_alias():
    cfg = _cfg()
    prior_up = {k.strip().upper(): v for k, v in cfg["priorisation"].items()}
    a = config_loader.resolve_priorisation(prior_up, "DECEL - TRANS TO CST SPD - COLD")
    b = config_loader.resolve_priorisation(prior_up, "DECEL TRANSITION TO CST SPEED - COLD")
    assert a and b
    assert a[0].get("speed_cols") == b[0].get("speed_cols")


def test_get_channel_aliases():
    channels = {
        "P: vehicle speed, vehicle Speed": 42.5,
        "max, Throttle Position": 88.0,
        "AccelerationChassis": -1.2,
        "Sous situation de vie, Sub Event Name": "tip-in #3",
    }
    assert get_channel(channels, "Vehicle Speed") == 42.5
    assert get_channel(channels, "Max Throttle Position") == 88.0
    assert get_channel(channels, "AccelerationChassis") == -1.2
    assert get_channel(channels, "Sub Event Name") == "tip-in #3"


def test_check_condition_or_list():
    assert check_condition("AUTO", "EGAL A", "AUTO;SPORT")
    assert not check_condition("ECO", "EGAL A", "AUTO;SPORT")
    assert check_condition("zzz", "NE CONTIENT PAS", "foo;bar")


def test_criterion_index_colors():
    # note below waterline -> RED penalty
    idx, color = scoring.criterion_index(5.0, wl=6.0, t=8.0, criticity=2)
    assert color == "RED"
    assert idx < 0
    # note above target -> GREEN, zero penalty
    idx2, color2 = scoring.criterion_index(9.0, wl=6.0, t=8.0, criticity=2)
    assert color2 == "GREEN"
    assert idx2 == 0.0


def test_classify_and_score_demo_pipeline():
    cfg = _cfg()
    canon = config_loader.build_canon_map(cfg["structure"], cfg["catalog"])
    raw = importer.generate_sample_events(cfg["definitions"], cfg["structure"], canon)
    docs = []
    for i, channels in enumerate(raw):
        sdv = classify_event(channels, cfg["definitions"])
        name = config_loader.canon_name(sdv, canon) if sdv else ""
        if name:
            docs.append({"id": str(i), "sdv": name, "channels": channels})
    assert len(docs) > 400
    project = {
        "version": "4.6", "priority": "PREMIUM", "mode": "AUTO",
        "odriv_milestone": "MDL2",
    }
    lookup = config_loader.build_targets_lookup(cfg["targets"], project)
    updates, sdv_results, glob = scoring.calculate_rating(
        project, docs, cfg, lookup, canon)
    assert len(sdv_results) >= 50
    assert glob["driv"] is not None and "index" in glob["driv"]
    assert glob["dyn"] is not None and glob["dyn"]["verdict"] in {
        "Low Risk", "Medium Risk", "High Risk",
    }
    assert updates  # events scored
