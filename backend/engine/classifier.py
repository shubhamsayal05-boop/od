"""Event classification engine — Python port of conditionSDV (Load_Data.bas).

Rule semantics (blueprint §9):
- Rules evaluated in ORDRE; first fully-matching active rule wins.
- VALEUR ';'-separated list = OR across the list; literal VIDE = empty string.
- Rows sharing a group number are OR'd; different groups are AND'd.
- Missing channel ("FFF" sentinel) -> the condition auto-passes.
"""

MISSING = "__FFF__"

POSITIVE_OPS = {"CONTIENT", "EGAL A", "INFERIEUR A", "INFERIEUR OU EGAL A",
                "SUPERIEUR A", "SUPERIEUR OU EGAL A"}
NEGATIVE_OPS = {"NE CONTIENT PAS", "DIFFERENT DE"}

# Semantic aliases used by priorisation grids / chart params that do not
# always match TRIE "Family, Channel" keys in acquisition files.
CHANNEL_ALIASES = {
    "vehicle speed": (
        "p: vehicle speed, vehicle speed",
        "vehicle speed",
        "p: vehicle_speed_act, vehicle_speed_act",
    ),
    "accelerationchassis": (
        "accelerationchassis",
        "val_a_tstart, accelerationchassis",
        "p: ax level start, ax level start",
    ),
    "throttle position": (
        "val_a_tstart, throttle position",
        "max, throttle position",
        "p: pedal, pedal",
        "p: pedal_act, pedal_act",
    ),
    "max throttle position": (
        "max, throttle position",
        "val_a_tstart, throttle position",
        "p: pedal, pedal",
    ),
    "pedal change time": (
        "p: pedal change time, pedal change time",
        "pedal change time",
    ),
    "delta pedal": (
        "p: delta pedal, delta pedal",
        "delta pedal",
    ),
    "sub event name": (
        "sous situation de vie, sub event name",
        "sub event name",
    ),
    "gear new": (
        "p: gear end, gear end",
        "p: gear new, gear new",
        "gear new",
    ),
}


def _to_float(v):
    try:
        return float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return None


def _check_single(value, op, target):
    sval = "" if value is None else str(value).strip().lower()
    starget = str(target).strip().lower()
    if starget == "vide":
        starget = ""
    if op == "CONTIENT":
        return starget in sval if starget else sval == ""
    if op == "NE CONTIENT PAS":
        return starget not in sval if starget else sval != ""
    if op == "EGAL A":
        return sval == starget
    if op == "DIFFERENT DE":
        return sval != starget
    fv, ft = _to_float(value), _to_float(target)
    if fv is None or ft is None:
        return False
    if op == "INFERIEUR A":
        return fv < ft
    if op == "INFERIEUR OU EGAL A":
        return fv <= ft
    if op == "SUPERIEUR A":
        return fv > ft
    if op == "SUPERIEUR OU EGAL A":
        return fv >= ft
    return False


def check_condition(value, op, target):
    """OR across ';'-list for positive ops, AND for negative ops."""
    if value == MISSING:
        return True  # missing-channel auto-pass
    parts = [p.strip() for p in str(target).split(";") if p.strip() != ""]
    if not parts:
        parts = [""]
    if op in NEGATIVE_OPS:
        return all(_check_single(value, op, p) for p in parts)
    return any(_check_single(value, op, p) for p in parts)


def _norm(s):
    return str(s).replace("\u00b7", ".").strip().lower()


def get_channel(channels, name):
    """Resolve a semantic channel name against 'Family, Channel' keys."""
    if not channels or name is None:
        return MISSING
    if name in channels:
        return channels[name]
    lname = _norm(name)

    def _lookup(target):
        for key, val in channels.items():
            k = _norm(key)
            if k == target:
                return val
            if "," in k and k.split(",", 1)[1].strip() == target:
                return val
        return None

    hit = _lookup(lname)
    if hit is not None:
        return hit
    # also allow matching just the channel part of the requested name
    if "," in lname:
        part = lname.split(",", 1)[1].strip()
        hit = _lookup(part)
        if hit is not None:
            return hit
    for alias in CHANNEL_ALIASES.get(lname, ()):
        hit = _lookup(alias)
        if hit is not None:
            return hit
    return MISSING


def classify_event(channels, rules):
    """Return the SDV name of the first fully-matching rule, or ''. """
    for rule in sorted(rules, key=lambda r: r.get("order", 0)):
        if not rule.get("active", True):
            continue
        groups = {}
        for cond in rule.get("conditions", []):
            groups.setdefault(cond.get("group", 1), []).append(cond)
        matched = True
        for _, conds in sorted(groups.items()):
            # within group: OR (findV latches a hit)
            group_ok = False
            for cond in conds:
                col = cond.get("column", "")
                # resolve the full "Family, Channel" key first, then channel part
                value = get_channel(channels, col)
                if value == MISSING and "," in col:
                    value = get_channel(channels, col.split(",", 1)[1].strip())
                if check_condition(value, cond.get("op", "CONTIENT"), cond.get("value", "")):
                    group_ok = True
                    break
            if not group_ok:
                matched = False
                break
        if matched and rule.get("conditions"):
            return rule["sdv"]
    return ""
