"""ODRIV scoring engine — faithful Python port of the §11-§14 compute chain,
driven by the real configuration extracted from ODRIV_v29_2_1_AT.xlsm.

Chain per SDV sheet: agreement index per criterion (Rating_IndiceAgrement),
per-event aggregation "C1 non-moyennés C2 moyennés", priorisation via the
CONFIGURATIONS SEETINGS speed x acceleration grids, point coloring vs
Waterline/Target, occurrence criticity (cfg_criticity + Seuil), SDV index
J5/BQ5 = 100*(1+sum(indice_occ)/N)^PUISS, status dots vs the SETTINGS
milestone matrices, weighted rate of low points, NoteGlobale_2 3x3 verdict.
"""
from .classifier import get_channel, _to_float, _check_single

COLOR_ORDER = {"GREEN": 0, "YELLOW": 1, "RED": 2}
VERDICT_MATRIX = [
    ["Low Risk", "Low Risk", "Medium Risk"],
    ["Medium Risk", "Medium Risk", "High Risk"],
    ["High Risk", "High Risk", "High Risk"],
]


def criterion_index(note, wl, t, criticity):
    """Agreement index for one criterion -> (index, color) or None.

    C = (3-c)/2 ; ZF = COEF1*WL + COEF2*T = 2*WL - T ; piecewise penalty.
    Participates only when WL>0, T>0, 0<note<=10.
    """
    try:
        criticity = int(criticity)
    except (TypeError, ValueError):
        return None
    C = (3 - criticity) / 2.0
    if C <= 0:
        return None
    if wl is None or t is None or wl <= 0 or t <= 0:
        return None
    if note is None or not (0 < note <= 10):
        return None
    zf = 2 * wl - t
    denom = t - zf
    if denom == 0:
        return None
    color = "RED" if note < wl else ("YELLOW" if note < t else "GREEN")
    if note < zf:
        return -C, color
    note_t = 10 * (note - zf) / denom
    wl_t = 10 * (wl - zf) / denom
    t_t = 10.0
    if note_t < wl_t:
        idx = C * (2 * note_t - t_t - wl_t) / (t_t + wl_t)
    elif note_t < t_t:
        idx = C * (note_t - t_t) / (t_t + wl_t)
    else:
        idx = 0.0
    return idx, color


def priority_from_grid(channels, configs):
    """P1/P2/P3 lookup in the CONFIGURATIONS SEETINGS grid:
    columns = vehicle-speed lower bounds, rows = |acceleration| bands."""
    if not configs:
        return 3
    cfg = configs[0]
    cols = cfg.get("speed_cols") or []
    grid = cfg.get("grid") or []
    if not cols or not grid:
        return 3
    speed = _to_float(get_channel(channels, cfg.get("ref_ligne") or "Vehicle Speed"))
    accel = _to_float(get_channel(channels, cfg.get("ref_colonne") or "AccelerationChassis"))
    if speed is None:
        return 3
    ci = 0
    for i, c in enumerate(cols):
        if c is not None and speed >= c:
            ci = i
    a = abs(accel) if accel is not None else 0.0
    row = grid[-1]
    for g in grid:
        try:
            lo, hi = [float(x) for x in g["band"].split("-")]
        except ValueError:
            continue
        if lo <= a < hi:
            row = g
            break
    try:
        p = row["priorities"][ci]
    except (IndexError, KeyError):
        p = None
    try:
        p = int(p)
    except (TypeError, ValueError):
        p = 3
    return p if p in (1, 2, 3) else 3


def criticity_level(color, priority, indice, criticity_cfg):
    """F_criticity: event color+indice -> drive-rating level + occurrence class."""
    seuil = (criticity_cfg.get("seuil") or [-0.16, -0.16, -0.1])
    s = seuil[priority - 1] if priority - 1 < len(seuil) else -0.16
    if color == "RED":
        level = "Red +" if indice <= -0.8 else "Red"
    elif color == "YELLOW":
        level = "Orange" if indice < (s if s is not None else -0.16) else "Jaune"
    else:
        level = "Green"
    table = criticity_cfg.get("table") or {}
    occ = None
    if level in table and priority - 1 < len(table[level]):
        occ = table[level][priority - 1]
    return level, occ


def score_event(channels, criteria_rows, part, coefficients, prior_cfgs,
                criticity_cfg):
    """Full per-event computation for one part ('driv'/'dyn').

    criteria_rows: {criterion_name: targets_row} for this SDV.
    """
    key = "driv" if part == "driv" else "resp"
    c1, c2 = [], []
    crit_colors = {}
    for name, row in criteria_rows.items():
        res = criterion_index(_to_float(get_channel(channels, name)),
                              _to_float(row.get("wl")), _to_float(row.get("t")),
                              row.get(key))
        if res is None:
            continue
        idx, color = res
        crit_colors[name.replace(".", "\u00b7")] = color
        crit = int(row.get(key))
        if crit == 1:
            c1.append(idx)
        else:
            c2.append(idx)
    if not c1 and not c2:
        return None
    indice = sum(c1) + (sum(c2) / len(c2) if c2 else 0.0)
    indice = max(indice, -1.0)  # hard clamp
    color = max(crit_colors.values(), key=lambda c: COLOR_ORDER[c]) if crit_colors else "GREEN"
    priority = priority_from_grid(channels, prior_cfgs)
    coef = coefficients[str(priority)][color]
    level, occ = criticity_level(color, priority, indice, criticity_cfg)
    return {
        "indice": round(indice, 4),
        "indice_occ": round(indice * coef / 100.0, 4),
        "color": color,
        "priority": priority,
        "criticity": level,
        "occurrence": occ,
        "crit_colors": crit_colors,
    }


def sdv_index(event_scores, puiss):
    """IndiceSDV = 1 + sum(indice_occ)/N ; J5 = round(100*IndiceSDV^PUISS, 1)."""
    n = len(event_scores)
    if n == 0:
        return None
    base = 1.0 + sum(e["indice_occ"] for e in event_scores) / n
    return round(100 * (max(base, 0.0) ** puiss), 1)


def status_dots(event_scores, m_idx, block, coef_yellow, facteur_redplus):
    """Note_SDV: per-priority status dot vs the SDV's SETTINGS milestone block."""
    pct = block.get("pct", {})
    nmin = block.get("nmin", {})
    out = {}
    for p in (1, 2, 3):
        evs = [e for e in event_scores if e["priority"] == p]
        n = len(evs)
        if n == 0:
            out[str(p)] = "NONE"
            continue
        redplus = sum(1 for e in evs if e["criticity"] == "Red +")
        red = sum(1 for e in evs if e["color"] == "RED") - redplus
        yellow = sum(1 for e in evs if e["color"] == "YELLOW")
        red_mass = red + facteur_redplus * redplus
        pct_red = 100.0 * red_mass / n
        pct_orange = 100.0 * (red_mass + yellow * coef_yellow) / n
        thr_r = (pct.get("P%dR" % p) or [15, 10, 5, 0])[m_idx]
        thr_o = (pct.get("P%dO" % p) or [25, 20, 15, 10])[m_idx]
        nm_r = (nmin.get("P%dR" % p) or [1, 1, 1, 1])[m_idx] or 1
        nm_o = (nmin.get("P%dO" % p) or [1, 1, 1, 1])[m_idx] or 1
        if pct_red > (thr_r or 0) and (red + redplus) >= nm_r:
            out[str(p)] = "RED"
        elif pct_orange > (thr_o or 0) and (red + redplus + yellow) >= nm_o:
            out[str(p)] = "ORANGE"
        else:
            out[str(p)] = "GREEN"
    return out


def low_point_fraction(event_scores, coef_yellow, facteur_redplus):
    n = len(event_scores)
    if n == 0:
        return 0.0
    redplus = sum(1 for e in event_scores if e["criticity"] == "Red +")
    red = sum(1 for e in event_scores if e["color"] == "RED") - redplus
    yellow = sum(1 for e in event_scores if e["color"] == "YELLOW")
    return (red + facteur_redplus * redplus + yellow * coef_yellow) / n


def counts_table(event_scores):
    table = {}
    total = len(event_scores)
    for color in ("GREEN", "YELLOW", "RED"):
        for p in (1, 2, 3):
            k = "%s_P%d" % (color, p)
            c = sum(1 for e in event_scores
                    if e["color"] == color and e["priority"] == p)
            table[k] = c
    table["total"] = total
    for p in (1, 2, 3):
        table["P%d_total" % p] = sum(1 for e in event_scores if e["priority"] == p)
    return table


def global_verdict(index_value, rate_low, m_idx, thr, globalpuiss):
    """NoteGlobale_2 3x3 decision matrix."""
    x = (index_value / 100.0) ** globalpuiss if index_value else 0.0
    if rate_low < thr["taux_vert_orange"][m_idx]:
        row = 0
    elif rate_low <= thr["taux_orange_rouge"][m_idx]:
        row = 1
    else:
        row = 2
    if x > thr["index_vert_orange"][m_idx]:
        col = 0
    elif x >= thr["index_orange_rouge"][m_idx]:
        col = 1
    else:
        col = 2
    return VERDICT_MATRIX[row][col], round(x, 4)


def milestone_index(project, configurations):
    label = str(project.get("odriv_milestone") or "")
    m = (configurations.get("milestones") or {}).get(label)
    try:
        m = int(m)
    except (TypeError, ValueError):
        m = 4
    return max(0, min(3, m - 1))


def calculate_rating(project, events, cfg, targets_lookup, canon_map):
    """Afficher_Calcul master loop -> (event_updates, sdv_results, global_results)."""
    sg = cfg["settings_global"]
    coefficients = sg["coefficients"]
    puiss = sg["constants"]["PUISS"]
    gpuiss = sg["constants"]["GLOBALPUISS"]
    calculs = cfg["calculs"]
    blocks = cfg["settings_blocks"]
    blocks_up = {k.strip().upper(): v for k, v in blocks.items()}
    prior_up = {k.strip().upper(): v for k, v in cfg["priorisation"].items()}
    tv_rows = {r["sdv"].strip().upper(): r for r in cfg["target_vehicle"]["rows"]}
    calculs_up = {k.strip().upper(): v for k, v in calculs.get("sdv", {}).items()}
    catalog = {c["name"].strip().upper(): c for c in cfg["catalog"]}

    grade = (project.get("priority") or "PREMIUM").strip().upper()
    coef_orange = calculs["coef_orange_mainstream"] if grade == "MAINSTREAM" \
        else calculs["coef_orange_premium"]
    coef_yellow = 1.0 / coef_orange
    facteur_redplus = calculs.get("facteur_redplus") or 2
    m_idx = milestone_index(project, cfg["configurations"])

    by_sdv = {}
    for ev in events:
        by_sdv.setdefault(ev["sdv"], []).append(ev)

    event_updates = {}
    sdv_results = []
    parts_acc = {"driv": [], "dyn": []}

    for sdv_name, evs in by_sdv.items():
        up = sdv_name.strip().upper()
        if len(evs) <= 2:
            continue  # UpdateTab requires >2 events
        criteria_rows = targets_lookup.get(up, {})
        if not criteria_rows:
            continue
        block = blocks_up.get(up, {})
        prior_cfgs = prior_up.get(up, [])
        cat = catalog.get(up, {})
        result = {"name": sdv_name, "group": cat.get("group", ""),
                  "order": cat.get("order", 999), "n_events": len(evs)}
        for part in ("driv", "dyn"):
            scores = []
            lowest = None
            for ev in evs:
                sc = score_event(ev["channels"], criteria_rows, part,
                                 coefficients, prior_cfgs, cfg["criticity"])
                if sc is None:
                    continue
                scores.append(sc)
                event_updates.setdefault(ev["id"], {})[part] = sc
                if lowest is None or sc["indice_occ"] < lowest[0]:
                    sub = get_channel(ev["channels"], "Sub Event Name")
                    lowest = (sc["indice_occ"], str(sub) if sub != "__FFF__" else ev["id"][:8])
            if not scores:
                result[part] = None
                continue
            idx = sdv_index(scores, puiss)
            tvr = tv_rows.get(up, {})
            result[part] = {
                "index": idx,
                "target_index": tvr.get(part if part == "driv" else "dyn"),
                "status": status_dots(scores, m_idx, block, coef_yellow, facteur_redplus),
                "status_pred": status_dots(scores, 3, block, coef_yellow, facteur_redplus),
                "counts": counts_table(scores),
                "low_frac": round(low_point_fraction(scores, coef_yellow, facteur_redplus), 5),
                "lowest_event": (lowest[1] if lowest and lowest[0] < 0 else None),
            }
            sdv_cc = calculs_up.get(up, {})
            if idx is not None:
                parts_acc[part].append({
                    "sdv": sdv_name, "index": idx,
                    "low_frac": result[part]["low_frac"],
                    "weight": block.get("weight") or 10,
                    "taux": sdv_cc.get("taux") or 0.001,
                    "target": result[part]["target_index"],
                })
        sdv_results.append(result)

    global_results = {"milestone_idx": m_idx}
    for part in ("driv", "dyn"):
        acc = parts_acc[part]
        thr = cfg["thresholds"][part]
        if not acc:
            global_results[part] = None
            continue
        wsum = sum(a["weight"] for a in acc)
        g_index = round(sum(a["index"] * a["weight"] for a in acc) / wsum, 1) if wsum else None
        tg = [a for a in acc if a["target"] is not None]
        g_target = round(sum(a["target"] * a["weight"] for a in tg) /
                         sum(a["weight"] for a in tg), 1) if tg else None
        taux_sum = sum(a["taux"] for a in acc)
        rate_low = round(sum(a["taux"] * a["low_frac"] for a in acc) / taux_sum, 5) if taux_sum else 0.0
        verdict, x = global_verdict(g_index or 0, rate_low, m_idx, thr, gpuiss)
        verdict_pred, _ = global_verdict(g_index or 0, rate_low, 3, thr, gpuiss)
        global_results[part] = {
            "index": g_index, "target_index": g_target,
            "normalized_x": x, "rate_low": rate_low,
            "verdict": verdict, "verdict_pred": verdict_pred,
            "full_scale": thr["taux_pleine_echelle"][m_idx],
        }
    sdv_results.sort(key=lambda r: r["order"])
    return event_updates, sdv_results, global_results
