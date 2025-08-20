
# === ADD to engine.py ===
def _edge_factor_from_ratio(ratio):
    """Convert ratio vs target to a gentle multiplier around 1.0.
    ratio <1 means worse than target when higher is better (or the inverse case handled by caller).
    Clamp to [0.7, 1.3] so it nudges but doesn't dominate.
    """
    try:
        r = float(ratio)
    except Exception:
        return 1.0
    from math import log
    # log curve around 1
    val = 1.0 + max(-0.4, min(0.4, log(r if r>0 else 1e-6)))
    return max(0.7, min(1.3, val))

def compute_edge_percentiles(templates, profile_key=None, measured=None, history=None):
    """Return edge multipliers per waste using benchmark targets and optional history.
    measured: dict like {'fpy_pct': 96, 'smed_changeover_min': 35, ...}
    history: dict of lists for percentiles, e.g., {'fpy_pct':[95,97,98]}
    """
    prio = templates.get("prioritization", {})
    metrics = prio.get("edge_metrics", {})
    prof = templates.get("profiles", {}).get(profile_key or "", {})
    bm = (prof.get("benchmarks", {}) if prof else {})
    edge = {}
    measured = measured or {}
    history = history or {}
    for waste, m in metrics.items():
        key = m.get("key")
        hib = bool(m.get("higher_is_better", True))
        target = bm.get(key)
        val = measured.get(key)
        if val is None or target is None:
            edge[waste] = 1.0
            continue
        # ratio vs target (>=1 good if hib; else <=1 good)
        ratio = (val/target) if hib else (target/max(val,1e-6))
        factor = _edge_factor_from_ratio(ratio)
        # optional: nudge with historical percentile (worse -> higher factor)
        hist = sorted([float(x) for x in history.get(key, []) if x is not None])
        if len(hist) >= 5:
            # simple percentile: position of val among history (for hib or reverse)
            import bisect
            if hib:
                p = bisect.bisect_left(hist, val)/len(hist)
                # low percentile (bad) -> factor up to +0.2
                factor *= (1.0 + max(0.0, 0.2*(0.5 - p)))
            else:
                # for lower-is-better metrics, invert
                p = 1.0 - (bisect.bisect_left(hist, val)/len(hist))
                factor *= (1.0 + max(0.0, 0.2*(0.5 - p)))
        edge[waste] = max(0.7, min(1.4, factor))
    return edge
