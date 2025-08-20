
from dataclasses import dataclass
from typing import Dict, Any, Tuple, List

@dataclass
class ProcessStep:
    id: str
    name: str
    ct_sec: float = 0.0
    wip_units_in: float = 0.0
    defect_pct: float = 0.0
    rework_pct: float = 0.0
    push_pull: str = "Push"
    process_type: str = "Manual"
    distance_m: float = 0.0
    layout_moves: int = 0
    waiting_starved_pct: float = 0.0
    safety_incidents: int = 0
    # New fields from the Excel model
    downtime_pct: float = 0.0
    changeover_freq: float = 0.0
    changeover_time_min: float = 0.0
    operators_n: float = 0.0
    touchpoints_n: float = 0.0
    answers: Dict[str, Any] = None

def categorize_theme(waste: str):
    w = (waste or '').lower()
    if w in ("overproduction","overprocessing","motion","transportation"): return "P","Production"
    if w in ("defects",): return "Q","Quality"
    if w in ("inventory",): return "C","Cost"
    if w in ("waiting",): return "D","Delivery"
    if w in ("safety",): return "S","Safety"
    if w in ("talent",): return "M","Morale"
    return "P","Production"

def get_questionnaire_effects(step: ProcessStep, templates: Dict[str,Any], waste: str) -> Tuple[float, List[str]]:
    delta = 0.0; snippets = []
    ans = (step.answers or {}).get(waste, {})
    if waste == "defects":
        tr = ans.get("trend")
        if tr == "Rising": delta += 1.0; snippets.append("Defect trend rising")
        elif tr == "Stable": delta += 0.3
        elif tr == "Falling": delta -= 0.2
    if waste == "waiting":
        freq = ans.get("frequency")
        if freq == "Frequent": delta += 1.0; snippets.append("Frequent starvation/blocks")
        elif freq == "Occasional": delta += 0.5
        elif freq == "Rare": delta += 0.1
    return delta, snippets

def score_wastes(step: ProcessStep, th: Dict[str,Any], templates: Dict[str,Any]) -> Dict[str,Any]:
    scores = {}
    scores["defects"] = min(5.0, (step.defect_pct or 0)/(th.get("defects_pct_high",3.0)/3)) if step.defect_pct else 0.0
    scores["waiting"] = min(5.0, ((step.waiting_starved_pct or 0) + (step.downtime_pct or 0)/2) / (th.get("waiting_pct_high",10.0)/3))
    scores["inventory"] = min(5.0, (step.wip_units_in or 0)/(th.get("inventory_wip_high",30.0)/3))
    scores["transportation"] = min(5.0, ((step.distance_m or 0) + (step.layout_moves or 0)*10 + (step.touchpoints_n or 0)*5) / (th.get("transport_distance_high_m",30.0)/3))
    scores["motion"] = min(5.0, (step.touchpoints_n or 0)/(th.get("touchpoints_high",6.0)/3)) + (1.0 if (step.process_type or 'Manual')=='Manual' else 0.3)
    scores["overprocessing"] = min(5.0, ((step.rework_pct or 0) + (step.changeover_time_min or 0)/th.get('changeover_time_high_min',30.0)) / (th.get("rework_pct_high",2.0)/3))
    scores["overproduction"] = 2.0 if (step.push_pull or "Push") == "Push" else 0.5
    scores["talent"] = 1.0 if ((step.answers or {}).get("talent")) else 0.5
    scores["safety"] = 5.0 if (step.safety_incidents or 0)>=(th.get("safety_incidents_high",1)) else (1.0 if (step.safety_incidents or 0)>0 else 0.2)
    deltas = {}
    for w in list(scores.keys()):
        dlt, _ = get_questionnaire_effects(step, templates, w)
        if dlt: scores[w] += dlt
        deltas[w] = dlt
    for k in scores: scores[k] = max(0.0, min(5.0, scores[k]))
    return {"scores": scores, "deltas": deltas}

def make_observation(step: ProcessStep, waste: str, waste_result: Dict[str,Any], templates: Dict[str,Any], th: Dict[str,Any]) -> Dict[str,Any]:
    sc = waste_result["scores"].get(waste, 0.0)
    if sc <= 0.0: return {}
    rpn_pct = min(100.0, sc/5.0*100.0)
    confidence = "High" if sc>=4.0 else ("Medium" if sc>=2.0 else "Low")
    parts = []
    parts.append(f"At {step.name} ({step.id}), {waste} was detected with score {sc:.1f}.")
    if waste == "defects": parts.append(f"Defect {step.defect_pct:.1f}%, rework {step.rework_pct:.1f}%.")
    if waste == "waiting": parts.append(f"Waiting/downtime ~{(step.waiting_starved_pct or 0.0)+(step.downtime_pct or 0.0):.1f}% of time.")
    if waste == "inventory": parts.append(f"WIP {step.wip_units_in:.0f} units.")
    if waste == "transportation": parts.append(f"Distance {step.distance_m:.0f} m, hand-offs {step.layout_moves}, touch-points {step.touchpoints_n:.0f}.")
    if waste == "motion": parts.append(f"Process: {step.process_type}, touch-points {step.touchpoints_n:.0f}.")
    if waste == "overprocessing": parts.append(f"Rework {step.rework_pct:.1f}%, changeover {step.changeover_time_min:.0f} min × {step.changeover_freq:.1f}/shift.")
    if waste == "overproduction": parts.append(f"Flow mode is {step.push_pull}; consider CONWIP/Kanban.")
    if waste == "safety": parts.append(f"Incidents: {step.safety_incidents}.")
    dlt, snippets = get_questionnaire_effects(step, templates, waste)
    if snippets: parts.append('; '.join(snippets))
    obs = ' '.join(parts)
    return {"step_id": step.id,"step_name": step.name,"waste": waste,"score_0_5": sc,"rpn_pct": rpn_pct,"confidence": confidence,"observation": obs}

def compute_lead_time(steps: List[ProcessStep], available_time_sec: float = 8*3600.0) -> Dict[str,Any]:
    lt = 0.0; by_step = {}; bottleneck=0.0
    for s in steps:
        availability = max(0.2, 1.0 - (s.downtime_pct or 0.0)/100.0)
        ct_eff = max(0.0, (s.ct_sec or 0.0) * (1.0 + (s.waiting_starved_pct or 0.0)/100.0)) / availability
        by_step[s.id] = {"ct_eff_sec": ct_eff}
        lt += ct_eff
        bottleneck = max(bottleneck, ct_eff)
    return {"lead_time_sec": lt, "ct_bottleneck_sec": bottleneck, "by_step": by_step}

def build_material_flow_narrative(steps: List[ProcessStep], templates: Dict[str,Any], factory_name: str, year: str, cost_text: str, sales_text: str) -> str:
    seq = []
    for s in steps:
        verb = "transported" if s.distance_m>0 else "transferred"
        seq.append(f"{s.name} ({s.id}) CT≈{int(s.ct_sec)}s; {verb} ~{int(s.distance_m)}m; WIP {int(s.wip_units_in)}; changeover {int(s.changeover_time_min)}m.")
    body = " ".join(seq)
    return f"Material Flow was observed within {factory_name}. {body} Handling and changeovers cost ~{cost_text} in {year}, with lost sales opportunities ~{sales_text}."



def score_vc_answers(vc_answers: dict, templates: dict, vc_confidence: dict=None, vc_followups: dict=None):
    """
    Returns per-stage ranked waste scores (0-5), issues, and a confidence index.
    vc_confidence: {stage:{qid: factor}} where factor in [0.4..1.0].
    vc_followups: {stage:{qid:{...}}} values from UI; added to issues.
    """
    out = {}
    qmap = (templates.get('value_chain',{}) or {}).get('questions',{})
    for stage, ans in (vc_answers or {}).items():
        qlist = qmap.get(stage, [])
        waste_scores = {}
        max_possible = 0.0
        issues = []
        conf_vals = []
        for q in qlist:
            qid = q.get('id')
            score = float(ans.get(qid, 0))
            ww = q.get('waste_weights', {}) or {}
            # confidence factor
            cf = 1.0
            if vc_confidence and stage in vc_confidence and qid in vc_confidence[stage]:
                cf = float(vc_confidence[stage][qid] or 1.0)
            conf_vals.append(cf)
            # accumulate
            local_max = max( (abs(v)*4.0 for v in ww.values()), default=0.0 )
            max_possible += local_max
            for w, wt in ww.items():
                waste_scores[w] = waste_scores.get(w, 0.0) + score * float(wt) * cf
            # capture high-severity issue line
            if score >= 3 and q.get('issue_if_high'):
                issues.append(q['issue_if_high'])
            # include followups values
            if vc_followups and stage in vc_followups and qid in vc_followups[stage]:
                fvals = vc_followups[stage][qid]
                if isinstance(fvals, dict):
                    for k,v in fvals.items():
                        if v not in (None, '', []):
                            issues.append(f"{q.get('text','')}: {k} = {v}")
        # normalize
        ranked = []
        denom = max(max_possible, 1e-6)
        for w, sc in waste_scores.items():
            ranked.append((w, max(0.0, min(5.0, 5.0*sc/denom))))
        ranked.sort(key=lambda x: x[1], reverse=True)
        conf_index = sum(conf_vals)/len(conf_vals) if conf_vals else 1.0
        out[stage] = {"ranked": ranked, "issues": issues, "confidence": conf_index}
    return out


def _get(d, path, default=None):
    cur = d or {}
    for p in path:
        if isinstance(cur, dict) and p in cur:
            cur = cur[p]
        else:
            return default
    return cur

def estimate_business_case(vc_summary, templates, vc_followups=None, assumptions=None):
    assumptions = (assumptions or {})
    labor_hr = float(assumptions.get("labor_cost_per_hour", 50.0))
    mat_cost = float(assumptions.get("material_cost_per_unit", 100.0))
    rework_min = float(assumptions.get("rework_time_min_per_unit", 10.0))
    fl_cost = float(assumptions.get("forklift_cost_per_hour", 120.0))
    finance_pct = float(assumptions.get("cost_of_capital_pct", 12.0))/100.0
    vol_month = float(assumptions.get("avg_monthly_volume_units", 10000.0))

    by_waste = {w:0.0 for w in ["defects","waiting","inventory","transportation","motion","overprocessing","overproduction","safety"]}

    def sev(sc): return max(0.0, min(1.0, sc/5.0))

    for row in (vc_summary or []):
        stage_name = row.get("stage_name","")
        sc_def = next((sc for (w,sc) in row.get("top3",[]) if w=="defects"), 0.0)
        if sc_def>0:
            f_unit = None
            if vc_followups and stage_name in vc_followups and "first_pass_yield" in vc_followups[stage_name]:
                f_unit = vc_followups[stage_name]["first_pass_yield"]
            cost_unit = float(assumptions.get("material_cost_per_unit", mat_cost))
            re_min = float(assumptions.get("rework_time_min_per_unit", rework_min))
            units = float(vol_month)
            if isinstance(f_unit, dict):
                cost_unit = float(f_unit.get("unit_material_cost") or cost_unit)
                re_min = float(f_unit.get("rework_time_min") or re_min)
                units = float(f_unit.get("monthly_volume_units") or units)
            defect_rate = 0.1 * sev(sc_def)
            annual_defect_units = units*12*defect_rate
            saving = annual_defect_units*(cost_unit + (re_min/60.0)*labor_hr) * 0.5
            by_waste["defects"] += saving

        sc_wait = next((sc for (w,sc) in row.get("top3",[]) if w=="waiting"), 0.0)
        f_chg = None
        if vc_followups and stage_name in vc_followups and "changeover_time" in vc_followups[stage_name]:
            f_chg = vc_followups[stage_name]["changeover_time"]
        ops = float((f_chg or {}).get("operators_n") or 0.0)
        chg_per_month = float((f_chg or {}).get("changeovers_per_month") or 0.0)
        if sc_wait>0 and ops>0 and chg_per_month>0:
            avoid_min = 30.0 * sev(sc_wait)
            saving = (avoid_min/60.0)*ops*chg_per_month*12* labor_hr
            by_waste["waiting"] += saving
            by_waste["overprocessing"] += 0.2*saving

        sc_inv = next((sc for (w,sc) in row.get("top3",[]) if w=="inventory"), 0.0)
        f_fg = None
        if vc_followups and stage_name in vc_followups and "aging_fg" in vc_followups[stage_name]:
            f_fg = vc_followups[stage_name]["aging_fg"]
        avg_fg = float((f_fg or {}).get("avg_fg_value") or 0.0)
        fin_rate = float((f_fg or {}).get("finance_rate_pct") or finance_pct*100.0)/100.0
        if sc_inv>0 and avg_fg>0:
            release = avg_fg * 0.2 * sev(sc_inv)
            saving = release * fin_rate
            by_waste["inventory"] += saving

        sc_tr = next((sc for (w,sc) in row.get("top3",[]) if w=="transportation"), 0.0)
        f_load = None
        if vc_followups and stage_name in vc_followups and "loading_time" in vc_followups[stage_name]:
            f_load = vc_followups[stage_name]["loading_time"]
        loads_per_day = float((f_load or {}).get("loads_per_day") or 0.0)
        fl = float((f_load or {}).get("forklift_cost_per_hour") or fl_cost)
        if sc_tr>0 and loads_per_day>0:
            avoid_min = 10.0 * sev(sc_tr)
            saving = (avoid_min/60.0) * loads_per_day * 300 * fl
            by_waste["transportation"] += saving

        sc_mo = next((sc for (w,sc) in row.get("top3",[]) if w=="motion"), 0.0)
        if sc_mo>0 and ops>0:
            saving = ops * labor_hr * 200 * 0.1 * sev(sc_mo)
            by_waste["motion"] += saving

        sc_op = next((sc for (w,sc) in row.get("top3",[]) if w=="overproduction"), 0.0)
        if sc_op>0 and avg_fg>0:
            by_waste["overproduction"] += avg_fg * 0.05 * sev(sc_op)

        sc_sa = next((sc for (w,sc) in row.get("top3",[]) if w=="safety"), 0.0)
        if sc_sa>0:
            by_waste["safety"] += 20000.0 * sev(sc_sa)

    total = sum(by_waste.values())
    return {"by_waste": by_waste, "notes": [], "total": total}
