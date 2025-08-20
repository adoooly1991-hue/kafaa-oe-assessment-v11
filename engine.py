
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
    answers: Dict[str, Any] = None

def categorize_theme(waste: str) -> Tuple[str,str]:
    w = (waste or '').lower()
    if w in ("overproduction","overprocessing","motion","transportation"): return "P","Production"
    if w in ("defects",): return "Q","Quality"
    if w in ("inventory",): return "C","Cost"
    if w in ("waiting",): return "D","Delivery"
    if w in ("safety",): return "S","Safety"
    if w in ("talent",): return "M","Morale"
    return "P","Production"

def get_questionnaire_effects(step: ProcessStep, templates: Dict[str,Any], waste: str) -> Tuple[float, List[str]]:
    delta = 0.0
    snippets = []
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
    scores["waiting"] = min(5.0, (step.waiting_starved_pct or 0)/(th.get("waiting_pct_high",10.0)/3))
    scores["inventory"] = min(5.0, (step.wip_units_in or 0)/(th.get("inventory_wip_high",30.0)/3))
    scores["transportation"] = min(5.0, ((step.distance_m or 0) + (step.layout_moves or 0)*10) / (th.get("transport_distance_high_m",30.0)/3))
    scores["motion"] = 2.0 if (step.process_type or "Manual") == "Manual" else 0.5
    scores["overprocessing"] = min(5.0, (step.rework_pct or 0)/(th.get("rework_pct_high",2.0)/3))
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
    if waste == "defects": parts.append(f"Defect rate is {step.defect_pct:.1f}% and rework {step.rework_pct:.1f}%.")
    if waste == "waiting": parts.append(f"Waiting/starved time is {step.waiting_starved_pct:.1f}% of available time.")
    if waste == "inventory": parts.append(f"Input WIP stands at {step.wip_units_in:.0f} units.")
    if waste == "transportation": parts.append(f"Travel distance {step.distance_m:.0f} m with {step.layout_moves} hand-offs.")
    if waste == "overprocessing": parts.append(f"Rework observed at {step.rework_pct:.1f}%.")
    if waste == "overproduction": parts.append(f"Flow mode is {step.push_pull}. Consider CONWIP/Kanban.")
    if waste == "motion": parts.append(f"Process type: {step.process_type}. Ergonomic review recommended.")
    if waste == "safety": parts.append(f"Recorded safety incidents: {step.safety_incidents}.")
    dlt, snippets = get_questionnaire_effects(step, templates, waste)
    if snippets: parts.append("; ".join(snippets))
    obs = " ".join(parts)
    return {"step_id": step.id,"step_name": step.name,"waste": waste,"score_0_5": sc,"rpn_pct": rpn_pct,"confidence": confidence,"observation": obs}

def compute_lead_time(steps: List[ProcessStep], available_time_sec: float = 8*3600.0) -> Dict[str,Any]:
    lt = 0.0; by_step = {}; bottleneck=0.0
    for s in steps:
        ct_eff = max(0.0, (s.ct_sec or 0.0) * (1.0 + (s.waiting_starved_pct or 0.0)/100.0))
        by_step[s.id] = {"ct_eff_sec": ct_eff}
        lt += ct_eff
        bottleneck = max(bottleneck, ct_eff)
    return {"lead_time_sec": lt, "ct_bottleneck_sec": bottleneck, "by_step": by_step}

def build_material_flow_narrative(steps: List[ProcessStep], templates: Dict[str,Any], factory_name: str, year: str, cost_text: str, sales_text: str) -> str:
    seq = []
    for s in steps:
        verb = "transported" if s.distance_m>0 else "transferred"
        seq.append(f"{s.name} ({s.id}) with CT ~ {int(s.ct_sec)}s; materials {verb} ~{int(s.distance_m)}m; WIP {int(s.wip_units_in)}.")
    body = " ".join(seq)
    return f"Material Flow was observed and traced within {factory_name}. {body} All this handling costs ~{cost_text} in {year}, and lost sales opportunities ~{sales_text}."
